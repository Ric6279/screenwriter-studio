// 编剧工坊 —— 宿主半（Host half）
//
// 三件事：
//   1. 注册 screenwriter_* 工具（让任何会话里的模型都能直接生成剧作蓝图 / 剧本）
//   2. 注册 screenwriter-studio 技能（把四本书的方法与九道工序交给模型）
//   3. 挂 /screenwriter/* 路由，给浏览器半的「九宫格工作台」提供状态与生成动作
//
// 真正的排版交给 engine/ 下的两个 Python 脚本：openpyxl 画表、python-docx 排版。
// 这一层只做参数翻译、串行执行、结果回传。

import { spawn } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { mkdir, readdir, readFile, stat, unlink, writeFile } from "node:fs/promises";
import { dirname, join, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));          // …/plugin/lib
const PLUGIN_DIR = resolve(HERE, "..");                        // …/plugin
const ENGINE_DIR = join(PLUGIN_DIR, "engine");
const STAGES_PATH = join(PLUGIN_DIR, "presets", "stages.json");
const SKILL_PATH = join(PLUGIN_DIR, "SKILL.md");

// 数据根：默认取插件目录的上一级 = 编剧工坊。
// 迁移到 mac / 其他目录、或插件被复制进 profile 时，用 SCREENWRITER_HOME 指回项目目录。
const DATA_ROOT = process.env.SCREENWRITER_HOME
  ? resolve(process.env.SCREENWRITER_HOME)
  : resolve(PLUGIN_DIR, "..");
const DEMO_DIR = join(DATA_ROOT, "demo");
const OUT_DIR = join(DATA_ROOT, "产出");
const KNOWLEDGE_DIR = join(DATA_ROOT, "知识库");
const UPLOAD_DIR = join(KNOWLEDGE_DIR, "上传");
const STATE_PATH = join(DATA_ROOT, "plugin-state.json");
const PROJECT_PATH = join(DEMO_DIR, "项目数据.json");
const SCRIPT_PATH = join(DEMO_DIR, "剧本_第01集.json");
const CURRENT_PROJECT_PATH = join(DEMO_DIR, "_当前项目.json");
const CURRENT_SCRIPT_PATH = join(DEMO_DIR, "_当前剧本.json");

const MAX_BODY_BYTES = 48 * 1024 * 1024;
const MAX_UPLOAD_BYTES = 30 * 1024 * 1024;
const KNOWLEDGE_EXTS = [".md", ".txt", ".docx", ".pdf", ".doc", ".rtf", ".epub"];
const PYTHON_CANDIDATES = [
  process.env.SCREENWRITER_PYTHON,
  process.platform === "win32"
    ? join(process.env.USERPROFILE ?? ".", ".dsh", "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "python.exe")
    : join(process.env.HOME ?? ".", ".dsh", "dsh-runtimes", "dsh-primary-runtime", "dependencies", "python", "bin", "python3"),
  process.platform === "win32" ? "python" : "python3",
  process.platform === "win32" ? "python3" : "python",
].filter(Boolean);

// OCR 依赖直接复用 course-kit 的 .pylibs（pymupdf + RapidOCR 的 ONNX 模型），
// 不往插件里再拷一份几百 MB 的库。库是按 Python 3.10 装的，所以 OCR/取字脚本
// 走单独的 Python；Excel/Word 仍用 DSH 自带 Python，互不影响。
const OCR_LIB_DIR = process.env.SCREENWRITER_OCR_LIBS
  ? resolve(process.env.SCREENWRITER_OCR_LIBS)
  : resolve(PLUGIN_DIR, "..", "..", "..", "course-kit", ".pylibs");
const OCR_PYTHON_CANDIDATES = [
  process.env.SCREENWRITER_OCR_PYTHON,
  process.platform === "win32"
    ? join(process.env.LOCALAPPDATA ?? "", "Programs", "Python", "Python310", "python.exe")
    : "python3.10",
].filter(Boolean);

const name = "screenwriter-studio";
const inject = ["tools", "skills", "webServer", "connection"];

// ------------------------------------------------------------------ 工具函数
function findPython() {
  for (const p of PYTHON_CANDIDATES) {
    if (p === "python" || p === "python3") return p;
    if (existsSync(p)) return p;
  }
  return "python";
}

function findOcrPython() {
  for (const p of OCR_PYTHON_CANDIDATES) {
    if (p === "python3.10") return p;
    if (existsSync(p)) return p;
  }
  return findPython();
}

function sendJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader("content-type", "application/json; charset=utf-8");
  res.setHeader("cache-control", "no-store");
  res.end(JSON.stringify(payload));
}

async function readBoundedBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.byteLength;
    if (size > MAX_BODY_BYTES) {
      req.resume();
      return null;
    }
    chunks.push(chunk);
  }
  return Buffer.concat(chunks, size).toString("utf8");
}

function cleanString(value, fallback) {
  return typeof value === "string" && value.trim() !== "" ? value.trim() : fallback;
}

async function readJson(path, fallback) {
  try {
    return JSON.parse(await readFile(path, "utf8"));
  } catch {
    return fallback;
  }
}

async function loadStages() {
  const def = await readJson(STAGES_PATH, { stages: [] });
  return Array.isArray(def.stages) ? def.stages : [];
}

async function loadProject() {
  return readJson(PROJECT_PATH, {
    project: {
      title: "未命名项目",
      subtitle: "",
      format: "12集 × 45分钟",
      genre: "",
      genre_cat: "",
      logline: "",
      theme: "",
      controlling_idea: "",
      audience: "",
      budget: "",
      created: "",
    },
  });
}

async function listFiles(dir, exts) {
  const out = [];
  try {
    for (const n of await readdir(dir)) {
      if (!exts.some((e) => n.toLowerCase().endsWith(e))) continue;
      const p = join(dir, n);
      const s = await stat(p);
      out.push({ name: n, path: p, size: s.size, mtime: s.mtimeMs });
    }
  } catch {
    /* 目录不存在时返回空 */
  }
  out.sort((a, b) => b.mtime - a.mtime);
  return out;
}

// ------------------------------------------------------------------ 项目卡
function emptyProject() {
  return {
    title: "", subtitle: "", format: "", genre: "", genre_cat: "",
    logline: "", theme: "", controlling_idea: "", audience: "", budget: "", tone: "",
  };
}

function projectFromData(data) {
  const p = data.project ?? {};
  return {
    title: p.title ?? "",
    subtitle: p.subtitle ?? "",
    format: p.format ?? "",
    genre: p.genre ?? "",
    genre_cat: p.genre_cat ?? "",
    logline: p.logline ?? "",
    theme: p.theme ?? "",
    controlling_idea: p.controlling_idea ?? "",
    audience: p.audience ?? "",
    budget: p.budget ?? "",
    tone: p.tone ?? "",
  };
}

// 情绪/类型词簇：只做“风格差异过大”的低成本检查，不调用模型。
const CLUSTERS = {
  "悬疑/犯罪": ["悬疑", "犯罪", "惊悚", "推理", "凶杀", "谋杀", "案件", "真相", "警", "侦探", "尸体", "绑架", "逃犯", "凶手"],
  "爱情/情感": ["爱情", "甜宠", "恋爱", "心动", "暗恋", "婚", "言情", "初恋", "前任", "CP"],
  "家庭/伦理": ["家庭", "亲情", "母女", "父子", "伦理", "代际", "故乡", "小镇", "姐妹", "兄弟", "母亲", "父亲"],
  "喜剧": ["喜剧", "搞笑", "幽默", "荒诞", "轻喜", "爆笑", "沙雕"],
  "科幻/灾难": ["科幻", "未来", "星际", "太空", "人工智能", "AI", "末日", "灾难", "怪兽", "外星", "飞船"],
  "古装/武侠": ["古装", "历史", "宫斗", "武侠", "江湖", "王朝", "仙侠", "朝堂", "侠"],
  "青春/成长": ["青春", "校园", "成长", "少年", "高中", "大学", "毕业", "高考", "初恋"],
};

function clusterScore(text) {
  const scores = {};
  for (const [name, words] of Object.entries(CLUSTERS)) {
    let n = 0;
    for (const w of words) {
      if (text.includes(w)) n += 1;
    }
    if (n > 0) scores[name] = n;
  }
  return scores;
}

function topCluster(scores) {
  const rows = Object.entries(scores).sort((a, b) => b[1] - a[1]);
  return rows[0] ?? [null, 0];
}

function checkProject(p) {
  const issues = [];
  const add = (field, message, severity = "warn") => issues.push({ field, message, severity });
  const text = [p.title, p.genre, p.genre_cat, p.logline, p.theme, p.tone].filter(Boolean).join(" ");

  if (!p.title?.trim()) add("title", "项目名还没填：总控与产物文件名都要用它。", "error");
  if (!p.format?.trim()) add("format", "格式没填：电影、剧集、短剧的节奏与页数完全不同。", "error");
  if (!p.genre?.trim()) add("genre", "类型没填：子智能体无法判断该用哪套类型合同。", "error");
  if (!p.audience?.trim()) add("audience", "受众没填：暴力、情感与幽默的尺度无法校准。", "error");
  if (!p.logline?.trim()) add("logline", "一句话梗概没填：九个阶段都缺少共同的故事核心。", "error");
  if (!p.theme?.trim()) add("theme", "主题没填：场景容易写成事件堆砌，而不是价值转变。", "warn");
  if (!p.tone?.trim()) add("tone", "基调没填：视觉与声音设计会各写各的。", "warn");
  if (!p.budget?.trim()) add("budget", "预算没填：类型与场景规模可能不匹配。", "warn");

  const loglineLen = (p.logline ?? "").trim().length;
  if (loglineLen > 0 && (loglineLen < 24 || loglineLen > 180)) {
    add("logline", `梗概目前 ${loglineLen} 字：建议 24—180 字，短了没钩子，长了没重点。`, "warn");
  }
  const themeLen = (p.theme ?? "").trim().length;
  if (themeLen > 80) add("theme", `主题 ${themeLen} 字：主题句最好一句话说清。`, "warn");

  const genreScores = clusterScore([p.genre, p.genre_cat].filter(Boolean).join(" "));
  const storyScores = clusterScore([p.logline, p.theme, p.tone].filter(Boolean).join(" "));
  const [gName, gScore] = topCluster(genreScores);
  const [sName, sScore] = topCluster(storyScores);
  if (gName && sName && gName !== sName && gScore >= 1 && sScore >= 1) {
    add("genre", `类型写的是「${gName}」，但梗概/主题/基调更像「${sName}」：请确认是不是同一部戏。`, "warn");
  }

  if (/儿童|少儿|亲子|低龄|全家欢/.test(p.audience ?? "") &&
      /犯罪|凶杀|尸体|惊悚|血腥|暴力|谋杀|自杀/.test(text)) {
    add("audience", "受众偏少儿/亲子，但故事里有犯罪、尸体或血腥元素：尺度可能冲突。", "error");
  }
  if (/低|小成本|中低成本|控制预算/.test(p.budget ?? "") &&
      /星际|太空|战争|灾难|爆炸|怪兽|特效|大场面|末日|飞船/.test(text)) {
    add("budget", "预算写的是低成本，但故事需要大场面/特效：制作上可能落不了地。", "error");
  }
  if (/轻喜|甜|温暖|治愈|合家欢/.test(p.tone ?? "") &&
      /黑暗|冷冽|压抑|血腥|绝望|惊悚|凶杀|尸体/.test([p.genre, p.logline].filter(Boolean).join(" "))) {
    add("tone", "基调写的是轻松/温暖，但类型与梗概偏黑暗/惊悚：观众会不知道自己在看什么。", "warn");
  }
  if (/冷冽|黑暗|压抑|沉重|克制/.test(p.tone ?? "") &&
      /甜宠|轻喜|搞笑|治愈|合家欢/.test([p.genre, p.logline].filter(Boolean).join(" "))) {
    add("tone", "基调写的是冷冽/沉重，但类型与梗概偏甜/轻喜：两种风格会互相抵消。", "warn");
  }

  let score = 100;
  for (const it of issues) score -= it.severity === "error" ? 14 : 7;
  score = Math.max(0, Math.min(100, score));
  // 只要还有任何提醒/缺口，就交给用户确认，不因总分高而静默通过。
  const hasError = issues.some((it) => it.severity === "error");
  const hasIssue = issues.length > 0;
  const level = hasError ? "bad" : hasIssue ? "warn" : "ok";
  return {
    score, level, issues,
    summary: level === "ok" ? "项目卡基本自洽，可以继续。"
      : level === "warn" ? "项目卡有几处风格/信息不一致，建议先改再继续。"
        : "项目卡缺口或冲突较多，直接继续会让九个阶段各写各的。",
    checkedAt: Date.now(),
  };
}

function buildBrief(project, stage) {
  const p = project;
  return [
    "【项目卡 · 总控同步】",
    `项目名：${p.title || "（未填）"}`,
    `格式：${p.format || "（未填）"}`,
    `类型：${p.genre || "（未填）"}${p.genre_cat ? "｜" + p.genre_cat : ""}`,
    `受众：${p.audience || "（未填）"}`,
    `预算：${p.budget || "（未填）"}`,
    `基调：${p.tone || "（未填）"}`,
    `一句话梗概：${p.logline || "（未填）"}`,
    `主题：${p.theme || "（未填）"}`,
    "",
    `【本阶段任务】${stage.summary}`,
    `【本阶段产出】${stage.output}`,
    "【要求】先复述你对项目卡的理解，再按本阶段默认预设开工；项目卡与本阶段预设冲突时，以项目卡为准。",
  ].join("\n");
}

async function listKnowledge() {
  const out = [];
  try {
    for (const n of await readdir(KNOWLEDGE_DIR)) {
      if (!n.toLowerCase().endsWith(".md")) continue;
      if (n.startsWith("_")) continue;
      const p = join(KNOWLEDGE_DIR, n);
      const s = await stat(p);
      out.push({ name: n, path: p, size: s.size, mtime: s.mtimeMs, kind: "method" });
    }
  } catch { /* 目录不存在 */ }
  try {
    for (const n of await readdir(UPLOAD_DIR)) {
      const ext = KNOWLEDGE_EXTS.find((e) => n.toLowerCase().endsWith(e));
      if (!ext) continue;
      const p = join(UPLOAD_DIR, n);
      const s = await stat(p);
      out.push({ name: n, path: p, size: s.size, mtime: s.mtimeMs, kind: n.endsWith(".摘录.md") ? "extract" : "upload" });
    }
  } catch { /* 还没有上传 */ }
  out.sort((a, b) => {
    if (a.kind !== b.kind) {
      const order = { method: 0, upload: 1, extract: 2 };
      return (order[a.kind] ?? 9) - (order[b.kind] ?? 9);
    }
    return b.mtime - a.mtime;
  });
  return out;
}


// ------------------------------------------------------------------ 演示种子
// 只有“载入示范案例”才生成这些摘要；空白项目的会话窗口保持为空。
function seedFor(stageId, data) {
  const p = data.project ?? {};
  const acts = data.acts ?? [];
  const beats = data.beats ?? [];
  const chars = data.characters ?? [];
  const scenes = data.scenes ?? [];
  const eps = data.episodes ?? [];
  const ep1 = data.ep1_scenes ?? [];

  const line = (s) => (s ?? "").toString();

  switch (stageId) {
    case "genre":
      return [
        `【${p.title ?? "项目"}】类型选择已完成。`,
        `主类型：${line(data.genre_analysis?.primary)}`,
        `副类型：${line(data.genre_analysis?.secondary)}`,
        `理由：${line(data.genre_analysis?.why_primary)}`,
        `观众钩子：${line(data.genre_analysis?.audience_hook)}`,
        `卖点：${(p.selling_points ?? []).slice(0, 3).map((x, i) => `${i + 1}. ${x}`).join("；")}`,
        "需要我换一个主类型、或补一份平台提案吗？",
      ].join("\n");
    case "outline":
      return [
        `【三幕八序列】共 ${acts.length} 个序列：`,
        ...acts.map((a) => `${a.seq}｜${a.name}｜${a.episodes}｜${a.timecode}\n  核心：${a.core_event}\n  转折：${a.turn}`),
        "告诉我哪一序列需要加厚或压缩，我按 25% / 50% / 75% 的时点重排。",
      ].join("\n\n");
    case "theme":
      return [
        `主题句：${line(p.theme)}`,
        `控制性理念：${line(p.controlling_idea)}`,
        `反理念：${line(p.counter_idea)}`,
        `主题问题：${line(p.theme_question)}`,
        "三个主题场：第1集“声音不会骗人，人会”（信念建立）；第6集“母亲当众否认”（信念失效）；第12集“摘下耳蜗听雨”（选择真相）。",
        "需要我把主题落到某一场的具体动作里吗？",
      ].join("\n");
    case "beats":
      return [
        `【十五节拍】已排到 540 分钟：`,
        ...beats.slice(0, 15).map((b) => `${String(b.no).padStart(2, "0")} ${b.beat}（${b.pct}，${b.episode}）：${b.event}`),
        "哪一拍最弱？我可以按“事件 + 价值转变 + 为什么有效”重写。",
      ].join("\n");
    case "plot":
      return [
        "【情节点】五颗钉子：",
        ...(data.plot_points ?? []).map((x) => `${x.no}. ${x.name}（${x.position}）：${x.content}`),
        "伏笔回馈已单独成表（05b）。要我把某颗钉子提前/延后吗？",
      ].join("\n");
    case "character":
      return [
        `【人物】${chars.length} 个主要人物：`,
        ...chars.map((c) => `· ${c.name}（${c.age}，${c.identity}）\n  欲望：${c.desire}\n  需求：${c.need}\n  秘密：${c.secret}`),
        "关系矩阵见 Excel「07 人物关系」。要我给某个人物补一场“救猫咪”时刻吗？",
      ].join("\n");
    case "setting":
      return [
        `【主要场景】${scenes.length} 个：`,
        ...scenes.map((s) => `· ${s.name}（${s.int_ext} / ${s.time}）：${s.function}`),
        "需要我为某一场写“空间如何阻止人物”的具体调度吗？",
      ].join("\n");
    case "sceneoutline":
      return [
        `【分集】12 集：`,
        ...eps.map((e) => `第${e.no}集《${e.title}》：${e.core}\n  钩子：${e.hook}`),
        `【第1集分场】${ep1.length} 场，合计约 ${ep1.reduce((n, s) => n + (parseInt(String(s.duration), 10) || 0), 0)} 分钟。`,
        "要我调整第 1 集的钩子或某几场的顺序吗？",
      ].join("\n");
    case "script":
      return [
        `【第1集完整剧本】已生成 ${(data.scenes ?? []).length} 场，格式按菲尔德规范。`,
        "场次：" + (data.scenes ?? []).map((s) => s.heading).join("｜"),
        "需要我加长某场对白、或把某场改成“晚进早出”吗？",
      ].join("\n");
    default:
      return "";
  }
}

function blankStage(s) {
  return {
    id: s.id,
    preset: s.preset,
    notes: "",
    session: [],
    brief: "",
    briefAt: 0,
    questions: [...(s.questions ?? [])],
    basis: [...(s.knowledge ?? [])],
    kbRefs: [],
  };
}

async function defaultState() {
  const def = await loadStages();
  return {
    version: 2,
    demoLoaded: false,
    project: emptyProject(),
    stages: def.map(blankStage),
    updatedAt: Date.now(),
  };
}

async function demoState() {
  const [def, data, script] = await Promise.all([
    loadStages(),
    loadProject(),
    readJson(SCRIPT_PATH, { scenes: [] }),
  ]);
  return {
    version: 2,
    demoLoaded: true,
    project: projectFromData(data),
    stages: def.map((s) => ({
      ...blankStage(s),
      session: [{
        role: "assistant",
        time: Date.now(),
        seed: true,
        text: seedFor(s.id, s.id === "script" ? script : data),
      }],
    })),
    updatedAt: Date.now(),
  };
}

let statePromise = null;

async function readState() {
  if (statePromise !== null) return statePromise;
  statePromise = (async () => {
    const base = await defaultState();
    const saved = await readJson(STATE_PATH, null);
    if (!saved) return base;
    const byId = new Map(base.stages.map((s) => [s.id, s]));
    for (const s of saved.stages ?? []) {
      const cur = byId.get(s.id);
      if (!cur) continue;
      cur.preset = cleanString(s.preset, cur.preset);
      cur.notes = typeof s.notes === "string" ? s.notes : "";
      cur.session = Array.isArray(s.session) ? s.session : cur.session;
      cur.brief = typeof s.brief === "string" ? s.brief : "";
      cur.briefAt = typeof s.briefAt === "number" ? s.briefAt : 0;
      cur.questions = Array.isArray(s.questions) ? s.questions.filter((x) => typeof x === "string") : cur.questions;
      cur.basis = Array.isArray(s.basis) ? s.basis.filter((x) => typeof x === "string") : cur.basis;
      cur.kbRefs = Array.isArray(s.kbRefs) ? s.kbRefs.filter((x) => typeof x === "string") : cur.kbRefs;
    }
    const savedProject = saved.project ?? {};
    return {
      ...base,
      demoLoaded: typeof saved.demoLoaded === "boolean" ? saved.demoLoaded : Boolean(savedProject.title),
      project: { ...base.project, ...savedProject },
      stages: [...byId.values()],
      updatedAt: saved.updatedAt ?? Date.now(),
    };
  })();
  return statePromise;
}

async function saveState(state) {
  statePromise = Promise.resolve(state);
  await mkdir(DATA_ROOT, { recursive: true });
  await writeFile(STATE_PATH, JSON.stringify(state, null, 2) + "\n", "utf8");
}

// ------------------------------------------------------------------ Python
let runQueue = Promise.resolve();

function runPython(script, args, opts = {}) {
  const task = runQueue.then(
    () =>
      new Promise((resolveRun) => {
        const file = join(ENGINE_DIR, script);
        if (!existsSync(file)) {
          resolveRun({ ok: false, log: `找不到引擎脚本：${file}` });
          return;
        }
        const pythonPath = existsSync(OCR_LIB_DIR)
          ? `${OCR_LIB_DIR}${process.platform === "win32" ? ";" : ":"}${process.env.PYTHONPATH ?? ""}`
          : process.env.PYTHONPATH ?? "";
        const env = {
          ...process.env,
          PYTHONUTF8: "1",
          PYTHONIOENCODING: "utf-8",
          PYTHONPATH: pythonPath,
          ...(opts.ocr ? { SCREENWRITER_OCR_LIBS: OCR_LIB_DIR } : {}),
        };
        const python = opts.ocr ? findOcrPython() : findPython();
        const child = spawn(python, ["-u", "-W", "ignore", file, ...args], {
          env,
          windowsHide: true,
          stdio: ["ignore", "pipe", "pipe"],
        });
        let out = "";
        let err = "";
        child.stdout.on("data", (c) => { out += c.toString("utf8"); });
        child.stderr.on("data", (c) => { err += c.toString("utf8"); });
        child.on("error", (e) => resolveRun({ ok: false, log: String(e?.message ?? e) }));
        child.on("close", (code) => {
          const tail = err.trim().split("\n").slice(-10).join("\n");
          resolveRun({ ok: code === 0, code, log: (out + (tail ? `\n[stderr]\n${tail}` : "")).trim() });
        });
      })
  );
  runQueue = task.catch(() => undefined);
  return task;
}

async function materialize(state) {
  const data = await loadProject();
  data.project = { ...data.project, ...state.project };
  await mkdir(DEMO_DIR, { recursive: true });
  await writeFile(CURRENT_PROJECT_PATH, JSON.stringify(data, null, 2) + "\n", "utf8");

  const script = await readJson(SCRIPT_PATH, null);
  if (script) {
    script.title = state.project.title || script.title;
    await writeFile(CURRENT_SCRIPT_PATH, JSON.stringify(script, null, 2) + "\n", "utf8");
  } else {
    await writeFile(CURRENT_SCRIPT_PATH, JSON.stringify({ title: state.project.title, scenes: [] }, null, 2) + "\n", "utf8");
  }
  return { project: CURRENT_PROJECT_PATH, script: CURRENT_SCRIPT_PATH };
}

async function buildTargets(target) {
  const state = await readState();
  if (!state.demoLoaded) {
    const msg = "当前是空白项目：九阶段还没有结构化内容，无法直接出 Excel/Word。先「载入示范案例」，或在各阶段完成内容后导入项目 JSON。";
    const kinds = target === "xlsx" ? ["xlsx"] : target === "docx" ? ["docx"] : ["xlsx", "docx"];
    return kinds.map((kind) => ({ kind, ok: false, log: msg, out: null }));
  }
  const paths = await materialize(state);
  const safeTitle = (state.project.title || "未命名项目").replace(/[\\/:*?"<>|\s]+/g, "-");
  const results = [];

  if (target === "xlsx" || target === "all") {
    const out = join(OUT_DIR, `01_剧作蓝图_${safeTitle}.xlsx`);
    const r = await runPython("build_xlsx.py", ["--project", paths.project, "--out", out]);
    results.push({ kind: "xlsx", ok: r.ok, log: r.log, out });
  }
  if (target === "docx" || target === "all") {
    const out = join(OUT_DIR, `02_第01集_完整剧本_${safeTitle}.docx`);
    const r = await runPython("build_docx.py", ["--script", paths.script, "--out", out]);
    results.push({ kind: "docx", ok: r.ok, log: r.log, out });
  }
  return results;
}

// ------------------------------------------------------------------ 知识库上传
function lastJsonLine(log) {
  const lines = String(log ?? "").split("\n").filter((l) => l.trim().startsWith("{"));
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    try { return JSON.parse(lines[i]); } catch { /* 继续 */ }
  }
  return null;
}

function safeUploadName(name) {
  const base = String(name ?? "").replace(/[\\/:*?"<>|]/g, "-").replace(/^\.+/, "").slice(0, 180);
  return base || "上传文件";
}

async function ingestKnowledge(target) {
  const lower = target.toLowerCase();
  if (![".docx", ".pdf", ".txt", ".md"].some((e) => lower.endsWith(e))) return null;
  const r = await runPython("ingest_knowledge.py", ["--src", target, "--outDir", UPLOAD_DIR], { ocr: true });
  return { ok: r.ok, result: lastJsonLine(r.log), log: r.ok ? undefined : r.log };
}



// ------------------------------------------------------------------ 总控同步
async function syncProjectCard(body = {}) {
  const prev = await readState();
  const project = { ...prev.project };
  if (body?.project && typeof body.project === "object") {
    for (const key of ["title", "subtitle", "format", "genre", "genre_cat", "logline", "theme", "controlling_idea", "audience", "budget", "tone"]) {
      if (typeof body.project[key] === "string") project[key] = body.project[key].trim();
    }
  }
  const check = checkProject(project);
  if (body?.force !== true && check.level !== "ok") {
    return { status: 409, payload: { ok: false, code: "PROJECT_CHECK", check, project } };
  }
  const def = await loadStages();
  const byId = new Map(def.map((s) => [s.id, s]));
  const now = Date.now();
  const stages = prev.stages.map((cur) => {
    const meta = byId.get(cur.id);
    if (!meta) return cur;
    return {
      ...cur,
      brief: buildBrief(project, meta),
      briefAt: now,
      session: [
        ...(cur.session ?? []),
        { role: "system", time: now, text: "总控已同步项目卡。请按上方「总控简报」开始本阶段。" },
      ].slice(-200),
    };
  });
  const next = { ...prev, project, stages, updatedAt: now };
  await saveState(next);
  return {
    status: 200,
    payload: {
      ok: true,
      check,
      project,
      synced: stages.length,
      stages: stages.map((s) => {
        const meta = byId.get(s.id) ?? {};
        return { ...meta, ...s };
      }),
    },
  };
}



const OUT_SCHEMA = { type: "object", additionalProperties: true };

const DEMO_TOOL = {
  name: "screenwriter_demo",
  description: "把《回声镇》示范案例载入工作台：填充项目卡、九阶段会话与示范产物数据；随后可用 screenwriter_build_all 生成 Excel/Word。",
  parameters: { type: "object", properties: {}, additionalProperties: false },
  execute: async () => {
    const demo = await demoState();
    await saveState(demo);
    return { isError: false, value: { ok: true, demoLoaded: true, project: demo.project, stages: demo.stages.length } };
  },
};

const NEW_PROJECT_TOOL = {
  name: "screenwriter_new_project",
  description: "把编剧工坊重置为空白项目（清空项目卡、九阶段会话与总控简报），不会删除产出文件。",
  parameters: { type: "object", properties: {}, additionalProperties: false },
  execute: async () => {
    const blank = await defaultState();
    await saveState(blank);
    return { isError: false, value: { ok: true, demoLoaded: false, project: blank.project } };
  },
};


function textRender(_args, value) {
  return [{ type: "text", text: JSON.stringify(value, null, 1) }];
}

function define(registry, def) {
  return registry.register({ ...def, output: { schema: OUT_SCHEMA, render: textRender } });
}

const STAGES_TOOL = {
  name: "screenwriter_stages",
  description:
    "列出编剧工坊的九个阶段（类型选择、三幕八序列、主题、详细小节、情节点、人物、场景、分场分集、完整剧本），" +
    "每阶段带默认预设、主责方法、产出格式与自检问题。写任何阶段前先调它，按该阶段预设工作。",
  parameters: { type: "object", properties: {}, additionalProperties: false },
  execute: async () => {
    const stages = await loadStages();
    return {
      isError: false,
      value: stages.map((s) => ({
        id: s.id, no: s.no, phase: s.phase, name: s.name, role: s.role,
        output: s.output, summary: s.summary, preset: s.preset, questions: s.questions,
      })),
    };
  },
};

const BUILD_XLSX_TOOL = {
  name: "screenwriter_build_xlsx",
  description:
    "把项目数据 JSON 生成《剧作蓝图》Excel（12 张表：封面、类型、梗概、三幕八序列、十五节拍、情节点、伏笔、人物、关系、场景、分集、分场、自检、情绪曲线）。" +
    "默认读取编剧工坊/demo/项目数据.json；out 不填则写入 产出/01_剧作蓝图_<标题>.xlsx。",
  parameters: {
    type: "object",
    properties: {
      project: { type: "string", description: "项目数据 JSON 路径（默认 demo/项目数据.json）" },
      out: { type: "string", description: "输出 xlsx 路径" },
    },
    additionalProperties: false,
  },
  execute: async (args) => {
    const source = args.project || PROJECT_PATH;
    const data = await readJson(source, null);
    if (!data) return { isError: true, value: { ok: false, message: `读不到项目数据：${source}` } };
    const safe = (data.project?.title || "未命名项目").replace(/[\\/:*?"<>|\s]+/g, "-");
    const out = args.out || join(OUT_DIR, `01_剧作蓝图_${safe}.xlsx`);
    const r = await runPython("build_xlsx.py", ["--project", source, "--out", out]);
    const files = await listFiles(OUT_DIR, [".xlsx"]);
    return { isError: !r.ok, value: { ok: r.ok, out, log: r.log, outputs: files } };
  },
};

const BUILD_DOCX_TOOL = {
  name: "screenwriter_build_docx",
  description:
    "把剧本 JSON 生成 Word 完整剧本（封面、人物表、场次标题、动作、人物、对白、转场、页码）。" +
    "默认读取编剧工坊/demo/剧本_第01集.json；out 不填则写入 产出/02_第01集_完整剧本_<标题>.docx。",
  parameters: {
    type: "object",
    properties: {
      script: { type: "string", description: "剧本 JSON 路径（默认 demo/剧本_第01集.json）" },
      out: { type: "string", description: "输出 docx 路径" },
    },
    additionalProperties: false,
  },
  execute: async (args) => {
    const source = args.script || SCRIPT_PATH;
    const data = await readJson(source, null);
    if (!data) return { isError: true, value: { ok: false, message: `读不到剧本数据：${source}` } };
    const safe = (data.title || "未命名项目").replace(/[\\/:*?"<>|\s]+/g, "-");
    const out = args.out || join(OUT_DIR, `02_第01集_完整剧本_${safe}.docx`);
    const r = await runPython("build_docx.py", ["--script", source, "--out", out]);
    const files = await listFiles(OUT_DIR, [".docx"]);
    return { isError: !r.ok, value: { ok: r.ok, out, log: r.log, outputs: files } };
  },
};

const BUILD_ALL_TOOL = {
  name: "screenwriter_build_all",
  description:
    "一键重建当前项目的 Excel 剧作蓝图与 Word 剧本，使用工作台里已保存的项目信息（标题/类型/梗概/主题）。",
  parameters: { type: "object", properties: {}, additionalProperties: false },
  execute: async () => {
    const results = await buildTargets("all");
    return {
      isError: results.some((r) => !r.ok),
      value: { ok: results.every((r) => r.ok), results, outputs: await listFiles(OUT_DIR, [".xlsx", ".docx"]) },
    };
  },
};

// ------------------------------------------------------------------ 技能
function skillContent() {
  try {
    return readFileSync(SKILL_PATH, "utf8");
  } catch {
    return "编剧工坊技能文档缺失。";
  }
}

// ------------------------------------------------------------------ 插件主体
function apply(ctx) {
  const definitions = [STAGES_TOOL, BUILD_XLSX_TOOL, BUILD_DOCX_TOOL, BUILD_ALL_TOOL, DEMO_TOOL, NEW_PROJECT_TOOL];
  for (const definition of definitions) {
    ctx.effect(() => define(ctx.tools, definition), `screenwriter: ${definition.name}`);
  }

  ctx.effect(
    () =>
      ctx.skills.register({
        name: "screenwriter-studio",
        description:
          "编剧工坊：以《故事》《救猫咪》《电影剧本写作基础》《打草稿》四本书为方法，按九个阶段（类型、三幕八序列、主题、小节、情节点、人物、场景、分场分集、完整剧本）产出 Excel 剧作蓝图与 Word 剧本。",
        whenToUse:
          "当用户要求写剧本、写故事大纲、设计人物、搭三幕/八序列/十五节拍、做分集分场、或把剧作资料整理成 Excel/Word 时使用。",
        source: "runtime",
        invocation: { modelInvocable: true, userInvocable: true },
        content: skillContent(),
        metadata: { engine: ENGINE_DIR, python: findPython(), root: DATA_ROOT },
      }),
    "screenwriter: skill"
  );

  const connectionOf = () => Reflect.get(ctx, "connection");
  const rejected = (req, res) => {
    const rejection = connectionOf().requestRejection(req);
    if (rejection === void 0) return false;
    res.statusCode = rejection;
    res.end();
    return true;
  };

  // GET/POST /screenwriter/state
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/state",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method === "GET") {
        const state = await readState();
        const def = await loadStages();
        sendJson(res, 200, {
          ok: true,
          root: DATA_ROOT,
          outDir: OUT_DIR,
          knowledgeDir: KNOWLEDGE_DIR,
          uploadDir: UPLOAD_DIR,
          python: findPython(),
          demoLoaded: state.demoLoaded,
          project: state.project,
          stages: def.map((s) => {
            const saved = state.stages.find((x) => x.id === s.id) ?? blankStage(s);
            return {
              ...s,
              preset: saved.preset,
              session: saved.session,
              notes: saved.notes,
              brief: saved.brief,
              briefAt: saved.briefAt,
              questions: saved.questions,
              basis: saved.basis,
              kbRefs: saved.kbRefs,
            };
          }),
          defaults: Object.fromEntries(def.map((s) => [s.id, s.preset])),
          outputs: await listFiles(OUT_DIR, [".xlsx", ".docx"]),
          knowledge: await listKnowledge(),
        });
        return;
      }
      if (req.method === "POST") {
        const text = await readBoundedBody(req);
        if (text === null) { sendJson(res, 413, { ok: false, message: "请求体过大" }); return; }
        let body;
        try { body = JSON.parse(text); } catch { sendJson(res, 400, { ok: false, message: "请求体必须是 JSON" }); return; }
        const prev = await readState();
        const next = { ...prev, project: { ...prev.project }, stages: prev.stages.map((s) => ({ ...s })) };
        if (body && typeof body.project === "object" && body.project !== null) {
          for (const key of ["title", "subtitle", "format", "genre", "genre_cat", "logline", "theme", "controlling_idea", "audience", "budget", "tone"]) {
            if (typeof body.project[key] === "string") next.project[key] = body.project[key].trim();
          }
        }
        if (Array.isArray(body?.stages)) {
          for (const patch of body.stages) {
            const cur = next.stages.find((s) => s.id === patch?.id);
            if (!cur) continue;
            if (typeof patch.preset === "string" && patch.preset.trim() !== "") cur.preset = patch.preset;
            if (Array.isArray(patch.session)) cur.session = patch.session.slice(-200);
            if (typeof patch.notes === "string") cur.notes = patch.notes;
            if (typeof patch.brief === "string") cur.brief = patch.brief;
            if (typeof patch.briefAt === "number") cur.briefAt = patch.briefAt;
            if (Array.isArray(patch.questions)) cur.questions = patch.questions.filter((x) => typeof x === "string");
            if (Array.isArray(patch.basis)) cur.basis = patch.basis.filter((x) => typeof x === "string");
            if (Array.isArray(patch.kbRefs)) cur.kbRefs = patch.kbRefs.filter((x) => typeof x === "string");
          }
        }
        next.updatedAt = Date.now();
        await saveState(next);
        sendJson(res, 200, { ok: true, project: next.project, demoLoaded: next.demoLoaded, outputs: await listFiles(OUT_DIR, [".xlsx", ".docx"]) });
        return;
      }
      res.statusCode = 405;
      res.setHeader("allow", "GET, POST");
      res.end();
    },
  }), "screenwriter: /screenwriter/state");

  // POST /screenwriter/demo —— 载入示范案例（项目卡 + 九阶段种子会话）
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/demo",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const demo = await demoState();
      await saveState(demo);
      sendJson(res, 200, { ok: true, demoLoaded: true, project: demo.project });
    },
  }), "screenwriter: POST /screenwriter/demo");

  // POST /screenwriter/project/new —— 回到空白项目
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/project/new",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const blank = await defaultState();
      await saveState(blank);
      sendJson(res, 200, { ok: true, demoLoaded: false, project: blank.project });
    },
  }), "screenwriter: POST /screenwriter/project/new");

  // POST /screenwriter/project/check —— 项目卡“最低花费”自检
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/project/check",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const state = await readState();
      const text = await readBoundedBody(req);
      let project = state.project;
      if (text !== null && text.trim() !== "") {
        try {
          const body = JSON.parse(text);
          if (body?.project && typeof body.project === "object") project = { ...project, ...body.project };
        } catch { /* 用已保存项目 */ }
      }
      sendJson(res, 200, { ok: true, check: checkProject(project) });
    },
  }), "screenwriter: POST /screenwriter/project/check");

  // POST /screenwriter/project/sync —— 总控：校验并同步项目卡到九阶段
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/project/sync",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const text = await readBoundedBody(req);
      let body = {};
      if (text !== null && text.trim() !== "") {
        try { body = JSON.parse(text); } catch { sendJson(res, 400, { ok: false, message: "请求体必须是 JSON" }); return; }
      }
      const result = await syncProjectCard(body);
      sendJson(res, result.status, result.payload);
    },
  }), "screenwriter: POST /screenwriter/project/sync");

  // POST /screenwriter/knowledge/upload —— 上传剧本/编剧书，落盘并抽文本
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/knowledge/upload",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const text = await readBoundedBody(req);
      if (text === null) { sendJson(res, 413, { ok: false, message: "文件太大（上限 30 MB）" }); return; }
      let body;
      try { body = JSON.parse(text); } catch { sendJson(res, 400, { ok: false, message: "请求体必须是 JSON" }); return; }
      const rawName = safeUploadName(body?.name);
      const ext = KNOWLEDGE_EXTS.find((e) => rawName.toLowerCase().endsWith(e));
      if (!ext) {
        sendJson(res, 400, { ok: false, message: `只支持 ${KNOWLEDGE_EXTS.join(" / ")}` });
        return;
      }
      const b64 = String(body?.data ?? "").replace(/^data:[^;]+;base64,/, "");
      let buf;
      try { buf = Buffer.from(b64, "base64"); } catch { sendJson(res, 400, { ok: false, message: "base64 解码失败" }); return; }
      if (buf.length === 0) { sendJson(res, 400, { ok: false, message: "空文件" }); return; }
      if (buf.length > MAX_UPLOAD_BYTES) { sendJson(res, 413, { ok: false, message: "文件超过 30 MB" }); return; }
      await mkdir(UPLOAD_DIR, { recursive: true });
      let name = rawName;
      if (existsSync(join(UPLOAD_DIR, name))) {
        const stem = name.slice(0, name.length - ext.length);
        name = `${stem}_${Date.now()}${ext}`;
      }
      const target = join(UPLOAD_DIR, name);
      await writeFile(target, buf);
      const ingest = await ingestKnowledge(target);
      sendJson(res, 200, {
        ok: true,
        file: { name, path: target, size: buf.length, kind: "upload" },
        ingest,
        knowledge: await listKnowledge(),
      });
    },
  }), "screenwriter: POST /screenwriter/knowledge/upload");

  // POST /screenwriter/knowledge/delete —— 只允许删上传目录里的文件
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/knowledge/delete",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      const text = await readBoundedBody(req);
      let body;
      try { body = JSON.parse(text); } catch { sendJson(res, 400, { ok: false, message: "请求体必须是 JSON" }); return; }
      const target = resolve(String(body?.path ?? ""));
      if (!target.startsWith(UPLOAD_DIR + sep)) {
        sendJson(res, 403, { ok: false, message: "只能删除上传目录里的文件" });
        return;
      }
      try {
        await unlink(target);
      } catch (e) {
        sendJson(res, 404, { ok: false, message: String(e?.message ?? e) });
        return;
      }
      sendJson(res, 200, { ok: true, knowledge: await listKnowledge() });
    },
  }), "screenwriter: POST /screenwriter/knowledge/delete");

  // POST /screenwriter/build
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/build",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "POST") {
        res.statusCode = 405;
        res.setHeader("allow", "POST");
        res.end();
        return;
      }
      let target = "all";
      const text = await readBoundedBody(req);
      if (text !== null && text.trim() !== "") {
        try {
          const body = JSON.parse(text);
          if (body?.target === "xlsx" || body?.target === "docx") target = body.target;
        } catch { /* 忽略，按 all 处理 */ }
      }
      const results = await buildTargets(target);
      sendJson(res, 200, {
        ok: results.every((r) => r.ok),
        results,
        outputs: await listFiles(OUT_DIR, [".xlsx", ".docx"]),
        outDir: OUT_DIR,
      });
    },
  }), "screenwriter: POST /screenwriter/build");

  // GET /screenwriter/files
  ctx.effect(() => ctx.webServer.register({
    kind: "exact",
    path: "/screenwriter/files",
    handler: async (req, res) => {
      if (rejected(req, res)) return;
      if (req.method !== "GET") {
        res.statusCode = 405;
        res.setHeader("allow", "GET");
        res.end();
        return;
      }
      sendJson(res, 200, {
        ok: true,
        outDir: OUT_DIR,
        uploadDir: UPLOAD_DIR,
        outputs: await listFiles(OUT_DIR, [".xlsx", ".docx"]),
        knowledge: await listKnowledge(),
      });
    },
  }), "screenwriter: /screenwriter/files");
}

export { apply, inject, name };
