# 编剧工坊 · 工作总结 / 交接与 macOS 迁移说明

> **2026-10-26 更新（三栏布局 + v2 功能 + OCR）**：初始空白，demo 改为可载入的
> 示范案例；总控/自检、可编辑阶段说明、知识库上传全部接回；扫描版 PDF 自动走本地
> RapidOCR。详见第 12 节。
>
> 给接手的会话看：这份文档是自包含的，读完可以直接开工，不需要上一轮对话。
> 写于 2026-10-19。项目在 Windows 上已跑通；本文重点是**搬到 macOS 后如何快速接手、验证、调整**。
> 先读本文件，再读 `README.md` 与 `plugin/README.md`；不要从零重写。

---

## 0. 三十秒交接

这是一个 DSH 插件 + 一套演示数据 + 一份四书知识库：

- 插件包：`@fatfish/screenwriter-studio`，侧栏名「✎ 编剧工坊」。
- 九个阶段（类型、三幕八序列、主题、详细小节、情节点、人物、场景、分场分集、完整剧本），
  每阶段一个可编辑的**默认预设**与一个**会话窗口**。
- 产物：`产出/01_剧作蓝图_回声镇.xlsx`（14 张可见表 + 1 张隐藏曲线数据表）、
  `产出/02_第01集_完整剧本_回声镇.docx`（13 页）。
- 演示项目：《回声镇》，12 集社会派悬疑短剧；第 1 集《回声》已写成完整剧本。

**接手后的第一件事不是改代码，而是按第 5 节在 Mac 上跑一遍 `screenwriter_build_all`。**
跑通了再谈界面与体验。

---

## 1. 当前状态

| 项 | 状态 | 证据 |
|---|---|---|
| 插件安装 | Windows profile 上已安装并 applied | `plugin_manager` 返回 `application: applied` |
| 宿主工具 | 4 个工具可用 | `screenwriter_stages` 返回 9 阶段；`screenwriter_build_all` 返回 ok |
| 客户端槽位 | 已注册 | `sidebar.panellist` 有 `screenwriter`；`main` 有 `screenwriter` |
| Excel | 已生成并渲染检查 | 14 张可见表；封面、三幕、人物、分集曲线、分场均检查过 |
| Word | 已生成并渲染检查 | 13 页；封面、人物表、正文、尾页均检查过 |
| 界面点击流 | **未做截图级验收** | 本机没有可用的浏览器控制通道 |
| macOS | **未测** | 本文第 6 节是迁移清单，不是已验证结论 |

### demo 故事一句话

声音设计师林晚回到故乡录“小镇声音档案”，在废品站捡到妹妹十年前的录音带；
磁带里有“关闸”和一个少年说“爸，我按了”。她越查越发现，那晚她也在场，
听见妹妹呼救却删掉了录音；供她做耳蜗手术、供她离开小镇的钱，正是母亲收下的封口费。

---

## 2. 目录与职责

```
screenwriter-tools/编剧工坊/
├─ README.md                 总说明（给人看）
├─ HANDOFF.md                本文件（给接手会话看）
├─ 知识库/
│  ├─ 00_四书整合_创作流程.md   九阶段分工、冲突裁决、流水线
│  ├─ 01_故事_麦基.md          控制性理念/激励事件/鸿沟/危机
│  ├─ 02_救猫咪_斯奈德.md      十类型/logline/15 节拍
│  ├─ 03_电影剧本写作基础_菲尔德.md  25/50/25/情节点/格式
│  ├─ 04_打草稿_编剧思维训练表.md    概念/人物/关系/情节/分场表
│  └─ _ocr_打草稿/            原书 p9—192 的 OCR 原文（184 页缓存）
├─ demo/
│  ├─ 项目数据.json            14 张表的结构化数据（31 KB）
│  └─ 剧本_第01集.json          14 场剧本正文（19 KB）
├─ plugin/
│  ├─ package.json            @fatfish/screenwriter-studio
│  ├─ cordis.patch.yml        插入一行 screenwriter
│  ├─ icon.svg / SKILL.md / README.md
│  ├─ presets/stages.json     九阶段默认预设（16.6 KB）
│  ├─ lib/host.js             宿主半（23 KB）：工具/技能/路由/状态
│  ├─ lib/client.js           浏览器半（30 KB）：九宫格工作台
│  └─ engine/
│     ├─ build_xlsx.py        项目 JSON → 剧作蓝图 xlsx
│     └─ build_docx.py        剧本 JSON → Word 剧本
└─ 产出/
   ├─ 01_剧作蓝图_回声镇.xlsx
   └─ 02_第01集_完整剧本_回声镇.docx
```

### 关键代码位置（行号以 2026-10-19 版本为准）

| 位置 | 作用 |
|---|---|
| `plugin/lib/host.js:25` | `DATA_ROOT`，支持 `SCREENWRITER_HOME` 覆盖 |
| `plugin/lib/host.js:38` | `PYTHON_CANDIDATES`，已按 win32/darwin 分支 |
| `plugin/lib/host.js:359` | `screenwriter_stages` 工具 |
| `plugin/lib/host.js:378` | `screenwriter_build_xlsx` |
| `plugin/lib/host.js:403` | `screenwriter_build_docx` |
| `plugin/lib/host.js:428` | `screenwriter_build_all` |
| `plugin/lib/host.js:459` | 技能注册 `screenwriter-studio` |
| `plugin/lib/host.js:485` | `GET/POST /screenwriter/state` |
| `plugin/lib/host.js:543` | `POST /screenwriter/build` |
| `plugin/lib/host.js:573` | `GET /screenwriter/files` |
| `plugin/lib/client.js:20` | 三个 API 相对路径 |
| `plugin/lib/client.js:673` | 侧栏槽位 `sidebar.panellist` |
| `plugin/lib/client.js:679` | 主面板槽位 `main` |
| `plugin/engine/build_xlsx.py:39` | 字体常量（按平台选择） |
| `plugin/engine/build_docx.py:33` | 字体常量（按平台选择） |

---

## 3. 架构与数据流

### 3.1 宿主半（`lib/host.js`）

- 注入 `tools`、`skills`、`webServer`、`connection`。
- 工具：`screenwriter_stages`、`screenwriter_build_xlsx`、`screenwriter_build_docx`、`screenwriter_build_all`。
- 路由（都过 `connection.requestRejection` 信任闸）：
  - `GET /screenwriter/state`：项目卡、九阶段预设/会话、产物、知识库列表；
  - `POST /screenwriter/state`：保存项目与阶段修改；
  - `POST /screenwriter/build`：`{target:"xlsx"|"docx"|"all"}`，串行跑 Python；
  - `GET /screenwriter/files`：产物与知识库。
- 状态：`编剧工坊/plugin-state.json`（首次 GET 不落盘，点保存才写）。
- 生成前 `materialize()`：把当前项目信息合并进 `demo/_当前项目.json` 与 `demo/_当前剧本.json`，
  再交给引擎；原始 demo 数据不被覆盖。

### 3.2 浏览器半（`lib/client.js`）

- 手写 CJS 包体：`window.__ModuleLoader__.load({id, factory})`，只 `require("react")`。
- 左栏：九道工序，按 开发/结构/人物/成稿 分组。
- 中栏三个标签：
  - 会话窗口：预置“默认预设产出”，可追加要求；按钮「记入会话」「复制简报」；
  - 默认预设：Textarea 可编辑、可重置；
  - 阶段说明：自检问题、依据、知识库路径。
- 右栏：项目卡、生成按钮、交付物、流程进度。
- 所有请求用相对路径 `screenwriter/state` 等；从 `http://127.0.0.1:<port>/` 解析。

### 3.3 引擎（`engine/*.py`）

- `build_xlsx.py`：读 `项目数据.json` → 14 张可见表 + `_曲线数据` 隐藏表；openpyxl。
- `build_docx.py`：读 `剧本_第01集.json` → 封面/人物表/正文/尾页；python-docx。
- 两个脚本都只依赖标准库 + openpyxl / python-docx；不联网、不读 PDF。
- 输出一行 JSON 到 stdout，宿主取最后一行；失败把 stderr 尾部回传。

### 3.4 数据模型

`demo/项目数据.json` 顶层键：

```
project, genre_analysis, acts[8], beats[15], plot_points[5],
characters[9], relationships{codes,edges[],matrix[]},
scenes[10], episodes[12], ep1_scenes[14], tension_curve[12]
```

`demo/剧本_第01集.json`：

```
title, episode, episode_title, format, draft, date, logline,
characters[], scenes[{heading, elements[{type,text}]}]
type ∈ action | character | parenthetical | dialogue | transition
```

`plugin/presets/stages.json`：`stages[]`，每项含 `id/no/phase/name/role/output/outputSheet/icon/summary/preset/questions/knowledge`。
改预设时两端都要考虑：文件改的是**默认值**；工作台改的是**已保存值**（`plugin-state.json` 优先）。

### 3.5 九阶段速查

| id | 名称 | 主责书 | 产出 |
|---|---|---|---|
| `genre` | 类型选择 | 救猫咪 | xlsx `01 类型与卖点` |
| `outline` | 三幕八序列大纲 | 麦基 + 菲尔德 | xlsx `03 三幕八序列` |
| `theme` | 主题撰写 | 麦基 | Markdown（暂不单独成表） |
| `beats` | 详细故事小节 | 斯奈德 | xlsx `04 十五节拍` |
| `plot` | 情节点 | 菲尔德 | xlsx `05 情节点` / `05b 伏笔回馈` |
| `character` | 主要人物设计 | 打草稿 + 麦基 | xlsx `06 人物设计` / `07 人物关系` |
| `setting` | 主要场景设计 | 菲尔德 + 麦基 | xlsx `08 主要场景` |
| `sceneoutline` | 分场 / 分集大纲 | 打草稿 + 斯奈德 | xlsx `09 分集大纲` / `10 第01集分场` |
| `script` | 完整剧本 | 菲尔德 | docx `02_第01集_完整剧本` |

---

## 4. 已验 / 未验

### 已验（Windows，2026-10-18/19）

- `screenwriter_stages` → 9 条。
- `screenwriter_build_all` → xlsx 44.5 KB、docx 44.8 KB，均 `ok: true`。
- xlsx：14 张可见表；`openpyxl` 读回 sheet 名正确；情绪曲线为折线图。
- docx：13 页；正文 14 场；封面与尾页正常；无空白页（曾修过 page break 造成的一页空白）。
- 渲染检查：封面、03 三幕、05 情节点、06 人物、07 关系、09 分集+曲线、10 分场、13 尾页。
- 已修复的坑：
  1. 封面打印区域被裁成 A—B 两列 → 显式 `print_area` + 右下角版权行撑到 H 列；
  2. 人物/情节点子表头重复设置列宽，覆盖主表宽度 → 子表头不再传 widths；
  3. 分集曲线数据列露在表右侧 → 移到隐藏表 `_曲线数据`，打印区 A1:H；
  4. 尾页多一页空白 → 用 `page_break_before` 替代空段 + break；
  5. Windows 控制台 cp950 打印中文路径报错 → 引擎结果 JSON 用 `ensure_ascii=True`。

### 未验

- **界面点击流**：槽位注册成功、工具调用成功，但没有浏览器控制通道，未截图验收。
- **macOS**：未跑过；第 5 节是迁移计划。
- **Numbers / Pages**：未测；建议用 Excel / Word 或 LibreOffice 打开验收。

---

## 5. macOS 迁移清单（按顺序做）

### P0 · 复制项目、确认路径

把整个 `编剧工坊/` 目录拷到 Mac，保持：

```
<任意目录>/编剧工坊/{demo,产出,知识库,plugin}
```

插件数据根默认是 `plugin/` 的上一级，因此**保持这个相对结构**最省事。
如果插件被复制/安装进 profile 的 `node_modules`（而不是 pnpm link），
`DATA_ROOT` 会指错；此时启动 DSH 前设：

```bash
export SCREENWRITER_HOME="/Users/<you>/.../编剧工坊"
```

验证：在会话里调 `screenwriter_stages`（能返回即可），或看 `/screenwriter/state` 返回的 `root`
是否是项目目录（浏览器开发者工具 Network；`curl` 会被 401 信任闸挡住，属正常）。

### P1 · Python 与依赖

DSH 自带运行时通常在：

```bash
ls ~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3
```

找不到就搜索：

```bash
find ~/.dsh -maxdepth 6 -name 'python3*' -type f 2>/dev/null
```

检查依赖（必须两个都能 import）：

```bash
~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3 \
  -c "import sys, openpyxl, docx; print(sys.executable, openpyxl.__version__, docx.__version__)"
```

缺了就用同一个解释器装（不要污染系统 Python）：

```bash
PY=~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3
$PY -m pip install --user "openpyxl==3.1.5" "python-docx==1.2.0"
```

如果 DSH 自带运行时没有 pip，可先用 Homebrew Python 3.12 建 venv：

```bash
python3 -m venv .venv
.venv/bin/pip install "openpyxl==3.1.5" "python-docx==1.2.0"
export SCREENWRITER_PYTHON="$PWD/编剧工坊/.venv/bin/python"
```

**先绕过插件，直接跑引擎验收**（最快定位问题）：

```bash
cd "/Users/<you>/.../编剧工坊"
$PY plugin/engine/build_xlsx.py \
  --project demo/项目数据.json \
  --out "产出/01_剧作蓝图_回声镇.xlsx"
$PY plugin/engine/build_docx.py \
  --script demo/剧本_第01集.json \
  --out "产出/02_第01集_完整剧本_回声镇.docx"
```

两个都输出 `{"ok": true, ...}` 才继续。

### P2 · 字体

代码已按平台选字体（`sys.platform == "darwin"`）：

- 标题：`PingFang SC`（黑体）
- 正文/剧本：`Songti SC`（宋体）

macOS 自带这两款，一般无需安装。若要强制指定：

```bash
export SCREENWRITER_FONT_SANS="PingFang SC"
export SCREENWRITER_FONT_SERIF="Songti SC"
```

检查中文字体是否被系统识别：

```bash
fc-list :lang=zh | head
```

若用 LibreOffice 做渲染验收、且提示缺字体，装 Noto：

```bash
brew install --cask font-noto-sans-cjk font-noto-serif-cjk
```

字体缺失**不会导致生成失败**，只会被 Excel/Word 替换；若版面错位，再调字号/列宽。

### P3 · 安装插件

在 DSH 会话里用 `plugin_manager`：

```
action: install_bundle
target: /Users/<you>/.../编剧工坊/plugin
```

- 需要 Full access 或审批。
- 成功后看 `application: "applied"` 与 `warnings: []`。
- 刷新页面；必要时重启 DSH（宿主半模块有 ESM 缓存，改 host.js 必须重启）。
- 验证：侧栏出现「✎ 编剧工坊」；`cordis_inspect_query` 的 `sidebar.panellist` 有 `screenwriter`。

### P4 · 验收清单

1. 工作台打开不报错；九道工序都在。
2. 点「生成全部产物」→ 状态栏显示完成；`产出/` 出现两份文件。
3. Excel：14 张可见表；`09 分集大纲` 有情绪曲线；`06 人物设计` 列宽正常。
4. Word：13 页；封面、人物表、正文、尾页正常。
5. 保存预设 → 刷新 → 修改仍在（写入 `plugin-state.json`）。
6. 改项目名 → 生成 → 新文件名出现（旧文件不会自动删）。

命令行的等价检查：

```bash
$PY -c "import openpyxl; wb=openpyxl.load_workbook('产出/01_剧作蓝图_回声镇.xlsx'); print([ws.title for ws in wb.worksheets])"
$PY -c "from docx import Document; d=Document('产出/02_第01集_完整剧本_回声镇.docx'); print(len(d.paragraphs))"
```

### P5 · 常见故障

| 现象 | 原因 | 处理 |
|---|---|---|
| `spawn ... ENOENT` | 找不到 Python | 设 `SCREENWRITER_PYTHON`；确认可执行权限 |
| 产物写进 `node_modules` | 插件被复制安装，`DATA_ROOT` 指错 | 设 `SCREENWRITER_HOME`，重启 DSH |
| `ModuleNotFoundError: openpyxl` | 运行时缺依赖 | 用同一个解释器 `pip install --user` |
| 页面 401 | `/screenwriter/*` 走 connection 信任闸 | 用页面本身；不要用 curl 验 |
| 侧栏没有图标 | 没刷新 / 安装未应用 | 刷新、重启；查 `plugin_manager` 与槽位 |
| Excel 图表在 Numbers 里丢了 | Numbers 对 openpyxl 图表兼容差 | 用 Excel / LibreOffice 打开 |
| 中文文件名乱码 | 终端 locale | 用 UTF-8 终端；Python 3 内部无碍 |
| 字体被替换、版面偏 | mac 无微软雅黑/宋体 | 已自动换 PingFang/Songti；仍偏就调字号 |

---

## 6. 关键决策（改之前先读）

1. **九个阶段就是九个“子智能体”**：它们目前是**预设 + 会话窗口**，不是自动跑模型的进程。
   工作台负责给 brief、存修改、复制简报；真正创作在会话里完成。若要做成真正的 agent preset，
   见第 9 节。
2. **Excel 一页一张表**：照《打草稿》的工作表逻辑；不要把多张表塞进一张，打印会崩。
3. **05b / 07b 拆表**：情节点与伏笔、关系矩阵与关系清单，列宽需求不同；合并会互相覆盖列宽。
4. **隐藏表 `_曲线数据`**：情绪曲线的数据源；不要删，图表会断。若必须挪，改 `sheet_episodes`
   里的 `Reference`。
5. **数据与排版分离**：`demo/*.json` 只存内容；`engine/*.py` 只负责排版。改故事别动 Python，
   改版式别动 demo。
6. **生成是确定性的**：同样输入必得同样结构；创意内容不靠模型临时发挥。

---

## 7. 已知限制与风险

- **界面未做截图验收**：只能确认槽位注册与工具调用；首次在 Mac 打开时，按 P4 逐项看。
- **会话窗口不调模型**：不要向用户承诺“点一下子智能体就自动写完”；它给的是 brief 与记录。
- **保存会忽略空字符串**：`host.js` 的 POST 只接受非空值；想清空字段要改逻辑。
- **改标题会产生新文件**：旧产物不自动删；需要手动清理或加“清理旧产物”按钮。
- **OCR 只做了《打草稿》**：另外三本（《故事》《救猫咪》《悉德·菲尔德》）只有方法提要，
  没有逐页 OCR；要引用原文需另跑 OCR。
- **客户端是手写 CJS 包体**：依赖 `window.__ModuleLoader__.load` 格式；DSH 加载器改动时要跟着改。
- **Windows 路径示例**：README 里仍有 `E:\...` 示例；Mac 上以本文与环境变量为准。
- **Excel 样式依赖字体**：mac 上字体替换后列宽可能略变；以打印预览为准微调。

---

## 8. 下一步建议（按优先级）

1. **P0**：Mac 上跑通 P1—P4；把 `SCREENWRITER_PYTHON` / `SCREENWRITER_HOME` 写进启动脚本或 `.env`。
2. **P1**：把九个阶段注册成真正的 DSH agent preset（每个 preset 挂 `persona` + 本插件工具 +
   `screenwriter` 技能），让“子智能体”有独立会话入口。
3. **P2**：把第 1 集扩写到 45 分钟体量（约 45—50 页 Word）；同时补 `03_分集大纲.docx`。
4. **P3**：给工作台加“删除旧产物 / 打开产出目录 / 导出 zip”三个按钮。
5. **P4**：把四本书的关键章节做 OCR，补进 `知识库/`，让引用能落到页码。

---

## 9. 新会话启动 prompt（直接复制）

```text
你在 macOS 上接手一个已完成的 DSH 插件项目「编剧工坊」，路径：
/Users/<you>/.../screenwriter-tools/编剧工坊

先读：
1) HANDOFF.md（交接与迁移清单）
2) README.md（产物与用法）
3) plugin/README.md（插件结构）
不要从零重写，不要改九个阶段的设计。

任务：
- 按 HANDOFF 第 5 节 P1—P4 在 macOS 上跑通：Python 依赖 → 引擎直跑 →
  plugin_manager 安装 → 工作台验收。
- 保留 Windows 上已验证的行为：Excel 14 张可见表、Word 13 页、九阶段预设不变。
- 所有平台差异用环境变量解决，不要写死 /Users 路径：
  SCREENWRITER_PYTHON、SCREENWRITER_HOME、SCREENWRITER_FONT_SANS/SERIF。
- 先跑引擎、再装插件、最后开界面；每步给出实际输出，不要只描述。

验收：
- build_xlsx.py / build_docx.py 各输出 {"ok": true}
- screenwriter_build_all 返回 ok
- 侧栏出现「编剧工坊」，九阶段可切换，保存后刷新仍在
- Excel 表名与 Word 页数与 Windows 一致

边界：
- 工作台不直接调用模型；会话窗口只存预设与记录。
- 未经我确认，不要删除产出/、知识库/、demo/ 里的文件。
```

---

## 10. 本轮变更（2026-10-19）

为便于 Mac 接手，已做三处低风险可移植性改动，并在 Windows 上回归通过：

1. `host.js`：`DATA_ROOT` 支持 `SCREENWRITER_HOME` 覆盖；Python 候选按 `process.platform`
   分支（Windows 找 `python.exe`，macOS 找 `.../python/bin/python3`，再退回 `python3/python`）。
2. `build_xlsx.py` / `build_docx.py`：字体按 `sys.platform` 选择
   （macOS：PingFang SC / Songti SC；Linux：Noto CJK；Windows 保持微软雅黑/宋体），
   并支持 `SCREENWRITER_FONT_SANS` / `SCREENWRITER_FONT_SERIF` 覆盖。
3. 重新生成并检查两份产物：xlsx 与 docx 均 `ok`，Windows 字体与版面未变。

Mac 上仍需人工确认的只有三件事：**Python 依赖、插件安装路径、界面点击流**。

---

## 11. 附：Windows 基线复现命令

在 Windows 工作区根目录执行，用来和 Mac 上的结果对比：

```powershell
$py = "$env:USERPROFILE\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe"
$proj = "screenwriter-tools\编剧工坊"

& $py "$proj\plugin\engine\build_xlsx.py" `
  --project "$proj\demo\项目数据.json" `
  --out "$proj\产出\01_剧作蓝图_回声镇.xlsx"

& $py "$proj\plugin\engine\build_docx.py" `
  --script "$proj\demo\剧本_第01集.json" `
  --out "$proj\产出\02_第01集_完整剧本_回声镇.docx"

# 结构基线
& $py -c "import openpyxl; wb=openpyxl.load_workbook(r'$proj\产出\01_剧作蓝图_回声镇.xlsx'); print(len(wb.worksheets), sum(1 for ws in wb.worksheets if ws.sheet_state=='visible'))"
# 期望：15 14（15 张表、14 张可见，含隐藏表 _曲线数据）

& $py -c "from docx import Document; d=Document(r'$proj\产出\02_第01集_完整剧本_回声镇.docx'); print(len(d.paragraphs))"
# 期望：257
```

Word 页数用 LibreOffice 转 PDF 后数：Windows 基线是 **13 页**。
Excel 可见表是 **14 张**，加上隐藏表 `_曲线数据` 共 **15 张**。

---

## 12. 2026-10-26 本轮：三栏布局 + v2 功能 + 知识库 OCR

### 12.1 当前界面

- 保持三栏：左九道工序 218px、中阶段内容、右项目卡 292px。
- 右栏 `box-sizing: border-box`、各栏 `minWidth: 0`、主体 `overflow: hidden`，
  防止右栏被挤出视口。
- 2026-10-26 修了一个真正的布局 bug：`ScreenwriterPanel` 的主体里，右栏
  `SidePane` 少关了一层 `</div>`，实际被塞进了中间栏，所以它会「跑中间」并遮住内容。
  现在中间栏先闭合，`SidePane` 才是主体的第三个 flex 子项。
- 初始进入是**空白项目**：项目卡空、九个会话空、总控简报空。
- 《回声镇》不再自动塞进状态，而是右栏「载入示范案例」按钮 / `screenwriter_demo`
  工具载入的**示范预设案例**；载入后总控、生成、交付物都可用。

### 12.2 总控

- 右栏「总控 · 按项目卡推进九阶段」按钮 → `POST /screenwriter/project/sync`。
- 先本地自检项目卡（缺项、长度、类型/梗概/基调冲突、受众、预算），有冲突弹
  `CheckModal`，可返回修改或「仍然继续」。
- 继续后把项目卡写成九份总控简报，写进九个阶段会话；不直接调用模型写正文。
- 界面上的「生成 Excel + Word」在空白项目下禁用，载入示范案例后可用。

### 12.3 阶段说明可编辑

每个阶段的「阶段说明」现在三块都可逐条改、删、加：

- **交付前必须回答**：`questions[]`，编号列表，可新增/删除/编辑。
- **本阶段依据**：`basis[]`，可新增/删除/编辑。
- **知识库参考**：`kbRefs[]`，从知识库弹层选择文件挂到当前阶段，可单独移除。
- 三块都随「保存」写入 `plugin-state.json`，并进入该阶段的简报；默认值仍来自
  `presets/stages.json`。
- 顶栏「知识库」打开管理弹层：上传 `.md/.txt/.docx/.pdf`、删除上传文件、
  把文件挂到阶段。

### 12.4 知识库 OCR

- 新增 `plugin/engine/ocr_pdf.py`：从 `course-kit/plugin/engine/pdf_ocr.py` 裁剪，
  复用 `course-kit/.pylibs` 里的 pymupdf + RapidOCR，本地 CPU 识别，不联网、不上传。
- `ingest_knowledge.py` 的 PDF 路径：有文字层直接读，没文字层的页才 OCR；
  `.docx / .txt / .md` 行为不变。
- 默认单次最多 OCR **120 页**，超出部分跳过并在结果里报数；用 `SCREENWRITER_OCR_PAGES`
  可改（0 = 整本）。
- 依赖目录默认 `<工作区>/course-kit/.pylibs`，用 `SCREENWRITER_OCR_LIBS` 覆盖；
  OCR 用的 Python 默认 `Python310/python.exe`（`.pylibs` 按 3.10 安装），用
  `SCREENWRITER_OCR_PYTHON` 覆盖。Excel/Word 仍走 DSH 自带 Python。
- 上传后状态栏显示抽字数与 OCR 页数；扫描件识别期间按钮显示「OCR 中…」。

### 12.5 本轮验证

- `ocr_pdf.py` 对《打草稿》第 1 页实测：149 字、4.9 秒、`method=pymupdf+rapidocr`。
- `ingest_knowledge.py` 同上文件实测 `{"ok": true, "method": "pymupdf+rapidocr"}`。
- `node --check client.js / host.js` 通过；host 模块 import 通过；两个 Python 脚本 `py_compile` 通过。
- 状态已重置为空白：`demoLoaded: false`、项目卡全空、九个会话 0 条。
- 未做截图级验收；宿主改动要重启 DSH 才重载，浏览器半刷新页面可见。

### 12.6 仍需注意

- 总控不会自动写剧本，它只把项目卡变成九份可执行的 brief。
- OCR 走同步上传，页数多时请求会久；默认 120 页是速度与完整度的折中。
- 生成 Excel/Word 仍要求已有结构化数据（示范案例或导入 JSON）；空白项目只做规划。
