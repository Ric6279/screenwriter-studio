// 编剧工坊 —— 浏览器半（Client half）
//
// 九宫格工作台：左侧九道工序，中间是该阶段的「默认预设 / 会话窗口 / 说明」，
// 右侧是项目卡、交付物与知识库。所有写入都走宿主半 /screenwriter/* 路由。
//
// 静态客户端包要求预构建的 CJS 包体（window.__ModuleLoader__.load 格式）；
// 本文件不 import 任何 npm 包，只用 shell 提供的 react + 浏览器原生 API。

window.__ModuleLoader__.load({
  id: "@fatfish/screenwriter-studio",
  factory: (require) => {
    var module = { exports: {} };
    var exports = module.exports;
    Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });

    const react = require("react");
    const { useState, useEffect, useRef, useMemo, useCallback } = react;
    const h = react.createElement;

    const API_STATE = "screenwriter/state";
    const API_BUILD = "screenwriter/build";
    const API_FILES = "screenwriter/files";
    const API_DEMO = "screenwriter/demo";
    const API_NEW = "screenwriter/project/new";
    const API_CHECK = "screenwriter/project/check";
    const API_SYNC = "screenwriter/project/sync";
    const API_KB_UPLOAD = "screenwriter/knowledge/upload";
    const API_KB_DELETE = "screenwriter/knowledge/delete";

    const C = {
      paper: "#F4EFE3",
      card: "#FFFDF7",
      card2: "#FBF6EA",
      ink: "#1B2A38",
      inkSoft: "#3E4C59",
      muted: "#7C7466",
      line: "#DCD2BE",
      lineSoft: "#EAE2D0",
      teal: "#1F6F78",
      tealSoft: "#DCEAEA",
      amber: "#B07D2B",
      amberSoft: "#F2E4C8",
      red: "#9B2C2C",
      green: "#3F6B45",
      white: "#FFFFFF",
    };

    const FONT = "ui-sans-serif, system-ui, 'Segoe UI', 'Microsoft YaHei', sans-serif";
    const SERIF = "'Songti SC', 'SimSun', 'Noto Serif SC', serif";
    const MONO = "ui-monospace, 'Cascadia Mono', Consolas, monospace";

    const PHASES = ["开发", "结构", "人物", "成稿"];

    function fmtSize(n) {
      if (!n && n !== 0) return "";
      if (n < 1024) return n + " B";
      if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
      return (n / 1024 / 1024).toFixed(2) + " MB";
    }

    function fmtTime(ms) {
      if (!ms) return "";
      const d = new Date(ms);
      const p = (x) => String(x).padStart(2, "0");
      return `${d.getMonth() + 1}/${d.getDate()} ${p(d.getHours())}:${p(d.getMinutes())}`;
    }

    async function apiRaw(path, options) {
      const res = await fetch(path, {
        headers: { "content-type": "application/json" },
        ...options,
      });
      let body = null;
      try { body = await res.json(); } catch { body = null; }
      return { status: res.status, ok: res.ok, body };
    }

    async function api(path, options) {
      const { status, ok, body } = await apiRaw(path, options);
      if (!ok || !body || body.ok === false) {
        const msg = body && (body.message || (body.check && body.check.summary) || (body.results && body.results.map((r) => r.log).join("\n"))) || ("HTTP " + status);
        throw new Error(msg);
      }
      return body;
    }

    function copyText(text) {
      try {
        navigator.clipboard.writeText(text);
        return true;
      } catch {
        return false;
      }
    }

    function basename(path) {
      return String(path || "").split(/[\\/]/).pop();
    }

    function readFileBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = () => reject(new Error("读取文件失败"));
        reader.readAsDataURL(file);
      });
    }

    // ------------------------------------------------------------------ 小部件
    function Btn(props) {
      const color = props.color || C.teal;
      return h("button", {
        onClick: props.onClick,
        disabled: props.disabled,
        title: props.title,
        style: {
          cursor: props.disabled ? "not-allowed" : "pointer",
          background: props.disabled ? C.line : props.kind === "ghost" ? "transparent" : color,
          color: props.disabled ? C.muted : props.kind === "ghost" ? color : "#fff",
          border: props.kind === "ghost" ? "1px solid " + color : "1px solid transparent",
          borderRadius: 5, padding: props.small ? "4px 9px" : "7px 14px",
          fontSize: props.small ? 11 : 12.5, fontWeight: 600, lineHeight: 1.2,
          opacity: props.disabled ? 0.7 : 1, whiteSpace: "nowrap",
        },
      }, props.children);
    }

    function Label(props) {
      return h("div", {
        style: {
          fontSize: 10.5, letterSpacing: ".08em", textTransform: "uppercase",
          color: C.muted, marginBottom: 5, fontWeight: 700,
        },
      }, props.children);
    }

    function TextInput(props) {
      return h("input", {
        value: props.value || "",
        onChange: (e) => props.onChange(e.target.value),
        placeholder: props.placeholder,
        spellCheck: false,
        style: {
          width: "100%", boxSizing: "border-box", background: C.card,
          color: C.ink, border: "1px solid " + C.line, borderRadius: 4,
          padding: "6px 8px", fontSize: 12, outline: "none", fontFamily: FONT,
        },
      });
    }

    function TextArea(props) {
      return h("textarea", {
        value: props.value || "",
        onChange: (e) => props.onChange(e.target.value),
        placeholder: props.placeholder,
        spellCheck: false,
        rows: props.rows || 4,
        style: {
          width: "100%", boxSizing: "border-box", background: C.card,
          color: C.ink, border: "1px solid " + C.line, borderRadius: 4,
          padding: "7px 9px", fontSize: props.size || 12,
          lineHeight: 1.7, outline: "none", resize: "vertical",
          fontFamily: props.mono ? MONO : FONT,
          ...(props.style || {}),
        },
      });
    }

    function Dot(props) {
      const color = props.ok ? C.green : props.warn ? C.amber : C.line;
      return h("span", {
        style: {
          display: "inline-block", width: 7, height: 7, borderRadius: "50%",
          background: color, marginRight: 6, flex: "0 0 auto",
        },
      });
    }

    // ------------------------------------------------------------------ 弹层
    function Modal(props) {
      useEffect(() => {
        const onKey = (e) => { if (e.key === "Escape" && props.onClose) props.onClose(); };
        window.addEventListener("keydown", onKey);
        return () => window.removeEventListener("keydown", onKey);
      }, [props.onClose]);
      return h("div", {
        style: {
          position: "fixed", inset: 0, zIndex: 60,
          background: "rgba(20,28,36,.42)", backdropFilter: "blur(2px)",
          display: "flex", alignItems: "center", justifyContent: "center", padding: 24,
        },
        onMouseDown: (e) => { if (e.target === e.currentTarget && props.onClose) props.onClose(); },
      },
        h("div", {
          style: {
            width: props.width || 620, maxWidth: "92vw", maxHeight: "86vh",
            background: C.paper, border: "1px solid " + C.line, borderRadius: 10,
            boxShadow: "0 24px 80px rgba(20,28,36,.35)", display: "flex", flexDirection: "column",
            overflow: "hidden",
          },
        },
          h("div", {
            style: {
              flex: "0 0 auto", display: "flex", alignItems: "center", gap: 10,
              padding: "14px 18px", background: C.card, borderBottom: "1px solid " + C.line,
            },
          },
            h("div", { style: { flex: 1 } },
              h("div", { style: { fontSize: 15, fontWeight: 800, color: C.ink } }, props.title),
              props.subtitle ? h("div", { style: { fontSize: 11, color: C.muted, marginTop: 3, lineHeight: 1.5 } }, props.subtitle) : null),
            h("button", {
              onClick: props.onClose,
              style: {
                border: "1px solid " + C.line, background: "transparent", color: C.muted,
                borderRadius: 5, width: 28, height: 28, cursor: "pointer", fontSize: 15,
              },
            }, "×")),
          h("div", { style: { flex: "1 1 auto", minHeight: 0, overflow: "auto", padding: "16px 18px" } }, props.children)));
    }

    function EditableList(props) {
      const { items, onChange, placeholder, numbered, addLabel } = props;
      const list = Array.isArray(items) ? items : [];
      const update = (i, v) => onChange(list.map((x, j) => (j === i ? v : x)));
      const remove = (i) => onChange(list.filter((_, j) => j !== i));
      return h("div", { style: { display: "flex", flexDirection: "column", gap: 5 } },
        list.map((v, i) => h("div", {
          key: i,
          style: { display: "flex", alignItems: "center", gap: 7 },
        },
          h("span", {
            style: { width: 20, flex: "0 0 auto", color: C.amber, fontSize: 11, fontWeight: 700 },
          }, numbered ? String(i + 1).padStart(2, "0") : "•"),
          h("input", {
            value: v,
            onChange: (e) => update(i, e.target.value),
            placeholder,
            spellCheck: false,
            style: {
              flex: 1, minWidth: 0, border: "1px solid transparent",
              borderBottom: "1px dashed " + C.line, background: "transparent",
              color: C.inkSoft, fontSize: 12, lineHeight: 1.7, padding: "3px 2px", outline: "none",
              fontFamily: FONT,
            },
          }),
          h("button", {
            onClick: () => remove(i),
            title: "删除",
            style: {
              flex: "0 0 auto", border: "none", background: "transparent",
              color: C.muted, cursor: "pointer", fontSize: 14, lineHeight: 1,
            },
          }, "×"))),
        h("button", {
          onClick: () => onChange([...list, ""]),
          style: {
            alignSelf: "flex-start", marginTop: 2, border: "1px dashed " + C.line,
            background: "transparent", color: C.teal, borderRadius: 4,
            padding: "3px 8px", fontSize: 10.5, cursor: "pointer",
          },
        }, addLabel || "+ 添加"));
    }

    function KnowledgeModal(props) {
      const { open, files, targetStage, onClose, onUpload, onDelete, onAdd, uploading } = props;
      const fileRef = useRef(null);
      if (!open) return null;
      return h(Modal, {
        title: "知识库",
        subtitle: "上传剧本、编剧书、案例拆解；摘要会自动抽取，供九个子智能体参考。支持 .md / .txt / .docx / .pdf。",
        width: 720,
        onClose,
      },
        h("div", {
          style: {
            display: "flex", alignItems: "center", gap: 10, padding: "10px 12px",
            background: C.card, border: "1px dashed " + C.line, borderRadius: 6, marginBottom: 14,
          },
        },
          h("input", {
            ref: fileRef, type: "file", multiple: true,
            accept: ".md,.txt,.docx,.pdf,.doc,.rtf,.epub",
            style: { display: "none" },
            onChange: async (e) => {
              const list = [...(e.target.files || [])];
              e.target.value = "";
              for (const f of list) await onUpload(f);
            },
          }),
          h(Btn, { onClick: () => fileRef.current && fileRef.current.click(), disabled: uploading },
            uploading ? "上传中…" : "上传文件"),
          h("span", { style: { fontSize: 10.5, color: C.muted, lineHeight: 1.6 } },
            "≤ 30 MB / 个。PDF 若是扫描件，抽不出文字时会保留原文件，建议先用 OCR 处理。"),
          h("span", { style: { flex: 1 } })),
        h("div", { style: { display: "flex", flexDirection: "column", gap: 6 } },
          (files || []).length === 0
            ? h("div", { style: { fontSize: 11.5, color: C.muted, padding: "12px 2px" } }, "还没有上传文件。")
            : files.map((f) => {
              const deletable = f.kind !== "method";
              return h("div", {
                key: f.path,
                style: {
                  display: "flex", alignItems: "center", gap: 8,
                  background: C.card, border: "1px solid " + C.lineSoft,
                  borderRadius: 5, padding: "7px 9px",
                },
              },
                h("span", {
                  style: {
                    fontSize: 9, color: f.kind === "method" ? C.teal : C.amber,
                    border: "1px solid " + (f.kind === "method" ? "#C9DEDC" : "#E4D2AE"),
                    borderRadius: 3, padding: "1px 5px", flex: "0 0 auto",
                  },
                }, f.kind === "method" ? "方法库" : f.kind === "extract" ? "摘录" : "上传"),
                h("span", {
                  title: f.path,
                  style: { flex: 1, minWidth: 0, fontSize: 11.5, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
                }, f.name),
                h("span", { style: { fontSize: 10, color: C.muted, flex: "0 0 auto" } }, fmtSize(f.size)),
                targetStage
                  ? h(Btn, { kind: "ghost", small: true, color: C.teal, onClick: () => onAdd(targetStage, f.path) }, "加入依据")
                  : null,
                h(Btn, { kind: "ghost", small: true, color: C.muted, onClick: () => copyText(f.path) }, "路径"),
                deletable
                  ? h(Btn, { kind: "ghost", small: true, color: C.red, onClick: () => onDelete(f) }, "删除")
                  : null);
            })));
    }

    function CheckModal(props) {
      const { data, onClose, onForce, busy } = props;
      if (!data) return null;
      const { check } = data;
      const level = check.level;
      const tone = level === "ok" ? C.green : level === "warn" ? C.amber : C.red;
      return h(Modal, {
        title: "项目卡自检",
        subtitle: "这是本地关键词/长度检查，不调用模型；只用来拦一下“明显不是同一部戏”的填写。",
        width: 600,
        onClose,
      },
        h("div", { style: { display: "flex", alignItems: "center", gap: 14, marginBottom: 14 } },
          h("div", {
            style: {
              width: 68, height: 68, borderRadius: "50%", flex: "0 0 auto",
              border: "4px solid " + tone, display: "flex", alignItems: "center",
              justifyContent: "center", flexDirection: "column", color: tone,
            },
          },
            h("span", { style: { fontSize: 20, fontWeight: 800, lineHeight: 1 } }, String(check.score)),
            h("span", { style: { fontSize: 9, marginTop: 2 } }, "自检分")),
          h("div", null,
            h("div", { style: { fontSize: 13, fontWeight: 700, color: tone } },
              level === "ok" ? "基本自洽" : level === "warn" ? "有风格/信息冲突" : "缺口较多"),
            h("div", { style: { fontSize: 11.5, color: C.muted, marginTop: 4, lineHeight: 1.6 } }, check.summary))),
        h("div", { style: { display: "flex", flexDirection: "column", gap: 7, marginBottom: 16 } },
          (check.issues || []).length === 0
            ? h("div", { style: { fontSize: 12, color: C.green } }, "没有发现问题。")
            : check.issues.map((it, i) => h("div", {
              key: i,
              style: {
                display: "flex", gap: 8, alignItems: "flex-start",
                background: C.card, border: "1px solid " + C.lineSoft,
                borderRadius: 5, padding: "8px 10px",
              },
            },
              h("span", {
                style: {
                  flex: "0 0 auto", fontSize: 9, padding: "1px 5px", borderRadius: 3,
                  color: it.severity === "error" ? "#fff" : C.amber,
                  background: it.severity === "error" ? C.red : C.amberSoft,
                },
              }, it.severity === "error" ? "缺口" : "提醒"),
              h("span", { style: { fontSize: 11.5, lineHeight: 1.7, color: C.inkSoft } }, it.message)))),
        h("div", { style: { display: "flex", gap: 8, justifyContent: "flex-end" } },
          h(Btn, { kind: "ghost", color: C.muted, onClick: onClose }, "返回修改"),
          h(Btn, { color: C.red, onClick: onForce, disabled: busy }, busy ? "同步中…" : "仍然继续")));
    }


    function StageRail(props) {
      const { stages, activeId, onPick } = props;
      return h("div", {
        style: {
          width: 218, flex: "0 0 auto", background: C.ink, color: "#F1EAD9",
          display: "flex", flexDirection: "column", overflow: "hidden",
        },
      },
        h("div", { style: { padding: "14px 16px 10px" } },
          h("div", { style: { fontSize: 15, fontWeight: 800, letterSpacing: ".12em" } }, "编 剧 工 坊"),
          h("div", { style: { fontSize: 10, color: "#B9AE97", marginTop: 4, letterSpacing: ".05em" } },
            "SCREENWRITER · 四书九道工序")),
        h("div", { style: { flex: "1 1 auto", overflow: "auto", padding: "4px 8px 16px" } },
          PHASES.map((phase) => {
            const list = stages.filter((s) => s.phase === phase);
            if (list.length === 0) return null;
            return h("div", { key: phase, style: { marginTop: 10 } },
              h("div", {
                style: {
                  fontSize: 9.5, letterSpacing: ".18em", color: "#8E8574",
                  padding: "4px 10px", fontWeight: 700,
                },
              }, phase),
              list.map((s) => {
                const active = s.id === activeId;
                return h("div", {
                  key: s.id,
                  onClick: () => onPick(s.id),
                  style: {
                    display: "flex", alignItems: "center", gap: 8,
                    padding: "8px 10px", marginBottom: 2, borderRadius: 6,
                    cursor: "pointer",
                    background: active ? "rgba(220,234,234,.12)" : "transparent",
                    borderLeft: "2px solid " + (active ? "#5FB3B8" : "transparent"),
                    color: active ? "#FFFFFF" : "#DCD2BE",
                  },
                },
                  h("span", {
                    style: {
                      width: 20, textAlign: "center", fontSize: 12,
                      color: active ? "#8FD0D3" : "#8E8574",
                    },
                  }, s.icon || String(s.no)),
                  h("span", { style: { flex: 1, fontSize: 12.5, fontWeight: active ? 700 : 500 } }, s.name),
                  h("span", { style: { fontSize: 9.5, color: "#8E8574" } }, String(s.no).padStart(2, "0")));
              }));
          })),
        h("div", {
          style: {
            padding: "10px 16px 14px", borderTop: "1px solid rgba(255,255,255,.08)",
            fontSize: 10, color: "#8E8574", lineHeight: 1.7,
          },
        },
          "麦基 · 斯奈德 · 菲尔德 · 打草稿",
          h("br"),
          "预设可改，会话可续。"));

    }

    // ------------------------------------------------------------------ 会话窗口
    function ChatPane(props) {
      const { stage, onAdd, draft, setDraft, onCopyBrief } = props;
      const listRef = useRef(null);
      useEffect(() => {
        if (listRef.current) listRef.current.scrollTop = listRef.current.scrollHeight;
      }, [stage.session.length, stage.brief]);
      const empty = stage.session.length === 0 && !stage.brief;
      return h("div", { style: { display: "flex", flexDirection: "column", height: "100%", minHeight: 0 } },
        h("div", {
          ref: listRef,
          style: { flex: "1 1 auto", overflow: "auto", padding: "18px 20px 8px", minHeight: 0 },
        },
          stage.brief ? h("div", {
            style: {
              background: C.amberSoft, border: "1px solid #E3CFA5", borderRadius: 7,
              padding: "11px 14px", marginBottom: 16,
            },
          },
            h("div", {
              style: {
                fontSize: 9.5, letterSpacing: ".12em", color: C.amber,
                fontWeight: 700, marginBottom: 6,
              },
            }, "总控简报" + (stage.briefAt ? " · " + fmtTime(stage.briefAt) : "")),
            h("div", {
              style: { fontSize: 11.5, lineHeight: 1.8, color: C.inkSoft, whiteSpace: "pre-wrap", fontFamily: SERIF },
            }, stage.brief)) : null,

          empty ? h("div", {
            style: {
              height: "100%", minHeight: 220, display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center", gap: 8, color: C.muted,
            },
          },
            h("div", { style: { fontSize: 26, color: "#C9BFA9" } }, stage.icon || "✎"),
            h("div", { style: { fontSize: 13, fontWeight: 700, color: C.inkSoft } }, "这个阶段还没有会话"),
            h("div", { style: { fontSize: 11.5, lineHeight: 1.8, maxWidth: 420, textAlign: "center" } },
              "先在右侧填写项目卡，点「总控」把设定同步给九个阶段；也可以直接在下方输入你的要求。")) : null,

          stage.session.map((m, i) => {
            if (m.role === "system") {
              return h("div", {
                key: i,
                style: {
                  textAlign: "center", fontSize: 10.5, color: C.muted,
                  margin: "4px 0 12px", lineHeight: 1.6,
                },
              }, "— " + m.text + (m.time ? " · " + fmtTime(m.time) : "") + " —");
            }
            const mine = m.role === "user";
            return h("div", {
              key: i,
              style: {
                display: "flex", justifyContent: mine ? "flex-end" : "flex-start",
                marginBottom: 12,
              },
            },
              h("div", {
                style: {
                  maxWidth: "86%", whiteSpace: "pre-wrap", wordBreak: "break-word",
                  background: mine ? C.teal : C.card,
                  color: mine ? "#F4FBFA" : C.ink,
                  border: "1px solid " + (mine ? C.teal : C.line),
                  borderRadius: mine ? "12px 12px 3px 12px" : "12px 12px 12px 3px",
                  padding: "10px 13px", fontSize: 12.5, lineHeight: 1.75,
                  fontFamily: SERIF,
                  boxShadow: "0 1px 2px rgba(27,42,56,.05)",
                },
              },
                m.seed ? h("div", {
                  style: {
                    fontSize: 9.5, letterSpacing: ".1em", color: C.amber,
                    marginBottom: 6, fontFamily: FONT, fontWeight: 700,
                  },
                }, "示范案例 · 默认预设产出" + (m.time ? " · " + fmtTime(m.time) : "")) : null,
                m.text));
          })),
        h("div", {
          style: {
            flex: "0 0 auto", borderTop: "1px solid " + C.line,
            background: C.card2, padding: "10px 14px 12px",
          },
        },
          h(TextArea, {
            value: draft, onChange: setDraft, rows: 3,
            placeholder: "在这个阶段的会话里补充要求：改哪一拍、换哪个人物、加哪场戏……",
          }),
          h("div", { style: { display: "flex", gap: 8, marginTop: 8, alignItems: "center" } },
            h(Btn, {
              onClick: () => { if (draft.trim()) { onAdd(draft.trim()); setDraft(""); } },
              disabled: !draft.trim(),
            }, "记入会话"),
            h(Btn, {
              kind: "ghost", color: C.amber, onClick: onCopyBrief,
              title: "复制一份可直接粘到新会话的完整简报（项目卡 + 总控简报 + 预设 + 会话）",
            }, "复制简报"),
            h("span", { style: { flex: 1 } }),
            h("span", { style: { fontSize: 10.5, color: C.muted } },
              stage.outputSheet ? "产出：Excel · " + stage.outputSheet : "产出：Word / Markdown"))));
    }

    // ------------------------------------------------------------------ 预设
    function PresetPane(props) {
      const { stage, defaults, onChange, onReset } = props;
      const changed = stage.preset !== defaults[stage.id];
      return h("div", { style: { display: "flex", flexDirection: "column", height: "100%", minHeight: 0 } },
        h("div", {
          style: {
            flex: "0 0 auto", padding: "12px 18px", display: "flex", gap: 8,
            alignItems: "center", borderBottom: "1px solid " + C.lineSoft, background: C.card2,
          },
        },
          h("span", { style: { fontSize: 11, color: C.muted } },
            changed ? "已修改（与默认预设不同）" : "当前为默认预设"),
          h("span", { style: { flex: 1 } }),
          h(Btn, { kind: "ghost", small: true, color: C.amber, onClick: onReset, disabled: !changed }, "重置为默认")),
        h("div", { style: { flex: "1 1 auto", padding: "14px 18px", minHeight: 0, display: "flex" } },
          h(TextArea, {
            value: stage.preset, onChange, mono: true, size: 12,
            style: { flex: 1 },
          })));
    }

    // ------------------------------------------------------------------ 阶段说明
    function InfoPane(props) {
      const { stage, onPatch, onOpenKb } = props;
      const refs = Array.isArray(stage.kbRefs) ? stage.kbRefs : [];
      return h("div", { style: { overflow: "auto", padding: "18px 22px 28px", height: "100%", minWidth: 0, boxSizing: "border-box" } },
        h("div", { style: { display: "flex", gap: 10, alignItems: "baseline", marginBottom: 8 } },
          h("span", { style: { fontSize: 26, color: C.amber } }, stage.icon || "·"),
          h("div", null,
            h("div", { style: { fontSize: 15, fontWeight: 800, color: C.ink } }, stage.role),
            h("div", { style: { fontSize: 11.5, color: C.muted, marginTop: 2 } },
              "阶段 " + stage.no + " · " + stage.phase + " · " + stage.output))),
        h("div", {
          style: {
            fontSize: 13, lineHeight: 1.9, color: C.inkSoft, fontFamily: SERIF,
            background: C.card, border: "1px solid " + C.line, borderRadius: 6,
            padding: "12px 14px", margin: "10px 0 18px",
          },
        }, stage.summary),

        h("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 6 } },
          h(Label, null, "交付前必须回答"),
          h("span", { style: { fontSize: 9.5, color: C.muted, marginBottom: 5 } }, "可直接改、可删")),
        h(EditableList, {
          items: stage.questions || [],
          onChange: (v) => onPatch({ questions: v }),
          placeholder: "这个阶段交付前要回答的问题",
          numbered: true,
          addLabel: "+ 添加一问",
        }),

        h("div", { style: { height: 18 } }),
        h("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 6 } },
          h(Label, null, "本阶段依据"),
          h("span", { style: { fontSize: 9.5, color: C.muted, marginBottom: 5 } }, "书/方法/案例，可改可删")),
        h(EditableList, {
          items: stage.basis || [],
          onChange: (v) => onPatch({ basis: v }),
          placeholder: "例如：01_故事_麦基.md · 第 8 章 激励事件",
          numbered: false,
          addLabel: "+ 添加依据",
        }),

        h("div", { style: { height: 18 } }),
        h("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 6 } },
          h(Label, null, "知识库参考"),
          h("span", { style: { fontSize: 9.5, color: C.muted, marginBottom: 5 } }, "本阶段独有的参考文件")),
        refs.length === 0
          ? h("div", { style: { fontSize: 11.5, color: C.muted, lineHeight: 1.7, marginBottom: 8 } },
            "还没有挂载参考文件。可以从知识库把剧本、编剧书或拆解案例加进来。")
          : h("div", { style: { display: "flex", flexDirection: "column", gap: 5, marginBottom: 8 } },
            refs.map((p, i) => h("div", {
              key: p + i,
              style: {
                display: "flex", alignItems: "center", gap: 8,
                background: C.card, border: "1px solid " + C.lineSoft,
                borderRadius: 4, padding: "6px 9px",
              },
            },
              h("span", { style: { color: C.teal, fontSize: 11 } }, "▤"),
              h("span", {
                title: p,
                style: { flex: 1, minWidth: 0, fontSize: 11.5, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
              }, basename(p)),
              h("button", {
                onClick: () => copyText(p),
                style: { border: "none", background: "transparent", color: C.muted, cursor: "pointer", fontSize: 10.5 },
              }, "路径"),
              h("button", {
                onClick: () => onPatch({ kbRefs: refs.filter((_, j) => j !== i) }),
                title: "从本阶段移除",
                style: { border: "none", background: "transparent", color: C.muted, cursor: "pointer", fontSize: 14, lineHeight: 1 },
              }, "×")))),
        h("div", { style: { display: "flex", gap: 8, marginTop: 4 } },
          h(Btn, { kind: "ghost", small: true, color: C.teal, onClick: () => onOpenKb(stage.id) }, "+ 从知识库添加"),
          h(Btn, { kind: "ghost", small: true, color: C.muted, onClick: () => onOpenKb(null) }, "管理知识库")),

        h("div", {
          style: {
            marginTop: 22, fontSize: 10, color: C.muted, lineHeight: 1.7,
            borderTop: "1px dashed " + C.line, paddingTop: 10,
          },
        }, "这些内容会随「保存」写入 plugin-state.json，并作为该子智能体的工作说明；默认预设仍可在「默认预设」里改。"));
    }

    // ------------------------------------------------------------------ 右栏
    function SidePane(props) {
      const { project, onChange, outputs, outDir, root, onGenerate, onSync, busy, onRefresh, onDemo, onNew, demoLoaded, canBuild } = props;
      const fields = [
        ["title", "项目名"],
        ["format", "格式"],
        ["genre", "类型"],
        ["audience", "受众"],
        ["budget", "预算"],
        ["tone", "基调"],
      ];
      const blank = !project.title;
      return h("div", {
        style: {
          width: 292, flex: "0 0 auto", borderLeft: "1px solid " + C.line,
          background: C.card2, overflowY: "auto", overflowX: "hidden",
          boxSizing: "border-box", padding: "14px 14px 20px",
        },
      },
        h("div", { style: { display: "flex", alignItems: "center", gap: 8 } },
          h(Label, null, "项目卡"),
          demoLoaded ? h("span", {
            style: {
              fontSize: 9, color: C.amber, border: "1px solid #E3CFA5",
              background: C.amberSoft, borderRadius: 3, padding: "1px 5px", marginBottom: 4,
            },
          }, "示范案例") : null),

        blank ? h("div", {
          style: {
            background: C.card, border: "1px dashed " + C.line, borderRadius: 6,
            padding: "14px 12px", marginBottom: 12,
          },
        },
          h("div", { style: { fontSize: 12.5, fontWeight: 700, color: C.inkSoft, marginBottom: 5 } }, "还没有项目"),
          h("div", { style: { fontSize: 10.5, color: C.muted, lineHeight: 1.7, marginBottom: 10 } },
            "可以直接在下面填写项目卡，也可以先载入《回声镇》示范案例，看九个阶段怎么跑。"),
          h(Btn, { small: true, onClick: onDemo }, "载入示范案例")) : null,

        h("div", { style: { display: "flex", flexDirection: "column", gap: 8, marginBottom: 12 } },
          fields.map(([key, label]) => h("div", { key },
            h("div", { style: { fontSize: 10, color: C.muted, marginBottom: 2 } }, label),
            h(TextInput, { value: project[key], onChange: (v) => onChange(key, v) }))),
          h("div", null,
            h("div", { style: { fontSize: 10, color: C.muted, marginBottom: 2 } }, "一句话梗概"),
            h(TextArea, { value: project.logline, onChange: (v) => onChange("logline", v), rows: 4, size: 11.5 })),
          h("div", null,
            h("div", { style: { fontSize: 10, color: C.muted, marginBottom: 2 } }, "主题"),
            h(TextArea, { value: project.theme, onChange: (v) => onChange("theme", v), rows: 2, size: 11.5 }))),

        h("button", {
          onClick: onSync,
          disabled: busy,
          title: "先做项目卡自检，再把最新设定同步给九个阶段",
          style: {
            width: "100%", boxSizing: "border-box", cursor: busy ? "not-allowed" : "pointer",
            background: busy ? C.line : C.amber, color: "#fff", border: "none",
            borderRadius: 6, padding: "9px 10px", fontSize: 12.5, fontWeight: 800,
            letterSpacing: ".02em", marginBottom: 4,
          },
        }, busy ? "处理中…" : "总控 · 按项目卡推进九阶段"),
        h("div", { style: { fontSize: 9.5, color: C.muted, lineHeight: 1.6, marginBottom: 10 } },
          "自检项目卡 → 生成总控简报 → 写入九个阶段会话。"),

        h("div", { style: { display: "flex", gap: 7, margin: "2px 0 12px" } },
          h(Btn, {
            onClick: onGenerate,
            disabled: busy || !canBuild,
            title: canBuild ? "按示范数据结构重建 Excel/Word" : "空白项目还没有结构化内容可生成",
          }, busy ? "生成中…" : "生成 Excel + Word"),
          h(Btn, { kind: "ghost", small: false, color: C.muted, onClick: onRefresh }, "刷新")),
        !canBuild ? h("div", {
          style: { fontSize: 9.5, color: C.muted, lineHeight: 1.6, marginBottom: 12 },
        }, "生成按钮需要示范数据（载入示范案例后可用）；空白项目请先在九个阶段制作内容。") : null,

        demoLoaded ? h("div", {
          style: { fontSize: 10, color: C.muted, marginBottom: 12 },
        },
          h("button", {
            onClick: onNew,
            style: {
              border: "none", background: "transparent", color: C.red,
              cursor: "pointer", fontSize: 10, padding: 0, textDecoration: "underline",
            },
          }, "清空演示，回到空白项目")) : null,

        h(Label, null, demoLoaded ? "交付物" : "历史产物"),
        !demoLoaded && (outputs || []).length > 0 ? h("div", {
          style: { fontSize: 9.5, color: C.muted, lineHeight: 1.6, marginBottom: 6 },
        }, "空白项目不会重新生成；载入示范案例后可重建。") : null,
        h("div", { style: { display: "flex", flexDirection: "column", gap: 6, marginBottom: 8 } },
          (outputs || []).length === 0
            ? h("div", { style: { fontSize: 11, color: C.muted } }, "还没有产物。")
            : outputs.map((f) => h("div", {
              key: f.path,
              title: f.path,
              onClick: () => copyText(f.path),
              style: {
                background: C.card, border: "1px solid " + C.line, borderRadius: 5,
                padding: "7px 9px", cursor: "pointer",
              },
            },
              h("div", {
                style: {
                  fontSize: 11.5, color: C.ink, fontWeight: 600, lineHeight: 1.4,
                  overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                },
              }, (f.name.endsWith(".xlsx") ? "▦ " : "✎ ") + f.name),
              h("div", { style: { fontSize: 10, color: C.muted, marginTop: 3 } },
                fmtSize(f.size) + " · " + fmtTime(f.mtime))))),
        h("div", {
          style: { fontSize: 9.5, color: C.muted, lineHeight: 1.6, marginBottom: 16, wordBreak: "break-all" },
        }, "点击复制路径 · " + (outDir || root || "")),

        h(Label, null, "流程"),
        h("div", { style: { display: "flex", flexDirection: "column", gap: 3 } },
          [["01", "类型与卖点", "01 类型与卖点"], ["02", "三幕八序列", "03 三幕八序列"],
           ["03", "十五节拍", "04 十五节拍"], ["04", "情节点", "05 情节点"],
           ["05", "人物与关系", "06 人物设计"], ["06", "主要场景", "08 主要场景"],
           ["07", "分集分场", "09 分集大纲"], ["08", "完整剧本", null]].map(([n, label, sheet]) => {
            const done = !!(outputs || []).length;
            return h("div", {
              key: n,
              style: { display: "flex", alignItems: "center", fontSize: 11, color: C.inkSoft, padding: "2px 0" },
            },
              h(Dot, { ok: done }),
              h("span", { style: { width: 22, color: C.muted, fontSize: 10 } }, n),
              h("span", { style: { flex: 1 } }, label),
              h("span", { style: { fontSize: 9.5, color: C.muted } }, sheet ? "xlsx" : "docx"));
          })));
    }

    // ------------------------------------------------------------------ 主面板
    function ScreenwriterPanel() {
      const [loading, setLoading] = useState(true);
      const [error, setError] = useState("");
      const [project, setProject] = useState({});
      const [stages, setStages] = useState([]);
      const [defaults, setDefaults] = useState({});
      const [knowledge, setKnowledge] = useState([]);
      const [outputs, setOutputs] = useState([]);
      const [root, setRoot] = useState("");
      const [outDir, setOutDir] = useState("");
      const [activeId, setActiveId] = useState("outline");
      const [tab, setTab] = useState("chat");
      const [drafts, setDrafts] = useState({});
      const [busy, setBusy] = useState(false);
      const [status, setStatus] = useState("");
      const [demoLoaded, setDemoLoaded] = useState(false);
      const [checkModal, setCheckModal] = useState(null);
      const [kbOpen, setKbOpen] = useState(false);
      const [kbTarget, setKbTarget] = useState(null);
      const [uploading, setUploading] = useState(false);

      const stage = useMemo(
        () => stages.find((s) => s.id === activeId) || stages[0] || null,
        [stages, activeId]
      );

      async function load() {
        setLoading(true);
        try {
          const data = await api(API_STATE, { method: "GET" });
          setProject(data.project || {});
          setStages(data.stages || []);
          setDefaults(data.defaults || {});
          setKnowledge(data.knowledge || []);
          setOutputs(data.outputs || []);
          setRoot(data.root || "");
          setOutDir(data.outDir || "");
          setDemoLoaded(Boolean(data.demoLoaded));
          setError("");
          if (!activeId && data.stages && data.stages[0]) setActiveId(data.stages[0].id);
        } catch (e) {
          setError(String(e.message || e));
        } finally {
          setLoading(false);
        }
      }

      useEffect(() => { load(); }, []);

      const patchStage = useCallback((id, patch) => {
        setStages((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch } : s)));
      }, []);

      async function save(silent) {
        const payload = {
          project,
          stages: stages.map((s) => ({
            id: s.id,
            preset: s.preset,
            session: s.session,
            notes: s.notes,
            brief: s.brief,
            briefAt: s.briefAt,
            questions: s.questions,
            basis: s.basis,
            kbRefs: s.kbRefs,
          })),
        };
        try {
          await api(API_STATE, { method: "POST", body: JSON.stringify(payload) });
          if (!silent) setStatus("已保存到 plugin-state.json · " + new Date().toLocaleTimeString());
          return true;
        } catch (e) {
          setStatus("保存失败：" + e.message);
          return false;
        }
      }

      async function generate() {
        setBusy(true);
        setStatus("正在生成…");
        await save(true);
        try {
          const data = await api(API_BUILD, { method: "POST", body: JSON.stringify({ target: "all" }) });
          setOutputs(data.outputs || []);
          const ok = data.results.filter((r) => r.ok).length;
          setStatus(`生成完成：${ok}/${data.results.length} · ` + new Date().toLocaleTimeString());
        } catch (e) {
          setStatus("生成失败：" + e.message);
        } finally {
          setBusy(false);
        }
      }

      async function refresh() {
        try {
          const data = await api(API_FILES, { method: "GET" });
          setOutputs(data.outputs || []);
          setKnowledge(data.knowledge || []);
          setStatus("产物已刷新");
        } catch (e) {
          setStatus("刷新失败：" + e.message);
        }
      }

      async function loadDemo() {
        if (!window.confirm("载入示范案例会覆盖当前项目卡与九阶段会话，继续？")) return;
        setBusy(true);
        try {
          await api(API_DEMO, { method: "POST", body: "{}" });
          await load();
          setStatus("已载入《回声镇》示范案例 · 可直接查看九个阶段");
        } catch (e) {
          setStatus("载入失败：" + e.message);
        } finally {
          setBusy(false);
        }
      }

      async function newProject() {
        if (!window.confirm("清空当前项目卡与九阶段会话，回到空白项目？")) return;
        setBusy(true);
        try {
          await api(API_NEW, { method: "POST", body: "{}" });
          await load();
          setStatus("已回到空白项目");
        } catch (e) {
          setStatus("清空失败：" + e.message);
        } finally {
          setBusy(false);
        }
      }

      async function syncProject(force) {
        setBusy(true);
        setStatus(force ? "项目卡有冲突，仍然同步…" : "正在自检并同步项目卡…");
        try {
          const res = await apiRaw(API_SYNC, {
            method: "POST",
            body: JSON.stringify({ project, force: force === true }),
          });
          if (res.status === 409 && res.body && res.body.code === "PROJECT_CHECK") {
            setCheckModal({ check: res.body.check });
            setStatus("自检发现冲突，等你确认");
            return;
          }
          if (!res.ok || !res.body || res.body.ok === false) {
            throw new Error((res.body && (res.body.message || res.body.check?.summary)) || ("HTTP " + res.status));
          }
          if (Array.isArray(res.body.stages)) setStages(res.body.stages);
          if (res.body.project) setProject(res.body.project);
          setStatus(`总控完成：自检 ${res.body.check?.score ?? "-"} 分，已同步 ${res.body.synced ?? 9} 个阶段 · ` + new Date().toLocaleTimeString());
          setTab("chat");
        } catch (e) {
          setStatus("同步失败：" + e.message);
        } finally {
          setBusy(false);
        }
      }

      async function uploadKnowledge(file) {
        if (file.size > 30 * 1024 * 1024) {
          setStatus("上传失败：文件超过 30 MB");
          return;
        }
        setUploading(true);
        setStatus(`正在读取「${file.name}」…扫描版 PDF 会在本机 OCR，页数多时要等一会儿`);
        try {
          const dataUrl = await readFileBase64(file);
          const res = await api(API_KB_UPLOAD, {
            method: "POST",
            body: JSON.stringify({ name: file.name, data: dataUrl }),
          });
          setKnowledge(res.knowledge || knowledge);
          const chars = res.ingest?.result?.chars;
          const ocrPages = res.ingest?.result?.ocr?.ocrPages;
          setStatus(
            `已上传 ${file.name}` +
            (chars ? `，抽到 ${chars} 字` : "，但没有抽到文字") +
            (ocrPages ? `（OCR ${ocrPages} 页）` : "")
          );
        } catch (e) {
          setStatus("上传失败：" + e.message);
        } finally {
          setUploading(false);
        }
      }

      async function deleteKnowledge(f) {
        if (!window.confirm(`删除「${f.name}」？`)) return;
        try {
          const res = await api(API_KB_DELETE, { method: "POST", body: JSON.stringify({ path: f.path }) });
          setKnowledge(res.knowledge || knowledge);
          setStages((prev) => prev.map((s) => ({
            ...s,
            kbRefs: (s.kbRefs || []).filter((p) => p !== f.path),
          })));
          setStatus("已删除 " + f.name);
        } catch (e) {
          setStatus("删除失败：" + e.message);
        }
      }

      function addKbRef(stageId, path) {
        const target = stages.find((s) => s.id === stageId);
        if (!target) return;
        const refs = target.kbRefs || [];
        if (refs.includes(path)) {
          setStatus("这个文件已经在本阶段依据里了");
          return;
        }
        patchStage(stageId, { kbRefs: [...refs, path] });
        setKbOpen(false);
        setTab("info");
        setStatus("已加入「" + target.name + "」的知识库参考");
      }

      function addMessage(text) {
        if (!stage) return;
        const msg = { role: "user", time: Date.now(), text };
        patchStage(stage.id, { session: [...stage.session, msg] });
        setStatus("已记入会话（记得保存）");
      }

      function copyBrief() {
        if (!stage) return;
        const lines = [
          `【编剧工坊 · ${stage.name}】`,
          `项目：${project.title || "（未填）"}`,
          `格式：${project.format || "（未填）"}｜类型：${project.genre || "（未填）"}`,
          `受众：${project.audience || "（未填）"}｜预算：${project.budget || "（未填）"}｜基调：${project.tone || "（未填）"}`,
          `一句话梗概：${project.logline || "（未填）"}`,
          `主题：${project.theme || "（未填）"}`,
          "",
          stage.brief ? "【总控简报】\n" + stage.brief : "",
          stage.brief ? "" : "",
          "【本阶段默认预设】",
          stage.preset,
          "",
          "【交付前必须回答】",
          ...(stage.questions || []).map((q, i) => `${i + 1}. ${q}`),
          "",
          "【本阶段依据】",
          ...(stage.basis || []).map((b) => "- " + b),
          ...(stage.kbRefs || []).map((p) => "- 知识库：" + p),
          "",
          "【会话记录】",
          ...stage.session.map((m) => (m.role === "user" ? "我：" : m.role === "system" ? "系统：" : "子智能体：") + m.text),
          "",
          "请按以上设定完成本阶段，并把结果填进「" + stage.output + "」。",
        ].filter((x) => x !== "");
        copyText(lines.join("\n"));
        setStatus("简报已复制，可粘到新会话继续");
      }

      if (loading) {
        return h("div", {
          style: {
            height: "100%", display: "flex", alignItems: "center", justifyContent: "center",
            background: C.paper, color: C.muted, fontFamily: FONT, fontSize: 13,
          },
        }, "正在打开编剧工坊…");
      }

      if (error) {
        return h("div", {
          style: { height: "100%", padding: 30, background: C.paper, color: C.ink, fontFamily: FONT },
        },
          h("div", { style: { fontSize: 15, fontWeight: 700, marginBottom: 10 } }, "编剧工坊没有加载成功"),
          h("div", { style: { fontSize: 12, color: C.red, whiteSpace: "pre-wrap", marginBottom: 14 } }, error),
          h(Btn, { onClick: load }, "重试"));
      }

      const tabBtn = (id, label) => h("button", {
        key: id,
        onClick: () => setTab(id),
        style: {
          background: "transparent", border: "none", cursor: "pointer",
          padding: "10px 2px 8px", marginRight: 18, fontSize: 12.5, fontWeight: tab === id ? 700 : 500,
          color: tab === id ? C.teal : C.muted,
          borderBottom: "2px solid " + (tab === id ? C.teal : "transparent"),
        },
      }, label);

      return h("div", {
        style: {
          display: "flex", flexDirection: "column", width: "100%", height: "100%", minWidth: 0,
          background: C.paper, color: C.ink, fontFamily: FONT, overflow: "hidden",
        },
      },
        // 顶栏
        h("div", {
          style: {
            flex: "0 0 auto", display: "flex", alignItems: "center", gap: 10, minWidth: 0,
            padding: "10px 16px", borderBottom: "1px solid " + C.line, background: C.card,
          },
        },
          h("span", { style: { fontSize: 16, fontWeight: 800, letterSpacing: ".02em" } }, "✎ 编剧工坊"),
          h("span", { style: { fontSize: 11, color: C.muted } }, "四书方法 · 九道工序 · Excel 蓝图 + Word 剧本"),
          h("span", { style: { flex: 1 } }),
          h("span", { style: { fontSize: 10.5, color: C.muted, maxWidth: 340, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, status),
          h(Btn, {
            kind: "ghost", color: C.teal,
            onClick: () => { setKbTarget(null); setKbOpen(true); },
          }, "知识库"),
          h(Btn, { kind: "ghost", color: C.inkSoft, onClick: () => save(false) }, "保存"),
          h(Btn, { onClick: generate, disabled: busy || !demoLoaded }, busy ? "生成中…" : "生成全部产物")),

        // 主体
        h("div", { style: { flex: "1 1 auto", display: "flex", minHeight: 0, minWidth: 0, overflow: "hidden" } },
          h(StageRail, { stages, activeId: stage ? stage.id : "", onPick: setActiveId }),
          h("div", { style: { flex: "1 1 auto", minWidth: 0, display: "flex", flexDirection: "column", background: C.paper, overflow: "hidden" } },
            // 阶段头
            h("div", {
              style: {
                flex: "0 0 auto", padding: "14px 20px 0", borderBottom: "1px solid " + C.lineSoft,
                background: C.card,
              },
            },
              h("div", { style: { display: "flex", alignItems: "baseline", gap: 10, flexWrap: "wrap" } },
                h("span", { style: { fontSize: 18, fontWeight: 800 } }, stage.name),
                h("span", { style: { fontSize: 11.5, color: C.teal, fontWeight: 600 } }, stage.role),
                h("span", { style: { flex: 1 } }),
                h("span", { style: { fontSize: 10.5, color: C.muted } }, stage.output)),
              h("div", { style: { fontSize: 11.5, color: C.muted, marginTop: 4, lineHeight: 1.6 } }, stage.summary),
              h("div", { style: { display: "flex", marginTop: 6 } },
                tabBtn("chat", "会话窗口"),
                tabBtn("preset", "默认预设"),
                tabBtn("info", "阶段说明"))),
            // 内容
            h("div", { style: { flex: "1 1 auto", minHeight: 0 } },
              tab === "chat"
                ? h(ChatPane, {
                  stage,
                  draft: drafts[stage.id] || "",
                  setDraft: (v) => setDrafts((d) => ({ ...d, [stage.id]: v })),
                  onAdd: addMessage,
                  onCopyBrief: copyBrief,
                })
                : tab === "preset"
                  ? h(PresetPane, {
                    stage, defaults,
                    onChange: (v) => patchStage(stage.id, { preset: v }),
                    onReset: () => patchStage(stage.id, { preset: defaults[stage.id] || stage.preset }),
                  })
                  : h(InfoPane, {
                    stage,
                    onPatch: (p) => patchStage(stage.id, p),
                    onOpenKb: (target) => { setKbTarget(target); setKbOpen(true); },
                  })),
          ),
          h(SidePane, {
            project,
            onChange: (k, v) => setProject((p) => ({ ...p, [k]: v })),
            outputs, outDir, root,
            onGenerate: generate, onSync: () => syncProject(false), busy, onRefresh: refresh,
            onDemo: loadDemo, onNew: newProject,
            demoLoaded, canBuild: demoLoaded,
          })),
        // 底栏
        h("div", {
          style: {
            flex: "0 0 auto", display: "flex", alignItems: "center", gap: 12,
            padding: "5px 16px", borderTop: "1px solid " + C.line,
            background: C.card2, fontSize: 10, color: C.muted,
          },
        },
          h("span", null, "项目：" + (project.title || "空白")),
          h("span", null, "阶段 " + (stage ? stage.no : "-") + "/9"),
          h("span", { style: { flex: 1 } }),
          h("span", null, "预设、会话、阶段说明保存在 plugin-state.json；产物在 产出/")),
        h(CheckModal, {
          data: checkModal,
          busy,
          onClose: () => setCheckModal(null),
          onForce: () => { setCheckModal(null); syncProject(true); },
        }),
        h(KnowledgeModal, {
          open: kbOpen,
          files: knowledge,
          targetStage: kbTarget,
          uploading,
          onClose: () => { setKbOpen(false); setKbTarget(null); },
          onUpload: uploadKnowledge,
          onDelete: deleteKnowledge,
          onAdd: addKbRef,
        }));
    }

    // ------------------------------------------------------------------ 侧栏图标
    function PanelIcon() {
      return h("span", {
        style: { display: "inline-flex", alignItems: "center", justifyContent: "center" },
        title: "编剧工坊",
      },
        h("svg", {
          width: 20, height: 20, viewBox: "0 0 24 24",
          fill: "none", stroke: "currentColor", strokeWidth: 1.6,
        },
          h("path", { d: "M4 19.5V6.2L7.5 4 11 6.2V19.5", strokeLinejoin: "round" }),
          h("path", { d: "M11 19.5V9.8L14.5 7.6 18 9.8v9.7", strokeLinejoin: "round" }),
          h("path", { d: "M3 19.5h18" }),
          h("path", { d: "M5.8 8.2h1.8M5.8 11.4h1.8M5.8 14.6h1.8M13.3 11.8h1.8M13.3 15h1.8" })));
    }

    // ------------------------------------------------------------------ 入口
    const inject = ["slots"];

    function apply(ctx) {
      ctx.slots.inject("sidebar.panellist", () => ctx.slots.register({
        name: "sidebar.panellist",
        id: "screenwriter",
        order: 2,
        label: "编剧工坊",
      }, PanelIcon));
      ctx.slots.inject("main", () => ctx.slots.register({
        name: "main",
        key: "screenwriter",
      }, ScreenwriterPanel));
    }

    exports.apply = apply;
    exports.inject = inject;
    return module.exports;
  },
});
