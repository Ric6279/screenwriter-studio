# @fatfish/screenwriter-studio（编剧工坊）

DSH 插件包。侧栏多一个「✎ 编剧工坊」，打开是九宫格工作台；同时在任意会话里注册
`screenwriter_*` 工具与 `screenwriter-studio` 技能。

## 结构

```
plugin/
├─ package.json          # @fatfish/screenwriter-studio
├─ cordis.patch.yml      # 把自己插为一行
├─ icon.svg
├─ SKILL.md              # 注册给模型的技能说明
├─ presets/stages.json   # 九道工序 + 默认预设 + 自检问题
├─ lib/
│  ├─ host.js            # 宿主半：工具 / 技能 / /screenwriter/* 路由
│  └─ client.js          # 浏览器半：九宫格工作台
└─ engine/
   ├─ build_xlsx.py      # 项目 JSON → 剧作蓝图 xlsx
   ├─ build_docx.py      # 剧本 JSON → Word 剧本
   ├─ ingest_knowledge.py# 上传的 docx/pdf/txt/md → 摘录
   └─ ocr_pdf.py         # PDF 取字：文字层优先，扫描页走 RapidOCR
```

## 宿主半

注入 `tools`、`skills`、`webServer`、`connection`。

工具：
- `screenwriter_stages`：九阶段与默认预设。
- `screenwriter_build_xlsx`：项目数据 → xlsx。
- `screenwriter_build_docx`：剧本数据 → docx。
- `screenwriter_build_all`：按工作台已保存的项目信息重建两份产物（空白项目会被拦下）。
- `screenwriter_demo`：载入《回声镇》示范案例。
- `screenwriter_new_project`：重置为空白项目。

路由（全部过 `connection.requestRejection` 信任闸）：
- `GET/POST /screenwriter/state`：项目卡、九阶段预设/会话/阶段说明、产物与知识库列表。
- `POST /screenwriter/build`：`{target:"xlsx"|"docx"|"all"}`，串行调用 Python。
- `GET /screenwriter/files`：产物与知识库。
- `POST /screenwriter/demo` / `project/new`：载入示范案例 / 清空为空白项目。
- `POST /screenwriter/project/check` / `project/sync`：项目卡自检 / 总控。
- `POST /screenwriter/knowledge/upload` / `delete`：上传/删除知识库文件。

状态写在 `编剧工坊/plugin-state.json`；生成时会把当前项目信息合并进
`demo/_当前项目.json` 与 `demo/_当前剧本.json`，再交给引擎，避免直接改动原始 demo 数据。

## 浏览器半

静态客户端包要求预构建的 CJS 包体（`window.__ModuleLoader__.load`），只用
`require("react")` 与浏览器原生 API。注册两个槽位：
- `sidebar.panellist`（id `screenwriter`，order 2）
- `main`（key `screenwriter`）

工作台（保持三栏布局）：
- 初始空白：项目卡空、九个会话空；《回声镇》通过右栏「载入示范案例」作为示范预设加载。
- 左：九道工序（开发 / 结构 / 人物 / 成稿）。
- 中：会话窗口、默认预设（可改可重置）、阶段说明。阶段说明里的
  「交付前必须回答 / 本阶段依据 / 知识库参考」三块都可逐条编辑、删除、新增，
  保存后写入 `plugin-state.json`。
- 右：项目卡（名称/格式/类型/受众/预算/基调/梗概/主题）、「总控」自检与九阶段简报、
  生成按钮、交付物、流程。
- 顶栏「知识库」：上传、删除、管理知识库文件，并可把文件挂到当前阶段的参考里。
  文字版 PDF 直接读文字层；扫描版 PDF 自动走本机 RapidOCR（不联网、不上传），
  默认最多识别 120 页。

OCR 复用 `course-kit/.pylibs` 里的 pymupdf + RapidOCR，不额外安装一套。可用环境变量覆盖：

| 变量 | 作用 |
|---|---|
| `SCREENWRITER_OCR_LIBS` | OCR 依赖目录（默认 `<工作区>/course-kit/.pylibs`） |
| `SCREENWRITER_OCR_PYTHON` | 跑 OCR 的 Python（该依赖按 3.10 安装；默认找 `Python310/python.exe`） |
| `SCREENWRITER_OCR_PAGES` | 单次上传最多 OCR 多少页，默认 120；0 表示整本 |

## 引擎

`build_xlsx.py` 生成 14 张可见表：封面、类型与卖点、一句话梗概、三幕八序列、十五节拍、
情节点、伏笔回馈、人物设计、人物关系矩阵、关系清单、主要场景、分集大纲（含情绪曲线）、
第 1 集分场、结构自检。排版按《打草稿》的“一页一张表”。

`build_docx.py` 按菲尔德格式排版：封面、人物表、场次标题、动作、人物名、对白、转场、
页眉页脚；中文宋体/微软雅黑。

## 安装与刷新

```powershell
# 安装（在 DSH 会话里用 plugin_manager，或命令行）
# plugin_manager action=install_bundle target=<此目录绝对路径>
```

- 宿主半改动需要重启 DSH 才会重载。
- 浏览器半改动：`install_bundle` 会重新构建；刷新页面可见。
- 引擎脚本是外部进程，改完立即生效（下一次点“生成”）。

## 已知边界

- 工作台不直接调用模型：会话窗口负责“记要求、存预设、复制简报”，真正的创作在会话里完成；
  九阶段的默认预设就是给模型/子智能体的 brief。
- 「总控」只做项目卡自检与九阶段 brief 广播，不直接生成正文；生成 Excel/Word 仍要有
  结构化数据（示范案例或导入 JSON）。
- 数据根默认 `screenwriter-tools/编剧工坊`，迁移时用 `SCREENWRITER_HOME` 覆盖。
- 生成是确定性的：Excel/Word 的结构与格式由脚本保证，创意内容来自 `demo/*.json`。
