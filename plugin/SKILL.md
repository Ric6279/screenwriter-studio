# 编剧工坊

用四本编剧书当方法、九道工序当流程，把一句想法做成可拍摄的剧作蓝图（Excel）与完整剧本（Word）。

## 四本书怎么用

- **《故事》麦基**：管“为什么”。控制性理念、激励事件、进展纠葛、危机/高潮/结局、期望—结果—鸿沟、场景与潜文本、伏笔与回馈。
- **《救猫咪》斯奈德**：管“什么时候”。十类型、一句话梗概四要素、15 节拍与百分比、故事板、救猫咪时刻。
- **《电影剧本写作基础》菲尔德**：管“怎么搭”。三幕 25/50/25、情节点Ⅰ（25%）与Ⅱ（75%）、场景晚进早出、剧本格式、一页≈一分钟。
- **《打草稿》工作手册**：管“填什么表”。概念表、角色表、关系表、情节点表、弧光表、情节表、一幕接一幕。

完整方法提要见工作区 `知识库/` 下五份 Markdown；默认预设见插件 `presets/stages.json`。

## 九道工序

1. **类型选择** → `01 类型与卖点`
2. **三幕八序列大纲** → `03 三幕八序列`
3. **主题撰写** → 控制性理念 / 反理念 / 三个主题场
4. **详细故事小节** → `04 十五节拍`
5. **情节点** → `05 情节点` + `05b 伏笔回馈`
6. **主要人物设计** → `06 人物设计` + `07 人物关系`
7. **主要场景设计** → `08 主要场景`
8. **分场 / 分集大纲** → `09 分集大纲` + `10 第01集分场`
9. **完整剧本** → `02_第01集_完整剧本.docx`

每一道的默认预设都写明：角色、输入、方法（引用书与章节）、输出格式、自检。不要跳过“表”，直接写剧本。

## 工具

- `screenwriter_stages`：列出九个阶段与默认预设；写任何阶段前先调它。
- `screenwriter_build_xlsx`：项目数据 JSON → 《剧作蓝图》xlsx（14 张可见表 + 情绪曲线）。
- `screenwriter_build_docx`：剧本 JSON → Word 剧本（封面、人物表、正文、页码）。
- `screenwriter_build_all`：用工作台当前项目信息一键重建两份产物（空白项目会被拦下）。
- `screenwriter_demo`：载入《回声镇》示范案例（项目卡 + 九阶段会话）。
- `screenwriter_new_project`：重置为空白项目。

默认路径：
- 项目数据：`编剧工坊/demo/项目数据.json`
- 剧本数据：`编剧工坊/demo/剧本_第01集.json`
- 产物：`编剧工坊/产出/`

## 项目数据 JSON 的关键字段

```
project: { title, subtitle, format, genre, genre_cat, logline, theme, controlling_idea, audience, budget }
genre_analysis: { primary, secondary, why_primary, why_secondary, audience_hook, platform_fit, rules[] }
acts[8]: { act, seq, name, episodes, timecode, core_event, turn, function }
beats[15]: { no, beat, pct, episode, event, why, page }
plot_points[5]: { no, name, position, page, content, effect }
characters[9]: { name, age, identity, role, desire, need, flaw, secret, arc, relation, voice }
relationships: { codes, edges[], matrix[] }
scenes[10]: { name, int_ext, time, atmosphere, motif, sound, function, episodes }
episodes[12]: { no, title, core, conflict, hook, emotion, seq, duration }
ep1_scenes[14]: { no, loc, int_ext, time, chars, summary, emotion, duration }
tension_curve[12]: { ep, value }
```

剧本 JSON：`{title, episode, episode_title, format, draft, date, logline, characters[], scenes[{heading, elements[{type,text}]}]}`；`type` 取 `action / character / parenthetical / dialogue / transition`。

## 工作台

侧栏「编剧工坊」打开九宫格：左侧选阶段，中间编辑该阶段的**默认预设**、**会话窗口**与**阶段说明**，右侧看项目卡与产物。

- 默认进入是空白项目；点「载入示范案例」看《回声镇》怎么跑。
- 项目卡填好后点「总控」：先本地自检（缺项、长度、类型/基调/受众/预算冲突），有问题会弹窗；继续后把项目卡写成九份总控简报，写入各阶段会话。
- 「阶段说明」里的自检问题、本阶段依据、知识库参考都可以逐条增删改。
- 顶栏「知识库」可上传剧本/编剧书（`.md/.txt/.docx/.pdf`），自动抽取文本，挂到某阶段作为参考。
- 改完点「保存」，点「生成」重建 Excel/Word（需要示范数据或已完成的结构化内容）。

## 自检

交付前逐项过 `11 结构自检`：激励事件、场景鸿沟、控制性理念、危机两难、开场/终场镜像、主题陈述、中点、救猫咪时刻、情节点 25%/75%、结尾回答开头、类型/受众/预算自洽、人物四要素、关系潜台词、分场可追溯。
