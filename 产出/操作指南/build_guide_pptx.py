# -*- coding: utf-8 -*-
"""生成《编剧工坊·三块工作区操作指南》PPT。"""

from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from pptx.oxml.xmlchemy import OxmlElement

ROOT = Path(r"E:\desktop\肥鱼专属空间\screenwriter-tools\编剧工坊\产出\操作指南")
ASSETS = ROOT / "assets"
OUT = ROOT / "编剧工坊_三块工作区操作指南.pptx"

PAPER = RGBColor(0xF5, 0xF0, 0xE5)
CARD = RGBColor(0xFF, 0xFD, 0xF7)
INK = RGBColor(0x1B, 0x2A, 0x38)
INK_SOFT = RGBColor(0x3E, 0x4C, 0x59)
MUTED = RGBColor(0x78, 0x6F, 0x62)
LINE = RGBColor(0xDC, 0xD2, 0xBE)
TEAL = RGBColor(0x1F, 0x6F, 0x78)
TEAL_SOFT = RGBColor(0xDC, 0xEA, 0xEA)
AMBER = RGBColor(0xB0, 0x7D, 0x2B)
AMBER_SOFT = RGBColor(0xF2, 0xE4, 0xC8)
DARK = RGBColor(0x14, 0x23, 0x31)
RED = RGBColor(0x9B, 0x2C, 0x2C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Microsoft YaHei"
SERIF = "SimSun"

SW, SH = 13.333, 7.5


def set_font(run, size=12, bold=False, color=INK, name=FONT, italic=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    f.name = name
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        node = rpr.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            rpr.append(node)
        node.set("typeface", name)


def text(slide, value, x, y, w, h, size=12, color=INK, bold=False,
         font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         line=1.25, space_after=0, margins=(0, 0, 0, 0), italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom = [Inches(m) for m in margins]
    tf.vertical_anchor = anchor
    tf.clear()
    first = True
    for line_text in str(value).split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = line_text
        set_font(r, size=size, bold=bold, color=color, name=font, italic=italic)
    return box


def rect(slide, x, y, w, h, fill=CARD, line_color=None, radius=0.04, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(0.75)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    shp.text_frame.text = ""
    return shp


def line(slide, x1, y1, x2, y2, color=LINE, width=0.75):
    shp = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    shp.line.color.rgb = color
    shp.line.width = Pt(width)


def badge(slide, number, x, y, d=0.34, fill=TEAL, color=WHITE, size=10):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(number)
    set_font(r, size=size, bold=True, color=color)


def header(slide, kicker, title, page):
    line(slide, 0.55, 0.34, 0.78, 0.34, color=AMBER, width=2.5)
    text(slide, kicker, 0.9, 0.22, 8.0, 0.3, size=10, color=MUTED, bold=True)
    text(slide, title, 0.55, 0.58, 12.2, 0.62, size=25, color=INK, bold=True)
    # footer
    line(slide, 0.55, 7.08, 12.78, 7.08, color=LINE, width=0.6)
    text(slide, "编剧工坊 · 三块工作区操作指南", 0.55, 7.15, 6.0, 0.22, size=8.5, color=MUTED)
    text(slide, f"{page:02d}", 12.25, 7.13, 0.5, 0.25, size=9, color=MUTED, align=PP_ALIGN.RIGHT, bold=True)


def picture(slide, name, x, y, w, h):
    """在给定框内等比放图并居中，加细边框。"""
    path = ASSETS / name
    iw, ih = Image.open(path).size
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    px, py = x + (w - pw) / 2, y + (h - ph) / 2
    pic = slide.shapes.add_picture(str(path), Inches(px), Inches(py), Inches(pw), Inches(ph))
    pic.line.color.rgb = LINE
    pic.line.width = Pt(0.75)
    return pic


def notes(slide, value):
    slide.notes_slide.notes_text_frame.text = value


prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
blank = prs.slide_layouts[6]

# ---------------------------------------------------------------- 01 封面
s = prs.slides.add_slide(blank)
bg = s.background.fill
bg.solid()
bg.fore_color.rgb = PAPER
rect(s, 0, 0, 4.85, SH, fill=DARK, line_color=None, shape=MSO_SHAPE.RECTANGLE)
# 装饰圆环
for cx, cy, d in [(3.4, 5.1, 3.2), (4.2, 6.6, 4.4)]:
    o = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - d / 2), Inches(cy - d / 2), Inches(d), Inches(d))
    o.fill.background()
    o.line.color.rgb = RGBColor(0x2C, 0x42, 0x51)
    o.line.width = Pt(0.9)

text(s, "SCREENWRITER STUDIO", 0.55, 0.72, 3.9, 0.3, size=9.5, color=RGBColor(0xB9, 0xAE, 0x97), bold=True)
text(s, "编\n剧\n工\n坊", 0.55, 1.18, 1.3, 2.5, size=36, color=WHITE, bold=True, line=1.0)
line(s, 1.85, 1.32, 1.85, 3.55, color=AMBER, width=2.2)
text(s, "四书方法 · 九道工序", 2.05, 1.35, 2.4, 0.35, size=13, color=RGBColor(0xE8, 0xDF, 0xCB), bold=True)
text(s, "Excel 蓝图 + Word 剧本", 2.05, 1.78, 2.4, 0.32, size=11, color=RGBColor(0xB9, 0xAE, 0x97))
text(s, "把一句想法，做成可开拍的剧作蓝图。", 0.55, 4.25, 3.6, 0.7, size=11.5, color=RGBColor(0xD8, 0xCF, 0xBD), line=1.6)

# 九个小格子
for i in range(9):
    row, col = divmod(i, 3)
    x, y = 0.55 + col * 1.12, 5.35 + row * 0.45
    r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(0.95), Inches(0.3))
    r.fill.solid()
    r.fill.fore_color.rgb = RGBColor(0x21, 0x36, 0x46)
    r.line.fill.background()
    text(s, f"{i + 1:02d}", x + 0.08, y + 0.04, 0.8, 0.22, size=7.5, color=RGBColor(0x9D, 0xC9, 0xCB), bold=True)

text(s, "操作指南 / 讲解版", 6.05, 0.95, 4.0, 0.3, size=11, color=TEAL, bold=True)
text(s, "三块工作区\n一套从空白到成稿的流程", 5.95, 1.35, 6.8, 1.55, size=30, color=INK, bold=True, line=1.25)
text(s, "左选工序 · 中做内容 · 右管项目", 6.0, 3.08, 6.4, 0.36, size=14, color=AMBER, bold=True)
text(s, "给第一次使用编剧工坊的同事：先认三块区域，再用五步串起一遍流程。",
     6.0, 3.62, 6.3, 0.8, size=12, color=INK_SOFT, line=1.7)
for i, (label, color) in enumerate([("左 · 选工序", TEAL), ("中 · 做内容", AMBER), ("右 · 管项目", INK)]):
    x = 6.0 + i * 2.12
    chip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(5.0), Inches(1.95), Inches(0.5))
    chip.fill.solid(); chip.fill.fore_color.rgb = color; chip.line.fill.background()
    tf = chip.text_frame; tf.margin_left = tf.margin_right = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label; set_font(r, size=10.5, bold=True, color=WHITE)

text(s, "2026 · 编剧工坊", 6.0, 6.55, 4.0, 0.3, size=9, color=MUTED)
notes(s, "开场先说明：这不是一份功能清单，而是一条使用路径。今天只讲三块区域和五个动作，讲完大家能自己从空白项目走到生成 Excel/Word。")

# ---------------------------------------------------------------- 02 全貌
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "01 / 先认界面", "三块区域：左选工序，中做内容，右管项目", 2)
picture(s, "ui_full.png", 0.5, 1.42, 8.45, 5.15)
text(s, "实际界面：空白项目状态", 0.55, 6.65, 4.0, 0.28, size=9, color=MUTED, italic=True)

cards = [
    ("01", "左侧 · 九道工序", "四个阶段、九个工序。点哪一项，中栏就切到哪一阶段。", TEAL),
    ("02", "中间 · 三页工作区", "会话窗口、默认预设、阶段说明。真正写要求、改 brief、挂资料都在这里。", AMBER),
    ("03", "右侧 · 项目卡", "填项目信息，点「总控」把设定同步给九个阶段；下面看交付物。", INK),
]
for i, (num, title, desc, color) in enumerate(cards):
    y = 1.45 + i * 1.72
    rect(s, 9.2, y, 3.6, 1.45, fill=CARD, line_color=LINE)
    badge(s, num, 9.38, y + 0.2, d=0.38, fill=color)
    text(s, title, 9.9, y + 0.18, 2.7, 0.3, size=13, color=INK, bold=True)
    text(s, desc, 9.38, y + 0.58, 3.15, 0.75, size=9.8, color=INK_SOFT, line=1.5)
text(s, "讲解时先指这三块，再按后面的 01 → 05 顺序走。", 9.2, 6.72, 3.7, 0.3, size=10, color=TEAL, bold=True)
notes(s, "先让大家认位置：左中右不是三个功能，而是三种角色——流程、内容、项目。后面所有操作都落在这三块里。")

# ---------------------------------------------------------------- 03 项目卡 + 总控
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "02 / 右栏 · 项目卡", "先把项目说清楚，再点总控", 3)
text(s, "项目卡是整条流程的输入。项目名、格式、类型、受众、预算、基调六项会参与自检；梗概与主题会写进九份简报。",
     0.55, 1.35, 7.9, 0.65, size=11.5, color=INK_SOFT, line=1.65)

steps = [
    ("填项目卡", "右侧从上到下填八项：项目名、格式、类型、受众、预算、基调、一句话梗概、主题。"),
    ("点「总控」", "按钮会先做本地自检：缺项、长度、类型/梗概/基调冲突、受众、预算。"),
    ("处理自检结果", "有冲突会弹窗。可以「返回修改」，也可以确认无误后「仍然继续」。"),
    ("接收九份简报", "继续后，项目卡被写成九份总控简报，写进九个阶段的会话窗口。"),
]
y = 2.2
for i, (t, d) in enumerate(steps):
    badge(s, f"{i + 1:02d}", 0.6, y + 0.04, d=0.42, fill=TEAL if i < 3 else AMBER, size=9)
    text(s, t, 1.18, y - 0.02, 3.0, 0.3, size=13, color=INK, bold=True)
    text(s, d, 1.18, y + 0.34, 6.9, 0.62, size=10.3, color=INK_SOFT, line=1.55)
    y += 0.98

# 模拟项目卡
rect(s, 9.0, 1.35, 3.85, 5.05, fill=CARD, line_color=LINE)
text(s, "项目卡", 9.22, 1.53, 1.4, 0.28, size=10.5, color=MUTED, bold=True)
text(s, "还没有项目", 9.22, 1.82, 2.2, 0.3, size=13, color=INK, bold=True)
text(s, "填完项目卡，或先载入示范案例看完整流程。", 9.22, 2.15, 3.35, 0.55, size=9.2, color=MUTED, line=1.5)
btn = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.22), Inches(2.68), Inches(1.35), Inches(0.36))
btn.fill.solid(); btn.fill.fore_color.rgb = TEAL; btn.line.fill.background()
tf = btn.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "载入示范案例"; set_font(r, size=8.8, bold=True, color=WHITE)
fields = [("项目名", ""), ("格式", ""), ("类型", ""), ("受众", ""), ("预算", ""), ("基调", "")]
for i, (label, _) in enumerate(fields):
    yy = 3.18 + i * 0.4
    text(s, label, 9.22, yy, 0.7, 0.2, size=7.8, color=MUTED)
    rect(s, 9.92, yy - 0.03, 2.65, 0.25, fill=WHITE, line_color=LINE, radius=0.02)
master = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.22), Inches(5.5), Inches(3.35), Inches(0.46))
master.fill.solid(); master.fill.fore_color.rgb = AMBER; master.line.fill.background()
tf = master.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "总控 · 按项目卡推进九阶段"; set_font(r, size=10.2, bold=True, color=WHITE)
text(s, "总控只广播 brief，不直接写剧本正文。", 9.22, 6.02, 3.35, 0.25, size=8.2, color=MUTED)
notes(s, "强调项目卡不是表单摆设：总控会拿它做自检并生成九份简报。先填清楚，后面每个阶段才有共同前提。")

# ---------------------------------------------------------------- 04 知识库
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "03 / 顶栏 · 知识库", "上传一次，挂到对应阶段", 4)
text(s, "顶栏「知识库」是资料入口。上传的编剧书、剧本、案例会进入知识库；扫描版 PDF 会在本机自动 OCR。",
     0.55, 1.32, 8.2, 0.55, size=11.5, color=INK_SOFT, line=1.6)

flow = [
    ("上传", "顶栏「知识库」→ 上传 .md / .txt / .docx / .pdf，单文件 ≤ 30 MB。"),
    ("取字", "文字版直接读文字层；扫描版逐页 OCR（默认最多 120 页，本地运行）。"),
    ("挂载", "到「阶段说明」→ 知识库参考 →「+ 从知识库添加」，选当前阶段要用的文件。"),
]
for i, (t, d) in enumerate(flow):
    x = 0.55 + i * 4.15
    rect(s, x, 2.05, 3.85, 1.38, fill=CARD, line_color=LINE)
    badge(s, f"{i + 1}", x + 0.2, 2.22, d=0.34, fill=TEAL, size=9)
    text(s, t, x + 0.65, 2.2, 2.8, 0.3, size=13, color=INK, bold=True)
    text(s, d, x + 0.2, 2.62, 3.45, 0.68, size=9.5, color=INK_SOFT, line=1.55)

# 左：知识库文件
rect(s, 0.55, 3.6, 4.0, 2.6, fill=CARD, line_color=LINE)
text(s, "知识库", 0.78, 3.76, 1.8, 0.28, size=10, color=MUTED, bold=True)
text(s, "已上传 3 个文件", 0.78, 4.02, 3.3, 0.28, size=12.5, color=INK, bold=True)
files = [
    ("方法", "01_故事_麦基.md", "6.4 KB", TEAL),
    ("上传", "救猫咪.pdf", "8.8 MB", AMBER),
    ("OCR", "打草稿_p9-192.ocr.txt", "146 KB", RED),
]
for i, (tag, name, size, color) in enumerate(files):
    yy = 4.45 + i * 0.52
    chip = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.78), Inches(yy + 0.04), Inches(0.52), Inches(0.24))
    chip.fill.solid(); chip.fill.fore_color.rgb = color; chip.line.fill.background()
    tf = chip.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = tag; set_font(r, size=7.2, bold=True, color=WHITE)
    text(s, name, 1.4, yy, 2.35, 0.28, size=9.2, color=INK_SOFT, bold=True)
    text(s, size, 3.75, yy + 0.03, 0.65, 0.22, size=7.5, color=MUTED, align=PP_ALIGN.RIGHT)

# 中间箭头
arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.62), Inches(4.72), Inches(0.32), Inches(0.34))
arrow.fill.solid(); arrow.fill.fore_color.rgb = AMBER; arrow.line.fill.background()

# 右：阶段说明里的挂载
rect(s, 4.98, 3.6, 3.3, 2.6, fill=CARD, line_color=LINE)
text(s, "阶段说明 · 类型选择", 5.18, 3.76, 2.8, 0.28, size=9.6, color=MUTED, bold=True)
text(s, "三块内容", 5.18, 4.03, 2.6, 0.28, size=12.5, color=INK, bold=True)
rows = [("交付前必须回答", "04 条"), ("本阶段依据", "03 条"), ("知识库参考", "01 个文件")]
for i, (label, meta) in enumerate(rows):
    yy = 4.44 + i * 0.47
    text(s, label, 5.18, yy, 2.1, 0.26, size=9.2, color=INK_SOFT)
    text(s, meta, 7.35, yy, 0.65, 0.24, size=8, color=MUTED, align=PP_ALIGN.RIGHT)
    if i < 2:
        line(s, 5.18, yy + 0.32, 7.95, yy + 0.32, color=LINE, width=0.5)
ref = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.18), Inches(5.74), Inches(2.85), Inches(0.3))
ref.fill.solid(); ref.fill.fore_color.rgb = TEAL_SOFT; ref.line.color.rgb = RGBColor(0xC9, 0xDE, 0xDC)
tf = ref.text_frame; tf.margin_left = Inches(0.08); p = tf.paragraphs[0]
r = p.add_run(); r.text = "· 02_救猫咪_斯奈德.md    ×"; set_font(r, size=7.8, color=TEAL)
text(s, "挂载后才会进入该阶段简报", 0.55, 6.3, 7.6, 0.22, size=8.6, color=MUTED, italic=True)

rect(s, 8.45, 3.65, 4.35, 2.55, fill=AMBER_SOFT, line_color=RGBColor(0xE3, 0xCF, 0xA5))
text(s, "讲解时提醒三件事", 8.68, 3.82, 3.8, 0.3, size=12, color=AMBER, bold=True)
for i, t in enumerate([
    "OCR 在本机跑，不上传、不联网。",
    "页数多的扫描件上传会久一点，按钮显示「OCR 中…」。",
    "上传完不等于已使用：还要在每个阶段里挂载。",
]):
    text(s, f"·  {t}", 8.68, 4.25 + i * 0.55, 3.85, 0.48, size=9.4, color=INK_SOFT, line=1.45)

notes(s, "知识库是第二步。关键不是上传，而是挂载：哪个阶段用哪本书的哪一章，要进到该阶段的阶段说明里选。")

# ---------------------------------------------------------------- 05 左侧工序
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "04 / 左栏 · 九道工序", "按编号走，就是交付顺序", 5)
text(s, "四个阶段、九个工序。点任意一项，中栏切到该阶段；编号 01—09 是推荐顺序，不是必须死守。",
     0.55, 1.32, 9.3, 0.45, size=11.5, color=INK_SOFT)

phases = [
    ("开发", "把类型和主题定下来", [
        ("01", "类型选择", "主类型 + 副类型 + 观众买点"),
        ("03", "主题撰写", "控制性理念 / 反理念 / 主题场"),
    ], TEAL),
    ("结构", "搭骨架，排节拍", [
        ("02", "三幕八序列大纲", "8 个序列，每序列一个转折"),
        ("04", "详细故事小节", "15 节拍，把每拍写成事件"),
        ("05", "情节点", "情节点 Ⅰ/Ⅱ、伏笔与回馈"),
    ], AMBER),
    ("人物", "人物与场景落地", [
        ("06", "主要人物设计", "欲望 / 缺陷 / 秘密 / 弧光"),
        ("07", "主要场景设计", "场景卡：母题、声音、功能"),
    ], RGBColor(0x5B, 0x6B, 0x4E)),
    ("成稿", "分场分集与剧本", [
        ("08", "分场 / 分集大纲", "12 集分场，钩子与情绪曲线"),
        ("09", "完整剧本", "按菲尔德格式写第 1 集"),
    ], RED),
]
for i, (name, sub, items, color) in enumerate(phases):
    x = 0.55 + i * 3.16
    rect(s, x, 1.95, 2.95, 4.6, fill=CARD, line_color=LINE)
    top = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.95), Inches(2.95), Inches(0.82))
    top.fill.solid(); top.fill.fore_color.rgb = color; top.line.fill.background()
    tf = top.text_frame
    tf.margin_left = Inches(0.18); tf.margin_top = Inches(0.1)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = name; set_font(r, size=15, bold=True, color=WHITE)
    p2 = tf.add_paragraph(); r2 = p2.add_run(); r2.text = sub; set_font(r2, size=8.8, color=WHITE)
    for j, (num, label, desc) in enumerate(items):
        yy = 2.95 + j * 1.08
        badge(s, num, x + 0.2, yy + 0.02, d=0.34, fill=color, size=7.5)
        text(s, label, x + 0.65, yy - 0.03, 2.15, 0.28, size=10.8, color=INK, bold=True, line=1.3)
        text(s, desc, x + 0.65, yy + 0.31, 2.15, 0.5, size=8.2, color=MUTED, line=1.4)
        if j < len(items) - 1:
            line(s, x + 0.2, yy + 0.89, x + 2.75, yy + 0.89, color=LINE, width=0.5)
text(s, "讲解提示：先选阶段，再去中栏。不要跳阶段直接写剧本。", 0.55, 6.72, 8.0, 0.26, size=9.5, color=TEAL, bold=True)
notes(s, "左侧是导航，也是方法。四书的方法被落成九个工序：先类型、再结构、再人物场景、最后成稿。")

# ---------------------------------------------------------------- 06 会话 & 预设
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "05 / 中栏 · 会话与预设", "会话窗口 & 默认预设", 6)
# 左卡
rect(s, 0.55, 1.4, 6.08, 5.25, fill=CARD, line_color=LINE)
text(s, "会话窗口", 0.8, 1.58, 3.0, 0.35, size=15, color=TEAL, bold=True)
text(s, "写补充要求、看总控简报、复制给子智能体", 0.8, 1.92, 5.5, 0.3, size=9.5, color=MUTED)
# 会话窗口 mock
rect(s, 0.78, 2.25, 5.6, 2.08, fill=WHITE, line_color=LINE, radius=0.03)
text(s, "类型选择 · 会话窗口", 0.95, 2.36, 3.0, 0.24, size=8.2, color=MUTED, bold=True)
rect(s, 0.95, 2.68, 4.2, 0.48, fill=AMBER_SOFT, line_color=None, radius=0.08)
text(s, "总控简报：主类型=社会派悬疑，副类型=家庭伦理；先把类型承诺说清楚。",
     1.05, 2.75, 3.98, 0.36, size=8.1, color=INK_SOFT, line=1.35)
rect(s, 2.55, 3.24, 3.55, 0.42, fill=TEAL_SOFT, line_color=None, radius=0.08)
text(s, "第 2 集结尾先不揭示磁带来源，只给声音。", 2.67, 3.31, 3.3, 0.28, size=8.2, color=INK_SOFT)
rect(s, 0.95, 3.82, 4.15, 0.34, fill=PAPER, line_color=LINE, radius=0.04)
text(s, "改哪一拍 / 换哪个人物 / 加哪场戏……", 1.05, 3.88, 3.9, 0.22, size=7.8, color=MUTED)
btn = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(3.82), Inches(0.98), Inches(0.34))
btn.fill.solid(); btn.fill.fore_color.rgb = TEAL; btn.line.fill.background()
tf = btn.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "记入会话"; set_font(r, size=7.6, bold=True, color=WHITE)
for i, t in enumerate([
    "空白项目显示引导；总控后顶部出现该阶段简报。",
    "在底部输入「改哪一拍 / 换哪个人物 / 加哪场戏」，点「记入会话」。",
    "「复制简报」把项目卡 + 预设 + 会话打包，粘到子智能体会话继续。",
]):
    text(s, f"{i + 1}.  {t}", 0.82, 4.62 + i * 0.57, 5.55, 0.48, size=9.2, color=INK_SOFT, line=1.42)

# 右卡
rect(s, 6.85, 1.4, 5.9, 5.25, fill=CARD, line_color=LINE)
text(s, "默认预设", 7.1, 1.58, 3.0, 0.35, size=15, color=AMBER, bold=True)
text(s, "这一页就是交给子智能体的 brief", 7.1, 1.92, 5.4, 0.3, size=9.5, color=MUTED)
# 预设 mock
rect(s, 7.08, 2.25, 5.42, 2.08, fill=WHITE, line_color=LINE, radius=0.03)
text(s, "类型策划 / 默认预设", 7.25, 2.36, 3.2, 0.24, size=8.2, color=MUTED, bold=True)
small = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.18), Inches(2.32), Inches(1.08), Inches(0.26))
small.fill.solid(); small.fill.fore_color.rgb = PAPER; small.line.color.rgb = LINE
tf = small.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "重置为默认"; set_font(r, size=6.8, color=MUTED)
preset_lines = [
    ("你是「类型策划」，只做一件事：为这个项目定类型。", INK_SOFT, True),
    ("【方法】① 十类型比对  ② 选主类型  ③ 查类型合同", INK_SOFT, False),
    ("【输出】七列表格 + 4 条规则 + 3 条卖点 + 3 部对标", INK_SOFT, False),
    ("【自检】主类型唯一？副类型只加一个？预算匹配？", MUTED, False),
]
for i, (t, c, b) in enumerate(preset_lines):
    text(s, t, 7.25, 2.68 + i * 0.34, 4.95, 0.28, size=7.9, color=c, bold=b)
for i, t in enumerate([
    "预设写清角色、输入、方法、输出格式、自检问题。",
    "可以直接改；改完点「保存」写入当前阶段。",
    "「重置为默认」恢复 presets/stages.json 的原始预设。",
]):
    text(s, f"{i + 1}.  {t}", 7.12, 4.62 + i * 0.57, 5.35, 0.48, size=9.2, color=INK_SOFT, line=1.42)
notes(s, "会话窗口负责沟通，默认预设负责把方法固定下来。两者配合：预设给底线，会话补当前项目的具体要求。")

# ---------------------------------------------------------------- 07 阶段说明
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "06 / 中栏 · 阶段说明", "三块内容都能单独改", 7)
picture(s, "ui_info.png", 0.45, 1.4, 7.85, 5.4)
text(s, "阶段说明：问题、依据、知识库参考", 0.55, 6.83, 6.5, 0.22, size=8.8, color=MUTED, italic=True)

items = [
    ("01", "交付前必须回答", "每一项都是输入框。直接改文字；右侧 × 删除；底部「+ 添加一问」新增。", TEAL),
    ("02", "本阶段依据", "书、章节、方法、案例。也可以逐条改、删、加，不必只用默认三本。", AMBER),
    ("03", "知识库参考", "点「+ 从知识库添加」，把上传的剧本/编剧书挂到当前阶段；× 单独移除。", RED),
]
y = 1.45
for num, title, desc, color in items:
    badge(s, num, 8.55, y + 0.02, d=0.38, fill=color, size=8.5)
    text(s, title, 9.08, y, 3.7, 0.3, size=12.5, color=INK, bold=True)
    text(s, desc, 8.55, y + 0.4, 4.2, 0.75, size=9.6, color=INK_SOFT, line=1.5)
    y += 1.2

rect(s, 8.55, 5.2, 4.2, 1.15, fill=AMBER_SOFT, line_color=RGBColor(0xE3, 0xCF, 0xA5))
text(s, "改完记得点顶栏「保存」", 8.78, 5.38, 3.7, 0.28, size=11.5, color=AMBER, bold=True)
text(s, "三块内容会写入 plugin-state.json，并进入该阶段简报；不保存只在当前页面生效。",
     8.78, 5.72, 3.75, 0.55, size=9, color=INK_SOFT, line=1.45)
notes(s, "这是本次最值得现场演示的一页。三块都能逐条改；上传的知识库必须在这里挂到阶段，才会进入该阶段的简报。")

# ---------------------------------------------------------------- 08 保存 / 生成 / 示范
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "07 / 收尾", "保存、生成，以及用示范案例对照", 8)
cards2 = [
    ("1", "保存", "顶栏「保存」把项目卡、九阶段预设、会话、阶段说明、知识库引用写入 plugin-state.json。", TEAL),
    ("2", "生成全部产物", "生成 Excel 剧作蓝图 + Word 完整剧本，写入 产出/。需要结构化数据，空白项目按钮不可用。", AMBER),
    ("3", "示范案例", "空白项目 →「载入示范案例」看完整结果；看完点「清空演示，回到空白项目」。", RED),
]
for i, (num, title, desc, color) in enumerate(cards2):
    x = 0.55 + i * 4.15
    rect(s, x, 1.5, 3.85, 2.05, fill=CARD, line_color=LINE)
    badge(s, num, x + 0.22, 1.72, d=0.4, fill=color, size=10)
    text(s, title, x + 0.78, 1.7, 2.7, 0.32, size=13.5, color=INK, bold=True)
    text(s, desc, x + 0.22, 2.2, 3.4, 1.15, size=9.4, color=INK_SOFT, line=1.5)

rect(s, 0.55, 3.9, 7.4, 2.55, fill=INK, line_color=None)
text(s, "产出目录", 0.82, 4.08, 2.0, 0.26, size=9, color=RGBColor(0x9D, 0xC9, 0xCB), bold=True)
text(s, "编剧工坊 / 产出 /", 0.82, 4.37, 3.2, 0.32, size=12.5, color=WHITE, bold=True)
for i, (icon, name, meta) in enumerate([
    ("▦", "01_剧作蓝图_回声镇.xlsx", "14 张可见表"),
    ("✎", "02_第01集_完整剧本_回声镇.docx", "13 页"),
]):
    yy = 4.95 + i * 0.62
    text(s, icon, 0.85, yy, 0.3, 0.28, size=12, color=AMBER, bold=True)
    text(s, name, 1.2, yy - 0.01, 4.7, 0.28, size=10.2, color=WHITE, bold=True)
    text(s, meta, 6.0, yy, 1.7, 0.25, size=8.5, color=RGBColor(0xC9, 0xC0, 0xAE), align=PP_ALIGN.RIGHT)

rect(s, 8.2, 3.9, 4.6, 2.55, fill=AMBER_SOFT, line_color=RGBColor(0xE3, 0xCF, 0xA5))
text(s, "一句话记住", 8.45, 4.1, 3.0, 0.3, size=12.5, color=AMBER, bold=True)
text(s, "空白项目只做规划；\n要看 Excel/Word 成品，先载入示范案例或导入结构化数据。",
     8.45, 4.5, 4.0, 0.95, size=10.3, color=INK_SOFT, line=1.65)
text(s, "生成前先保存，避免刚改的设定没进产物。", 8.45, 5.62, 3.9, 0.35, size=9, color=MUTED, italic=True)
notes(s, "最后把闭环讲清楚：保存是持久化，生成是出文件；空白项目不能直接生成，必须用示范案例或结构化数据。")

# ---------------------------------------------------------------- 09 总结
s = prs.slides.add_slide(blank)
s.background.fill.solid(); s.background.fill.fore_color.rgb = PAPER
header(s, "08 / 一页带走", "讲解顺序：一条线走完五步", 9)
steps = [
    ("1", "填项目卡"), ("2", "点总控"), ("3", "选阶段"),
    ("4", "写会话 / 改预设 / 挂知识库"), ("5", "保存并生成"),
]
for i, (num, label) in enumerate(steps):
    x = 0.55 + i * 2.52
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.82), Inches(1.85), Inches(0.62), Inches(0.62))
    circ.fill.solid(); circ.fill.fore_color.rgb = TEAL if i < 3 else AMBER; circ.line.fill.background()
    tf = circ.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num; set_font(r, size=12, bold=True, color=WHITE)
    text(s, label, x, 2.65, 2.35, 0.62, size=10.5, color=INK, bold=True, align=PP_ALIGN.CENTER, line=1.35)
    if i < 4:
        line(s, x + 1.5, 2.16, x + 2.3, 2.16, color=LINE, width=1.4)

for i, (label, desc, color) in enumerate([
    ("左", "选阶段：九道工序是交付顺序", TEAL),
    ("中", "做内容：会话 / 预设 / 阶段说明", AMBER),
    ("右", "管项目：项目卡 / 总控 / 产出", INK),
]):
    x = 0.55 + i * 4.15
    rect(s, x, 3.85, 3.85, 1.55, fill=CARD, line_color=LINE)
    text(s, label, x + 0.22, 4.03, 0.6, 0.45, size=20, color=color, bold=True)
    text(s, desc, x + 0.82, 4.1, 2.85, 0.65, size=10.5, color=INK_SOFT, line=1.5)

text(s, "先让项目卡说清楚，再让每个阶段各自开工。", 0.55, 5.85, 11.8, 0.55, size=21, color=INK, bold=True)
line(s, 0.55, 6.48, 3.4, 6.48, color=AMBER, width=2)
text(s, "编剧工坊 · 操作指南 / 讲解结束", 0.55, 6.6, 6.0, 0.28, size=9, color=MUTED)
notes(s, "收尾只重复五个动作。现场如果时间有限，就按这条线演示一遍，其他细节留给文档。")

prs.save(str(OUT))
print("saved", OUT, "slides", len(prs.slides))
