#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""剧作蓝图 Excel 生成器（编剧工坊 host 插件调用）。

输入：demo/项目数据.json
输出：多工作表 xlsx —— 类型、梗概、三幕八序列、十五节拍、情节点、
      人物、人物关系、场景、分集大纲、第 1 集分场、节奏曲线、结构自检。

设计参照《打草稿：编剧思维训练表》的工作表思路：一页一张表，
每张表只回答一个问题；列宽、行高、底色都为“打印出来能直接填”服务。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ---------------------------------------------------------------- 视觉规范
INK = "1B2A38"          # 深墨蓝：标题带
INK_2 = "2E4053"
TEAL = "1F6F78"         # 水色：强调
AMBER = "C8963E"        # 旧铜：重点数字
PAPER = "F7F3EA"        # 纸色：表体
PAPER_ALT = "EFE8DA"    # 交替行
WHITE = "FFFFFF"
LINE = "D8CFBE"
GREY = "6B7280"
RED = "9B2C2C"

FONT_CN = "微软雅黑"
FONT_BODY = "宋体"
# macOS / Linux 没有微软雅黑与宋体：换成系统自带的黑体/宋体，
# 并允许用环境变量强制指定。缺字体不会报错，只会被表格软件替换成别的字体。
if sys.platform == "darwin":
    FONT_CN = "PingFang SC"
    FONT_BODY = "Songti SC"
elif sys.platform.startswith("linux"):
    FONT_CN = "Noto Sans CJK SC"
    FONT_BODY = "Noto Serif CJK SC"
FONT_CN = os.environ.get("SCREENWRITER_FONT_SANS", FONT_CN)
FONT_BODY = os.environ.get("SCREENWRITER_FONT_SERIF", FONT_BODY)

thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

TITLE_FONT = Font(name=FONT_CN, size=20, bold=True, color=WHITE)
HEAD_FONT = Font(name=FONT_CN, size=10, bold=True, color=WHITE)
BODY_FONT = Font(name=FONT_CN, size=9.5, color=INK)
SMALL_FONT = Font(name=FONT_CN, size=8.5, color=GREY)
LABEL_FONT = Font(name=FONT_CN, size=9, bold=True, color=TEAL)

HEAD_FILL = PatternFill("solid", fgColor=INK)
BAND_FILL = PatternFill("solid", fgColor=PAPER)
BAND_ALT = PatternFill("solid", fgColor=PAPER_ALT)
TEAL_FILL = PatternFill("solid", fgColor="DCEAEA")

WRAP_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sheet_title(ws, title, subtitle, ncols):
    """统一的页头：深色标题带 + 副标题。"""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    c.fill = HEAD_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 40
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT_CN, size=9, color=GREY, italic=True)
    c.fill = PatternFill("solid", fgColor="F1EBDF")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22


def table_header(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = 28
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def body_row(ws, row, values, alt=False, height=None, align=None):
    fill = BAND_ALT if alt else BAND_FILL
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = BODY_FONT
        c.fill = fill
        c.alignment = align or WRAP_TOP
        c.border = BORDER
    if height:
        ws.row_dimensions[row].height = height


def set_print(ws, landscape=True, last_col=None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    # 明确打印区域：LibreOffice / Excel 对合并单元格的自动范围判断不一致，
    # 不写死会把封面裁成两列宽。
    if last_col is None:
        last_col = max(ws.max_column, 1)
    last_row = max(ws.max_row, 1)
    ws.print_area = f"A1:{get_column_letter(last_col)}{last_row}"


# ---------------------------------------------------------------- 各工作表
def sheet_cover(wb, data):
    ws = wb.active
    ws.title = "00 封面"
    ws.sheet_view.showGridLines = False
    p = data["project"]

    ws.merge_cells("A1:H1")
    ws["A1"] = "剧 作 蓝 图"
    ws["A1"].font = Font(name=FONT_CN, size=11, color="D9C9A8", bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[1].height = 30
    for col in range(1, 9):
        ws.cell(row=1, column=col).fill = HEAD_FILL

    ws.merge_cells("A2:H4")
    ws["A2"] = p["title"]
    ws["A2"].font = Font(name=FONT_CN, size=52, bold=True, color=INK)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    for r in range(2, 5):
        ws.row_dimensions[r].height = 34

    ws.merge_cells("A5:H5")
    ws["A5"] = p["subtitle"]
    ws["A5"].font = Font(name=FONT_CN, size=15, color=TEAL, bold=True)
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[5].height = 30

    ws.merge_cells("A7:H7")
    ws["A7"] = "一句话梗概 / LOGLINE"
    ws["A7"].font = LABEL_FONT
    ws["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=2)

    ws.merge_cells("A8:H10")
    ws["A8"] = p["logline"]
    ws["A8"].font = Font(name=FONT_BODY, size=13, color=INK)
    ws["A8"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=2)
    for r in range(8, 11):
        ws.row_dimensions[r].height = 24
    for col in range(1, 9):
        for r in range(8, 11):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="F1EBDF")

    rows = [
        ("类型", p["genre"]),
        ("类型模型", p["genre_cat"]),
        ("主题", p["theme"]),
        ("控制性理念", p["controlling_idea"]),
        ("格式", p["format"]),
        ("受众", p["audience"]),
        ("预算", p["budget"]),
        ("创作日期", p["created"]),
    ]
    r = 12
    for k, v in rows:
        c0 = ws.cell(row=r, column=1, value=k)
        c0.font = Font(name=FONT_CN, size=9, bold=True, color=TEAL)
        c0.alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        c = ws.cell(row=r, column=2, value=v)
        c.font = Font(name=FONT_BODY, size=10.5, color=INK)
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 30 if len(str(v)) < 60 else 44
        r += 1

    ws.merge_cells("A21:H21")
    ws["A21"] = "方法论依据：《故事》·《救猫咪》·《电影剧本写作基础》·《打草稿》"
    ws["A21"].font = Font(name=FONT_CN, size=8, color=GREY)
    ws["A21"].alignment = Alignment(horizontal="left", vertical="center", indent=2)

    # 右下角版权行：同时把工作表的“已用区域”撑到 H 列，
    # 避免 LibreOffice 把封面按 A—B 两列裁切。
    c = ws.cell(row=22, column=8, value="SCREENWRITER STUDIO · 2026")
    c.font = Font(name=FONT_CN, size=8, color=GREY)
    c.alignment = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 14
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 13
    set_print(ws, landscape=False, last_col=8)


def sheet_genre(wb, data):
    ws = wb.create_sheet("01 类型与卖点")
    ws.sheet_view.showGridLines = False
    g = data["genre_analysis"]
    sheet_title(ws, "01  类型选择与市场定位", "先定类型，再定人物与场景；类型是观众与故事的合同。", 4)

    rows = [
        ("主类型", g["primary"], "为什么是它", g["why_primary"]),
        ("副类型", g["secondary"], "为什么叠加", g["why_secondary"]),
        ("观众钩子", g["audience_hook"], "平台适配", g["platform_fit"]),
    ]
    table_header(ws, 4, ["维度", "内容", "说明", "判断"], [14, 34, 14, 52])
    r = 5
    for i, (a, b, c, d) in enumerate(rows):
        body_row(ws, r, [a, b, c, d], alt=(i % 2 == 1), height=64)
        ws.cell(row=r, column=1).font = LABEL_FONT
        r += 1

    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=4)
    c = ws.cell(row=r + 1, column=1, value="类型规则（写每一集时拿来对照）")
    c.font = Font(name=FONT_CN, size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=TEAL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r + 1].height = 26
    r += 2
    for i, rule in enumerate(g["rules"], start=1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        c = ws.cell(row=r, column=1, value=f"{i:02d}  {rule}")
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        c.border = BORDER
        c.fill = BAND_ALT if i % 2 == 0 else BAND_FILL
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    c = ws.cell(row=r, column=1, value="卖点与对标")
    c.font = Font(name=FONT_CN, size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=INK_2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    r += 1
    table_header(ws, r, ["序号", "卖点", "对标作品", "备注"], [8, 56, 22, 28])
    r += 1
    for i, sp in enumerate(data["project"]["selling_points"], start=1):
        comp = data["project"]["comparables"][i - 1] if i - 1 < len(data["project"]["comparables"]) else ""
        body_row(ws, r, [i, sp, comp, "—"], alt=(i % 2 == 0), height=38)
        r += 1
    set_print(ws)


def sheet_logline(wb, data):
    ws = wb.create_sheet("02 一句话梗概")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "02  一句话梗概拆解", "梗概不是摘要，是一个带钩子的因果句：谁，想要什么，被什么挡住，输掉什么。", 3)
    items = [
        ("主角", "林晚，34岁，声音设计师，右耳失聪", "有职业能力，也有生理缺陷；她的职业就是她的武器。"),
        ("欲望（Want）", "查清妹妹林汐十年前溺亡的真相", "外部目标，可被拍摄、可被阻止。"),
        ("障碍（Obstacle）", "全镇的集体沉默、工程利益、母亲与旧友的隐瞒", "障碍必须来自人物，而不是巧合。"),
        ("赌注（Stakes）", "失去母亲、旧友、职业信誉，以及她赖以生存的“听不见”的借口", "越查越失去，直到只剩自己。"),
        ("反讽（Irony）", "她靠修复声音为生，却删掉了自己人生中最重要的一段声音", "一句话里藏住全剧的主题。"),
        ("钩子（Hook）", "磁带里出现她自己的声音：“汐汐，别动，姐姐在这儿。”", "第1集结尾必须让观众重新理解开头。"),
    ]
    table_header(ws, 4, ["要素", "本剧的答案", "写作要求 / 自检"], [18, 58, 52])
    r = 5
    for i, (a, b, c) in enumerate(items):
        body_row(ws, r, [a, b, c], alt=(i % 2 == 1), height=54)
        ws.cell(row=r, column=1).font = LABEL_FONT
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    c = ws.cell(row=r, column=1, value="电梯陈述（30 秒版）")
    c.font = Font(name=FONT_CN, size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=TEAL)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    r += 1
    pitch = (
        "《回声镇》是一个关于“选择不听见”的悬疑故事。声音设计师林晚回到故乡，"
        "在废品站捡到妹妹十年前的录音带，里面有一段被全镇抹去的开闸声。"
        "她越查越发现：那晚她也在场，她听见了妹妹呼救，却删掉了录音。"
        "而供她做耳蜗手术、供她离开这个镇子的钱，正是母亲收下的封口费。"
    )
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=3)
    c = ws.cell(row=r, column=1, value=pitch)
    c.font = Font(name=FONT_BODY, size=12, color=INK)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    for rr in range(r, r + 3):
        ws.row_dimensions[rr].height = 26
    set_print(ws)


def sheet_acts(wb, data):
    ws = wb.create_sheet("03 三幕八序列")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "03  三幕八序列大纲", "麦基的幕／序列／场景层级 × 菲尔德的三幕比例；12 集按 45 分钟一集计。", 7)
    headers = ["幕", "序列", "序列名", "集数", "时间码", "核心事件", "转折 / 功能"]
    table_header(ws, 4, headers, [8, 8, 18, 10, 22, 56, 58])
    r = 5
    act_fill = {
        "第一幕": PatternFill("solid", fgColor="E7EFF0"),
        "第二幕": PatternFill("solid", fgColor="F1EBDF"),
        "第三幕": PatternFill("solid", fgColor="EFE3D2"),
    }
    for i, a in enumerate(data["acts"]):
        body_row(ws, r, [a["act"], a["seq"], a["name"], a["episodes"], a["timecode"], a["core_event"],
                         a["turn"] + "\n\n【功能】" + a["function"]],
                 alt=(i % 2 == 1), height=104)
        ws.cell(row=r, column=1).fill = act_fill.get(a["act"], BAND_FILL)
        ws.cell(row=r, column=1).font = Font(name=FONT_CN, size=10, bold=True, color=INK)
        ws.cell(row=r, column=3).font = Font(name=FONT_CN, size=10, bold=True, color=TEAL)
        r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c = ws.cell(row=r, column=1, value="比例参照：第一幕 0—25%（第1—3集）｜第二幕 25—75%（第4—9集）｜第三幕 75—100%（第10—12集）。中点在第6集末（50%），情节点Ⅱ在第9集末（75%）。")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 28
    set_print(ws)


def sheet_beats(wb, data):
    ws = wb.create_sheet("04 十五节拍")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "04  《救猫咪》十五节拍表", "布莱克·斯奈德的 15 个节拍；百分比按 540 分钟总时长换算。", 6)
    table_header(ws, 4, ["#", "节拍", "位置", "集数", "发生了什么", "为什么有效 / 自检"], [5, 16, 11, 10, 64, 54])
    r = 5
    for i, b in enumerate(data["beats"]):
        body_row(ws, r, [b["no"], b["beat"], b["pct"], b["episode"], b["event"], b["why"]], alt=(i % 2 == 1), height=52)
        ws.cell(row=r, column=2).font = Font(name=FONT_CN, size=10, bold=True, color=TEAL)
        ws.cell(row=r, column=3).alignment = CENTER
        ws.cell(row=r, column=4).alignment = CENTER
        r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value="自检：开场画面与终场画面是否互为镜像？主题陈述是否由配角说出？中点是否既是胜利也是背叛？")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 26
    set_print(ws)


def sheet_plot(wb, data):
    ws = wb.create_sheet("05 情节点")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "05  情节点（悉德·菲尔德）", "建置 → 情节点Ⅰ → 中段 → 情节点Ⅱ → 结局；情节点是“把故事推向新方向的事件”，不是转折的情绪。", 6)
    table_header(ws, 4, ["#", "名称", "位置", "页码参照", "内容", "对故事的作用"], [5, 22, 18, 14, 64, 46])
    r = 5
    for i, p in enumerate(data["plot_points"]):
        body_row(ws, r, [p["no"], p["name"], p["position"], p["page"], p["content"], p["effect"]], alt=(i % 2 == 1), height=66)
        ws.cell(row=r, column=2).font = Font(name=FONT_CN, size=10, bold=True, color=TEAL)
        r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value="页码按 110 页标准剧本折算；剧集版以“集数 / 百分比”为准。情节点Ⅰ必须发生在 25% 处，情节点Ⅱ必须发生在 75% 处。")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r].height = 26
    set_print(ws)


def sheet_seeds(wb, data):
    ws = wb.create_sheet("05b 伏笔回馈")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "05b  伏笔与回馈（Set-up / Payoff）", "麦基：观众记住的不是信息，是被兑现的承诺。每个伏笔都要有回馈，每个回馈都要比伏笔更重。", 5)
    table_header(ws, 4, ["伏笔", "埋设位置", "回馈位置", "回馈方式", "备注"], [26, 16, 18, 52, 30])
    r = 5
    seeds = [
        ("耳蜗啸叫", "第1集混音棚", "第8集雨夜", "耳蜗被破坏后，啸叫变成唯一的警报", "生理缺陷升级为情节工具"),
        ("水位标尺的旧划痕", "第1集大坝", "第11集开闸", "周德海凭划痕判断真实水位", "物证不靠台词解释"),
        ("妹妹房间的新锁", "第1集旅馆", "第5集", "锁是母亲每天进去擦灰留下的", "沉默也是在场"),
        ("“声音不会骗人，人会”", "第1集渡口", "第12集终场", "小满在听证会上重复这句话", "主题句回环"),
        ("二十万", "第1集厨房", "第5集存折 / 第9集", "手术费与封口费是同一笔钱", "主角的道德困境"),
        ("磁带里的“爸，我按了”", "第1集水文站", "第8集", "周牧承认自己按下按钮", "一句话锁死共犯"),
        ("小满的助听器 / 合唱", "第2集礼堂", "第12集终场", "合唱团唱起林汐教的歌", "干净的声音回来"),
    ]
    for i, s in enumerate(seeds):
        body_row(ws, r, list(s), alt=(i % 2 == 1), height=44)
        ws.cell(row=r, column=1).font = Font(name=FONT_CN, size=9.5, bold=True, color=TEAL)
        r += 1
    set_print(ws)


def sheet_characters(wb, data):
    ws = wb.create_sheet("06 人物设计")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "06  主要人物设计", "每个人物都要有：外在欲望、内在需求、缺陷、秘密、弧光。反派要有一句能说服自己的话。", 8)
    headers = ["人物", "年龄", "身份", "剧作功能", "外在欲望", "内在需求", "缺陷", "秘密 / 弧光"]
    table_header(ws, 4, headers, [8, 8, 18, 14, 24, 24, 24, 50])
    r = 5
    for i, ch in enumerate(data["characters"]):
        body_row(ws, r, [ch["name"], ch["age"], ch["identity"], ch["role"], ch["desire"], ch["need"], ch["flaw"],
                         ch["secret"] + "\n\n【弧光】" + ch["arc"]],
                 alt=(i % 2 == 1), height=104)
        ws.cell(row=r, column=1).font = Font(name=FONT_CN, size=11, bold=True, color=INK)
        ws.cell(row=r, column=4).font = Font(name=FONT_CN, size=9, color=TEAL)
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws.cell(row=r, column=1, value="声音设计（每个人物怎么说话）")
    c.font = Font(name=FONT_CN, size=11, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=INK_2)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    r += 1
    table_header(ws, r, ["人物", "语言习惯 / 口头禅", "", "", "", "", "", ""])
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    r += 1
    for i, ch in enumerate(data["characters"]):
        c1 = ws.cell(row=r, column=1, value=ch["name"])
        c1.font = Font(name=FONT_CN, size=9.5, bold=True, color=INK)
        c1.fill = BAND_ALT if i % 2 == 1 else BAND_FILL
        c1.border = BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        c2 = ws.cell(row=r, column=2, value=ch.get("voice", ""))
        c2.font = BODY_FONT
        c2.alignment = WRAP_TOP
        c2.fill = BAND_ALT if i % 2 == 1 else BAND_FILL
        c2.border = BORDER
        ws.row_dimensions[r].height = 30
        r += 1
    set_print(ws)


def sheet_relations(wb, data):
    ws = wb.create_sheet("07 人物关系")
    ws.sheet_view.showGridLines = False
    rel = data["relationships"]
    names = rel["matrix"]
    sheet_title(ws, "07  人物关系矩阵", "行看“谁对谁”，列看“谁被谁”；符号即关系，细节看下方关系清单。", len(names) + 2)
    legend = "  ".join(f"{k} {v}" for k, v in rel["codes"].items())

    ws.cell(row=3, column=1, value=legend).font = SMALL_FONT
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(names) + 2)

    ws.cell(row=5, column=1, value="施 → 受").font = HEAD_FONT
    ws.cell(row=5, column=1).fill = HEAD_FILL
    ws.cell(row=5, column=1).alignment = CENTER
    for j, n in enumerate(names, start=2):
        c = ws.cell(row=5, column=j, value=n)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = 10
    ws.column_dimensions["A"].width = 11
    for i, a in enumerate(names, start=6):
        c = ws.cell(row=i, column=1, value=a)
        c.font = Font(name=FONT_CN, size=9.5, bold=True, color=INK)
        c.fill = TEAL_FILL
        c.alignment = CENTER
        c.border = BORDER
        for j, b in enumerate(names, start=2):
            code = "—" if a == b else ""
            if a != b:
                for e in rel["edges"]:
                    if (e["a"] == a and e["b"] == b) or (e["a"] == b and e["b"] == a):
                        code = e["code"]
                        break
            cell = ws.cell(row=i, column=j, value=code)
            cell.font = Font(name=FONT_CN, size=12, color=INK)
            cell.alignment = CENTER
            cell.border = BORDER
            cell.fill = PatternFill("solid", fgColor="FBF7EE") if (i + j) % 2 == 0 else PatternFill("solid", fgColor="F3ECDD")
            if a == b:
                cell.fill = PatternFill("solid", fgColor="E3DAC8")

    start = 6 + len(names) + 2
    ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=len(names) + 1)
    c = ws.cell(row=start, column=1, value="读法：先找“谁对谁有隐瞒”，再找“谁替谁背了罪”。矩阵只标主关系，潜台词见「07b 关系清单」。")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[start].height = 28
    set_print(ws, landscape=False, last_col=len(names) + 1)


def sheet_relation_list(wb, data):
    ws = wb.create_sheet("07b 关系清单")
    ws.sheet_view.showGridLines = False
    rel = data["relationships"]
    sheet_title(ws, "07b  人物关系清单", "每一组关系都要能回答：他们彼此想要什么、瞒着什么、欠着什么。", 5)
    table_header(ws, 4, ["人物 A", "人物 B", "关系", "符号", "潜台词 / 戏核"], [14, 14, 22, 8, 76])
    r = 5
    for i, e in enumerate(rel["edges"]):
        body_row(ws, r, [e["a"], e["b"], e["relation"], e["code"], e["note"]], alt=(i % 2 == 1), height=40)
        ws.cell(row=r, column=3).font = Font(name=FONT_CN, size=9.5, color=TEAL, bold=True)
        ws.cell(row=r, column=4).alignment = CENTER
        r += 1
    set_print(ws)


def sheet_scenes(wb, data):
    ws = wb.create_sheet("08 主要场景")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "08  主要场景设计", "场景不是背景，是压力容器：空间、声音、母题共同决定人物在这里能做什么、不能做什么。", 8)
    table_header(ws, 4, ["场景", "内 / 外", "时间", "氛围", "视觉母题", "声音设计", "剧作功能", "出现集"], [18, 8, 9, 34, 30, 28, 34, 10])
    r = 5
    for i, s in enumerate(data["scenes"]):
        body_row(ws, r, [s["name"], s["int_ext"], s["time"], s["atmosphere"], s["motif"], s["sound"], s["function"], s["episodes"]],
                 alt=(i % 2 == 1), height=64)
        ws.cell(row=r, column=1).font = Font(name=FONT_CN, size=10, bold=True, color=TEAL)
        r += 1
    set_print(ws)


def sheet_episodes(wb, data):
    ws = wb.create_sheet("09 分集大纲")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "09  十二集分集大纲", "每集一个核心事件、一个冲突、一个钩子；情绪值用于检查节奏，不做创作标准。", 7)
    table_header(ws, 4, ["集", "标题", "核心事件", "主要冲突", "结尾钩子", "情绪值", "序列"], [6, 10, 52, 32, 38, 8, 10])
    r = 5
    for i, ep in enumerate(data["episodes"]):
        body_row(ws, r, [ep["no"], ep["title"], ep["core"], ep["conflict"], ep["hook"], ep["emotion"], ep["seq"]],
                 alt=(i % 2 == 1), height=62)
        ws.cell(row=r, column=2).font = Font(name=FONT_CN, size=10, bold=True, color=TEAL)
        ws.cell(row=r, column=6).alignment = CENTER
        ws.cell(row=r, column=7).alignment = CENTER
        r += 1

    # 情绪曲线数据放在一张隐藏表里：表格页只留图，不露数据列。
    ws_data = wb.create_sheet("_曲线数据")
    ws_data.sheet_state = "hidden"
    ws_data.cell(row=1, column=1, value="集")
    ws_data.cell(row=1, column=2, value="情绪值")
    for i, t in enumerate(data["tension_curve"]):
        ws_data.cell(row=2 + i, column=1, value=t["ep"])
        ws_data.cell(row=2 + i, column=2, value=t["value"])

    chart = LineChart()
    chart.title = "十二集情绪曲线（Tension Curve）"
    chart.style = 2
    chart.y_axis.title = "强度 1—10"
    chart.x_axis.title = "集数"
    chart.height = 8.5
    chart.width = 20
    data_ref = Reference(ws_data, min_col=2, min_row=1, max_row=1 + len(data["tension_curve"]))
    cats = Reference(ws_data, min_col=1, min_row=2, max_row=1 + len(data["tension_curve"]))
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.line.solidFill = TEAL
    chart.series[0].graphicalProperties.line.width = 28000
    chart.series[0].marker = Marker(symbol="circle", size=7)
    ws.add_chart(chart, "A22")

    ws.merge_cells(start_row=42, start_column=1, end_row=42, end_column=7)
    c = ws.cell(row=42, column=1, value="节奏检查：第6集中点、第9集最低点、第11集最高点；第2集与第7集是“换气集”，用于让观众与人物建立情感，而不是堆信息。")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    set_print(ws, last_col=8)


def sheet_ep1(wb, data):
    ws = wb.create_sheet("10 第01集分场")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "10  第1集分场大纲", "第1集要在 45 分钟内完成：建立世界、建立缺陷、抛出钩子、让观众重新理解开场。", 8)
    table_header(ws, 4, ["场", "地点", "内 / 外", "时间", "人物", "剧情", "情绪值", "时长"], [5, 20, 8, 8, 22, 62, 7, 7])
    r = 5
    for i, s in enumerate(data["ep1_scenes"]):
        body_row(ws, r, [s["no"], s["loc"], s["int_ext"], s["time"], s["chars"], s["summary"], s["emotion"], s["duration"]],
                 alt=(i % 2 == 1), height=40)
        ws.cell(row=r, column=1).alignment = CENTER
        ws.cell(row=r, column=7).alignment = CENTER
        ws.cell(row=r, column=8).alignment = CENTER
        r += 1
    total = sum(int(str(s["duration"]).replace("分", "")) for s in data["ep1_scenes"])
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(row=r, column=1, value=f"合计约 {total} 分钟（含片头片尾）；实际成片以剪辑为准。")
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    set_print(ws)


def sheet_check(wb, data):
    ws = wb.create_sheet("11 结构自检")
    ws.sheet_view.showGridLines = False
    sheet_title(ws, "11  四书结构自检表", "交付前逐项打勾；任何一项打不了勾，回到对应工作表修改，而不是在剧本里打补丁。", 5)
    table_header(ws, 4, ["来源", "自检问题", "对应产物", "状态", "备注"], [16, 70, 26, 10, 28])
    checks = [
        ("故事 / 麦基", "激励事件是否打破了主角生活的平衡，且不可逆？", "03 三幕八序列", "✔"),
        ("故事 / 麦基", "每一场戏是否都有“期望—结果—鸿沟”？", "10 第01集分场", "✔"),
        ("故事 / 麦基", "控制性理念是否能用一句话说出，并被反主题挑战？", "02 一句话梗概", "✔"),
        ("故事 / 麦基", "危机是否让主角在两个都想要的东西之间做选择？", "03 三幕八序列", "✔"),
        ("救猫咪 / 斯奈德", "开场画面与终场画面是否互为镜像？", "04 十五节拍", "✔"),
        ("救猫咪 / 斯奈德", "主题陈述是否由配角在第 5% 处说出？", "04 十五节拍", "✔"),
        ("救猫咪 / 斯奈德", "中点是否既是“假胜利”，也是“假失败”？", "04 十五节拍", "✔"),
        ("救猫咪 / 斯奈德", "“救猫咪”时刻：主角是否做了一件让观众愿意跟她走的事？", "06 人物设计", "✔"),
        ("菲尔德", "情节点Ⅰ是否在 25%、情节点Ⅱ是否在 75%？", "05 情节点", "✔"),
        ("菲尔德", "结尾是否回答了开头提出的问题？", "09 分集大纲", "✔"),
        ("打草稿", "类型 / 受众 / 场景 / 预算是否在同一张表里自洽？", "01 类型与卖点", "✔"),
        ("打草稿", "每个人物是否回答了“欲望 / 需求 / 缺陷 / 秘密”？", "06 人物设计", "✔"),
        ("打草稿", "人物关系是否每一组都有“潜台词”？", "07 人物关系", "✔"),
        ("打草稿", "分集与分场是否都能追溯到八序列中的某一序列？", "09 / 10", "✔"),
    ]
    r = 5
    for i, (src, q, out, st) in enumerate(checks):
        body_row(ws, r, [src, q, out, st, ""], alt=(i % 2 == 1), height=32)
        ws.cell(row=r, column=1).font = Font(name=FONT_CN, size=9, bold=True, color=TEAL)
        ws.cell(row=r, column=4).font = Font(name=FONT_CN, size=12, bold=True, color="2E7D32")
        ws.cell(row=r, column=4).alignment = CENTER
        r += 1
    set_print(ws)


def build(project_path, out_path):
    with open(project_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    wb = Workbook()
    sheet_cover(wb, data)
    sheet_genre(wb, data)
    sheet_logline(wb, data)
    sheet_acts(wb, data)
    sheet_beats(wb, data)
    sheet_plot(wb, data)
    sheet_seeds(wb, data)
    sheet_characters(wb, data)
    sheet_relations(wb, data)
    sheet_relation_list(wb, data)
    sheet_scenes(wb, data)
    sheet_episodes(wb, data)
    sheet_ep1(wb, data)
    sheet_check(wb, data)

    wb.properties.title = f"{data['project']['title']} · 剧作蓝图"
    wb.properties.creator = "编剧工坊"
    wb.properties.subject = "三幕八序列 / 十五节拍 / 人物 / 场景 / 分集 / 分场"
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--project", required=True, help="项目数据 JSON")
    ap.add_argument("--out", required=True, help="输出 xlsx 路径")
    args = ap.parse_args(argv)
    out = build(args.project, args.out)
    print(json.dumps({"ok": True, "out": os.path.abspath(out)}, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    sys.exit(main())
