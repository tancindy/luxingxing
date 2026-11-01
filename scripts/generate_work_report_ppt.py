# -*- coding: utf-8 -*-
"""生成《2026 年度工作汇报》PPT（中文，16:9，共 13 页）。

用法:
    python3 scripts/generate_work_report_ppt.py

输出:
    ./工作汇报.pptx
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- 主题配色
BLUE_DARK = RGBColor.from_string("1F4E79")   # 主色：深蓝
BLUE = RGBColor.from_string("2E75B6")        # 次色：亮蓝
BLUE_LIGHT = RGBColor.from_string("DEEBF7")  # 浅蓝
ORANGE = RGBColor.from_string("ED7D31")      # 强调色：橙
GREEN = RGBColor.from_string("70AD47")       # 辅助色：绿
DARK = RGBColor.from_string("262626")        # 正文深灰
TEXT = RGBColor.from_string("404040")        # 正文灰
GRAY = RGBColor.from_string("8C8C8C")        # 辅助灰
LIGHT = RGBColor.from_string("F2F6FA")       # 卡片浅底
LINE = RGBColor.from_string("D9E2F0")        # 分隔线
WHITE = RGBColor.from_string("FFFFFF")

FONT = "微软雅黑"
SLIDE_W, SLIDE_H = 13.333, 7.5
TOTAL_PAGES = 13


# ---------------------------------------------------------------- 基础工具
def set_run_font(run, size=16, bold=False, color=TEXT, name=FONT, italic=False):
    f = run.font
    f.name = name
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    # 同时设置东亚字体，保证中文正常显示
    rPr = run._r.get_or_add_rPr()
    latin = rPr.get_or_add_latin()
    latin.set("typeface", name)
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = parse_xml(
            '<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            ' typeface="%s"/>' % name
        )
        latin.addnext(ea)
    else:
        ea.set("typeface", name)


def add_rect(slide, x, y, w, h, fill, line=None, rounded=False, radius=0.08):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, line_spacing=1.25, space_after=0):
    """runs: [[(text, kwargs), ...], ...]  每个子列表是一个段落。"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para_runs in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        for text, kwargs in para_runs:
            r = p.add_run()
            r.text = text
            set_run_font(r, **kwargs)
    return tb


def add_bullets(slide, items, x, y, w, h, size=16, gap=8, line_spacing=1.3,
                bullet_color=ORANGE, text_color=TEXT):
    """items: [str] 或 [(粗体前缀, 其余文本)]"""
    runs = []
    for item in items:
        para = [("▪  ", dict(size=size, bold=True, color=bullet_color))]
        if isinstance(item, tuple):
            para.append((item[0], dict(size=size, bold=True, color=DARK)))
            para.append((item[1], dict(size=size, color=text_color)))
        else:
            para.append((item, dict(size=size, color=text_color)))
        runs.append(para)
    return add_text(slide, x, y, w, h, runs, line_spacing=line_spacing, space_after=gap)


def add_header(slide, title, page_no):
    add_rect(slide, 0.55, 0.5, 0.09, 0.52, ORANGE)
    add_text(slide, 0.78, 0.42, 11.5, 0.7,
             [[(title, dict(size=27, bold=True, color=BLUE_DARK))]])
    add_rect(slide, 0.55, 1.18, 12.23, 0.018, LINE)
    add_text(slide, 0.55, 7.02, 5.0, 0.35,
             [[("2026 年度工作汇报", dict(size=10, color=GRAY))]])
    add_text(slide, 11.3, 7.02, 1.48, 0.35,
             [[("%02d / %d" % (page_no, TOTAL_PAGES), dict(size=10, color=GRAY))]],
             align=PP_ALIGN.RIGHT)


def style_cell(cell, text, size=13, bold=False, color=TEXT, fill=None,
               align=PP_ALIGN.CENTER):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run_font(r, size=size, bold=bold, color=color)


# ---------------------------------------------------------------- 幻灯片内容
def slide_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, BLUE_DARK)
    # 装饰
    add_rect(s, -1.2, -1.2, 3.6, 3.6, BLUE, rounded=True, radius=0.5)
    add_rect(s, 11.4, 5.2, 3.4, 3.4, BLUE, rounded=True, radius=0.5)
    add_rect(s, 12.2, -0.9, 2.0, 2.0, RGBColor.from_string("2A5F8F"),
             rounded=True, radius=0.5)
    add_rect(s, 1.4, 2.62, 1.6, 0.10, ORANGE)
    add_text(s, 1.4, 2.85, 10.5, 1.2,
             [[("2026 年度工作汇报", dict(size=44, bold=True, color=WHITE))]])
    add_text(s, 1.42, 4.05, 10.5, 0.5,
             [[("回顾总结  ·  成果展示  ·  规划展望", dict(size=18, color=BLUE_LIGHT))]])
    add_rect(s, 1.42, 4.78, 6.2, 0.015, RGBColor.from_string("5B8DB8"))
    add_text(s, 1.42, 5.05, 10.5, 0.9,
             [[("汇报人：XXX          部门：XXX 部", dict(size=14, color=BLUE_LIGHT))],
              [("汇报时间：2026 年 11 月", dict(size=14, color=BLUE_LIGHT))]],
             line_spacing=1.5)
    return s


def slide_agenda(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "目 录", 2)
    items = [
        ("01", "工作概述", "年度目标与整体进展"),
        ("02", "重点工作完成情况", "关键任务与完成质量"),
        ("03", "数据与成果展示", "核心指标与业务数据"),
        ("04", "问题与经验总结", "不足之处与收获体会"),
        ("05", "下阶段工作计划", "目标、里程碑与资源需求"),
        ("06", "总结与展望", "回顾全年，展望未来"),
    ]
    x0, y0, w, h, gx, gy = 0.75, 1.62, 5.95, 1.55, 0.45, 0.28
    for i, (num, title, desc) in enumerate(items):
        col, row = i % 2, i // 2
        x = x0 + col * (w + gx)
        y = y0 + row * (h + gy)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_rect(s, x, y, 0.08, h, BLUE if col == 0 else ORANGE, rounded=True, radius=0.5)
        add_text(s, x + 0.3, y + 0.18, 1.1, 1.2,
                 [[(num, dict(size=30, bold=True, color=BLUE if col == 0 else ORANGE))]])
        add_text(s, x + 1.45, y + 0.28, w - 1.7, 1.1,
                 [[(title, dict(size=18, bold=True, color=DARK))],
                  [(desc, dict(size=12, color=GRAY))]],
                 line_spacing=1.35, space_after=4)
    return s


def slide_overview(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "一、工作概述", 3)
    add_bullets(s, [
        ("总体思路：", "围绕年度战略目标，聚焦核心业务与重点项目，稳中求进、提质增效。"),
        ("执行情况：", "全年累计完成 XX 项重点工作，XX 个里程碑节点按时交付，整体进展符合预期。"),
        ("团队建设：", "团队规模达 XX 人，覆盖研发、运营、市场三大职能，人才结构持续优化。"),
    ], 0.75, 1.55, 11.8, 1.9, size=16)
    kpis = [
        ("92%", "年度目标完成率", BLUE),
        ("12 个", "重点项目交付", ORANGE),
        ("95%", "客户满意度", GREEN),
        ("+28%", "营收同比增长", BLUE_DARK),
    ]
    x0, y, w, h, gap = 0.75, 3.85, 2.87, 1.75, 0.24
    for i, (num, label, color) in enumerate(kpis):
        x = x0 + i * (w + gap)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_rect(s, x, y, w, 0.07, color, rounded=True, radius=0.5)
        add_text(s, x, y + 0.35, w, 0.8,
                 [[(num, dict(size=32, bold=True, color=color))]], align=PP_ALIGN.CENTER)
        add_text(s, x, y + 1.12, w, 0.5,
                 [[(label, dict(size=13, color=TEXT))]], align=PP_ALIGN.CENTER)
    add_text(s, 0.75, 5.95, 11.8, 0.6,
             [[("※ 以上数据为模板示例，请根据实际情况替换。", dict(size=11, color=GRAY))]])
    return s


def slide_keywork(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "二、重点工作完成情况", 4)
    rows, cols = 6, 5
    left, top, width, height = 0.75, 1.6, 11.83, 4.55
    shape = s.shapes.add_table(rows, cols, Inches(left), Inches(top),
                               Inches(width), Inches(height))
    table = shape.table
    col_w = [3.0, 2.35, 3.25, 1.45, 1.78]
    for i, w in enumerate(col_w):
        table.columns[i].width = Inches(w)
    headers = ["重点工作", "目标要求", "完成情况", "完成度", "负责人"]
    data = [
        ["XXX 平台升级改造", "Q2 完成上线", "已上线并稳定运行", "100%", "张三"],
        ["客户增长专项", "新增 500 家", "新增 560 家", "112%", "李四"],
        ["内部流程优化", "效率提升 20%", "效率提升 25%", "125%", "王五"],
        ["团队建设与招聘", "招聘 10 人", "入职 8 人", "80%", "赵六"],
        ["品牌宣传推广", "媒体曝光 50 次", "曝光 46 次", "92%", "钱七"],
    ]
    for c, htext in enumerate(headers):
        style_cell(table.cell(0, c), htext, size=14, bold=True, color=WHITE, fill=BLUE_DARK)
    for r, row in enumerate(data, start=1):
        fill = WHITE if r % 2 == 1 else LIGHT
        for c, val in enumerate(row):
            bold = (c == 3)
            color = GREEN if (c == 3 and val not in ("80%", "92%")) else (ORANGE if c == 3 else TEXT)
            if c == 0:
                style_cell(table.cell(r, c), val, size=13, bold=True, color=DARK,
                           fill=fill, align=PP_ALIGN.LEFT)
            else:
                style_cell(table.cell(r, c), val, size=13, bold=bold, color=color, fill=fill)
    table.rows[0].height = Inches(0.55)
    for r in range(1, rows):
        table.rows[r].height = Inches(0.8)
    add_text(s, 0.75, 6.35, 11.8, 0.5,
             [[("※ 表格内容为示例，可按实际项目增删行、修改列。", dict(size=11, color=GRAY))]])
    return s


def slide_data(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "三、数据与成果展示", 5)
    # 图表
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("实际销售额（万元）", (420, 560, 610, 780))
    chart_data.add_series("销售目标（万元）", (400, 520, 600, 700))
    gf = s.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.75), Inches(1.55), Inches(7.35), Inches(4.9), chart_data,
    )
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.font.size = Pt(11)
    chart.font.name = FONT
    plot = chart.plots[0]
    plot.gap_width = 90
    plot.has_data_labels = True
    plot.data_labels.number_format = "0"
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.font.size = Pt(9)
    s1, s2 = chart.series[0], chart.series[1]
    s1.format.fill.solid()
    s1.format.fill.fore_color.rgb = BLUE
    s2.format.fill.solid()
    s2.format.fill.fore_color.rgb = RGBColor.from_string("BDD7EE")
    # 右侧要点
    add_rect(s, 8.45, 1.55, 4.13, 4.9, LIGHT, rounded=True)
    add_text(s, 8.75, 1.85, 3.6, 0.5,
             [[("核心要点", dict(size=17, bold=True, color=BLUE_DARK))]])
    add_bullets(s, [
        ("Q4 销售额 780 万元", "，同比增长 28%，超额完成年度目标。"),
        ("新客户贡献营收占比 35%", "，客户结构持续优化。"),
        ("运营成本率下降 3.2 个百分点", "，盈利能力稳步提升。"),
        ("复购率提升至 62%", "，产品与服务口碑向好。"),
    ], 8.72, 2.45, 3.65, 3.8, size=13, gap=12, line_spacing=1.35)
    return s


def slide_highlights(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "四、项目亮点展示", 6)
    cards = [
        ("01", "XXX 平台成功上线", "系统可用性达 99.9%，用户活跃度提升 40%，获用户好评。", BLUE),
        ("02", "客户服务体系升级", "平均响应时长缩短至 2 小时，客户满意度提升至 95%。", ORANGE),
        ("03", "数字化运营转型", "自动化流程覆盖 80% 核心业务，人效同比提升 25%。", GREEN),
        ("04", "品牌影响力提升", "行业峰会主题演讲 3 次，主流媒体曝光 46 次。", BLUE_DARK),
    ]
    x0, y0, w, h, gx, gy = 0.75, 1.55, 5.94, 2.42, 0.45, 0.3
    for i, (num, title, desc, color) in enumerate(cards):
        col, row = i % 2, i // 2
        x = x0 + col * (w + gx)
        y = y0 + row * (h + gy)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_rect(s, x, y, 0.09, h, color, rounded=True, radius=0.5)
        add_text(s, x + 0.35, y + 0.3, 1.0, 1.0,
                 [[(num, dict(size=26, bold=True, color=color))]])
        add_text(s, x + 1.35, y + 0.32, w - 1.65, 1.9,
                 [[(title, dict(size=17, bold=True, color=DARK))],
                  [(desc, dict(size=13, color=TEXT))]],
                 line_spacing=1.35, space_after=6)
    return s


def slide_problems(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "五、存在的问题与不足", 7)
    blocks = [
        ("目标执行方面", [
            "部分项目进度滞后于原定计划，需加强过程管控；",
            "跨部门协同效率有待提升，信息同步不够及时；",
            "需求变更管理不够规范，导致资源重复投入。",
        ], BLUE),
        ("团队能力方面", [
            "高端人才与关键技术岗位储备不足；",
            "新技术、新工具的应用能力仍需加强；",
            "新人培养周期偏长，上手速度有待提高。",
        ], ORANGE),
        ("流程管理方面", [
            "部分制度流程繁琐，影响执行效率；",
            "数据统计口径不统一，分析结论存在偏差；",
            "风险预警机制尚不健全，应对偏被动。",
        ], GREEN),
    ]
    x0, y, w, h, gap = 0.75, 1.55, 3.83, 4.9, 0.27
    for i, (title, items, color) in enumerate(blocks):
        x = x0 + i * (w + gap)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_rect(s, x, y, w, 0.62, color, rounded=True, radius=0.18)
        add_rect(s, x, y + 0.32, w, 0.3, color)  # 覆盖圆角下沿，形成平头标题条
        add_text(s, x, y + 0.09, w, 0.5,
                 [[(title, dict(size=16, bold=True, color=WHITE))]], align=PP_ALIGN.CENTER)
        add_bullets(s, items, x + 0.28, y + 0.95, w - 0.5, h - 1.2,
                    size=13, gap=10, line_spacing=1.35, bullet_color=color)
    return s


def slide_lessons(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "六、经验总结与收获", 8)
    add_bullets(s, [
        ("坚持目标导向：", "以季度目标拆解任务、责任到人，确保方向不偏移、执行不落空。"),
        ("强化数据驱动：", "建立周度数据复盘机制，及时发现问题、快速纠偏调整。"),
        ("重视人才培养：", "通过“老带新”、专项培训与轮岗锻炼，团队战斗力明显增强。"),
        ("优化协作机制：", "推行双周同步会与共享文档，跨部门信息差明显减少。"),
        ("沉淀方法资产：", "将项目经验固化为 SOP 与模板库，复用效率显著提升。"),
    ], 0.75, 1.65, 11.8, 3.6, size=16, gap=16)
    # 底部金句
    add_rect(s, 0.75, 5.65, 11.83, 1.05, BLUE_DARK, rounded=True)
    add_text(s, 0.75, 5.87, 11.83, 0.6,
             [[("“ 把每一次复盘，都变成下一次进步的起点。 ”",
                dict(size=17, bold=True, color=WHITE))]],
             align=PP_ALIGN.CENTER)
    return s


def slide_plan(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "七、下阶段工作计划", 9)
    cards = [
        ("计划一", "推进 XXX 二期建设", "Q1 完成需求评审与方案设计，Q2 完成开发测试并正式上线。", BLUE),
        ("计划二", "深耕重点行业客户", "目标新增签约 80 家，营收增长 30%，打造 5 个标杆案例。", ORANGE),
        ("计划三", "完善人才梯队建设", "新增核心技术岗位 5 人，开展专项培训 12 场，完善考核机制。", GREEN),
    ]
    x0, y, w, h, gap = 0.75, 1.6, 3.83, 2.75, 0.27
    for i, (tag, title, desc, color) in enumerate(cards):
        x = x0 + i * (w + gap)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_rect(s, x + 0.3, y + 0.32, 1.15, 0.45, color, rounded=True, radius=0.35)
        add_text(s, x + 0.3, y + 0.38, 1.15, 0.4,
                 [[(tag, dict(size=12, bold=True, color=WHITE))]], align=PP_ALIGN.CENTER)
        add_text(s, x + 0.3, y + 1.0, w - 0.6, 0.55,
                 [[(title, dict(size=16, bold=True, color=DARK))]])
        add_text(s, x + 0.3, y + 1.6, w - 0.6, 1.0,
                 [[(desc, dict(size=12.5, color=TEXT))]], line_spacing=1.35)
    add_text(s, 0.75, 4.75, 11.8, 0.5,
             [[("执行保障：", dict(size=15, bold=True, color=BLUE_DARK))]])
    add_bullets(s, [
        ("责任到人：", "每项计划明确负责人与关键结果，纳入季度考核。"),
        ("节奏管控：", "月度检查、季度复盘，关键节点设置预警与应急预案。"),
    ], 0.75, 5.25, 11.8, 1.4, size=14, gap=8)
    return s


def slide_timeline(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "八、时间安排与里程碑", 10)
    # 时间轴主线
    add_rect(s, 1.1, 3.62, 11.1, 0.035, LINE)
    nodes = [
        ("Q1", "年度规划启动", "完成目标拆解、方案评审、二期项目立项启动", BLUE),
        ("Q2", "系统上线复盘", "二期系统正式上线，完成半年度经营复盘", ORANGE),
        ("Q3", "客户攻坚优化", "重点行业客户攻坚，流程优化专项落地", GREEN),
        ("Q4", "总结规划筹备", "年度总结与考核，次年规划与预算编制", BLUE_DARK),
    ]
    centers = [2.15, 5.15, 8.15, 11.15]
    d = 1.0
    for i, ((q, title, desc, color), cx) in enumerate(zip(nodes, centers)):
        add_rect(s, cx - d / 2, 3.62 - d / 2 + 0.017, d, d, color,
                 rounded=True, radius=0.5)
        add_text(s, cx - d / 2, 3.62 - 0.24, d, 0.5,
                 [[(q, dict(size=18, bold=True, color=WHITE))]], align=PP_ALIGN.CENTER)
        card_w, card_h = 2.75, 1.65
        above = (i % 2 == 0)
        cy = 1.62 if above else 4.55
        add_rect(s, cx - card_w / 2, cy, card_w, card_h, LIGHT, rounded=True)
        strip_x = cx - card_w / 2
        add_rect(s, strip_x, cy if above else cy + card_h - 0.07, card_w, 0.07,
                 color, rounded=True, radius=0.5)
        add_text(s, strip_x + 0.18, cy + 0.18, card_w - 0.36, 1.35,
                 [[(title, dict(size=14, bold=True, color=DARK))],
                  [(desc, dict(size=11, color=TEXT))]],
                 line_spacing=1.25, space_after=5)
    return s


def slide_resources(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "九、团队协作与资源需求", 11)
    rows, cols = 5, 4
    left, top, width, height = 0.75, 1.6, 11.83, 3.6
    shape = s.shapes.add_table(rows, cols, Inches(left), Inches(top),
                               Inches(width), Inches(height))
    table = shape.table
    col_w = [1.9, 4.9, 1.85, 3.18]
    for i, w in enumerate(col_w):
        table.columns[i].width = Inches(w)
    headers = ["需求类别", "具体内容", "预计投入", "预期效果"]
    data = [
        ["人力支持", "新增研发 3 人、运营 2 人，补齐关键岗位", "Q1 到位", "保障二期项目按期交付"],
        ["预算支持", "市场推广预算 XX 万元、培训预算 XX 万元", "XX 万元", "支撑增长目标达成"],
        ["技术支持", "引入 XX 数据平台，升级内部协作工具", "Q2 完成", "数据驱动决策提效"],
        ["跨部门协作", "建立联合项目组，明确接口人与同步机制", "长期机制", "协同效率提升 30%"],
    ]
    for c, htext in enumerate(headers):
        style_cell(table.cell(0, c), htext, size=14, bold=True, color=WHITE, fill=BLUE_DARK)
    for r, row in enumerate(data, start=1):
        fill = WHITE if r % 2 == 1 else LIGHT
        for c, val in enumerate(row):
            if c == 0:
                style_cell(table.cell(r, c), val, size=13, bold=True, color=BLUE_DARK, fill=fill)
            else:
                style_cell(table.cell(r, c), val, size=12.5, color=TEXT, fill=fill,
                           align=PP_ALIGN.LEFT if c == 1 else PP_ALIGN.CENTER)
    table.rows[0].height = Inches(0.55)
    for r in range(1, rows):
        table.rows[r].height = Inches(0.76)
    add_rect(s, 0.75, 5.5, 11.83, 1.15, LIGHT, rounded=True)
    add_text(s, 1.05, 5.72, 11.3, 0.8,
             [[("协作机制：", dict(size=14, bold=True, color=ORANGE)),
               ("双周项目同步会 + 共享文档实时更新 + 关键节点专项跟进，确保信息透明、问题不过夜。",
                dict(size=14, color=TEXT))]],
             line_spacing=1.35)
    return s


def slide_summary(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(s, "十、总结与展望", 12)
    add_bullets(s, [
        ("回顾全年：", "各项重点任务基本达成，成绩来之不易，凝聚了每一位成员的付出与智慧。"),
        ("正视不足：", "对标更高要求，仍存在短板与差距，将持续改进、把问题转化为成长机会。"),
        ("展望未来：", "围绕公司战略方向，聚焦核心目标，以更饱满的状态迎接新一年的挑战。"),
    ], 0.75, 1.7, 11.8, 2.6, size=16, gap=16)
    # 三个关键词
    kws = [("聚焦", "战略不偏移", BLUE), ("协同", "聚力共前行", ORANGE), ("突破", "再上新台阶", GREEN)]
    x0, y, w, h, gap = 0.75, 4.55, 3.83, 1.75, 0.27
    for i, (kw, desc, color) in enumerate(kws):
        x = x0 + i * (w + gap)
        add_rect(s, x, y, w, h, LIGHT, rounded=True)
        add_text(s, x, y + 0.28, w, 0.7,
                 [[(kw, dict(size=26, bold=True, color=color))]], align=PP_ALIGN.CENTER)
        add_text(s, x, y + 1.05, w, 0.5,
                 [[(desc, dict(size=13, color=TEXT))]], align=PP_ALIGN.CENTER)
    return s


def slide_thanks(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, BLUE_DARK)
    add_rect(s, -1.0, 5.0, 3.2, 3.2, BLUE, rounded=True, radius=0.5)
    add_rect(s, 11.6, -1.1, 3.2, 3.2, BLUE, rounded=True, radius=0.5)
    add_rect(s, SLIDE_W / 2 - 0.8, 2.35, 1.6, 0.1, ORANGE)
    add_text(s, 1.5, 2.65, 10.33, 1.2,
             [[("感谢聆听  ·  欢迎交流指导", dict(size=40, bold=True, color=WHITE))]],
             align=PP_ALIGN.CENTER)
    add_text(s, 1.5, 4.05, 10.33, 0.6,
             [[("THANKS FOR YOUR ATTENTION", dict(size=16, color=BLUE_LIGHT))]],
             align=PP_ALIGN.CENTER)
    add_text(s, 1.5, 5.15, 10.33, 0.5,
             [[("汇报人：XXX     |     联系方式：xxx@xxx.com     |     Q & A",
                dict(size=13, color=BLUE_LIGHT))]],
             align=PP_ALIGN.CENTER)
    return s


# ---------------------------------------------------------------- 主流程
def main():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    slide_cover(prs)
    slide_agenda(prs)
    slide_overview(prs)
    slide_keywork(prs)
    slide_data(prs)
    slide_highlights(prs)
    slide_problems(prs)
    slide_lessons(prs)
    slide_plan(prs)
    slide_timeline(prs)
    slide_resources(prs)
    slide_summary(prs)
    slide_thanks(prs)

    out = "工作汇报.pptx"
    prs.save(out)
    print("已生成：%s（共 %d 页）" % (out, len(prs.slides._sldIdLst)))


if __name__ == "__main__":
    main()
