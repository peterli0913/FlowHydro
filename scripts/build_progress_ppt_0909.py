"""
Build a 5-slide progress report (16:9) for 27 Aug – 9 Sep 2026.

Same chrome as the 0715 / 0729 / 0811 / 0826 decks (numbered tile, English
primary, small Chinese secondary, coloured header bar).  Written so the file
can be sent to SW / Keith as-is.

Sources
-------
- 8 Sep technical session with Keith (gasket, PIC-028, 3D / cabinet direction)
- 9 Sep working-programme alignment + 0909 timeline workbook
- Sandwich visit agenda (3D + HAZOP), proposed 12–22 October

What is deliberately omitted
----------------------------
Internal discussion of why a window moved, delay attribution, leadership
briefing points, and any other content that is not appropriate in a UK-facing
progress note.  Downstream bars that were disturbed during the live 9 Sep
edit are not restated as new end dates; only the near-term freeze points
agreed that morning are shown as dates.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

FONT_EN = "Segoe UI"
FONT_ZH = "Microsoft YaHei"

BODY = "1E2B3A"
SUB = "8A97AB"
FOOT = "9AA6B8"
LINE = "E3E8F0"

PERIOD = "Progress 2026-08-27 ~ 09-09"
TOTAL = "05"
OUT = "Progress_Report_0909.pptx"


def rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _run(p, text, *, font=FONT_EN, size=18, bold=False, color=BODY, italic=False):
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = rgb(color)
    rPr = r._r.get_or_add_rPr()
    rPr.append(rPr.makeelement(qn("a:ea"), {"typeface": FONT_ZH}))
    return r


def box(slide, kind, x, y, w, h, *, fill=None, line=None, lw=1.0):
    sh = slide.shapes.add_shape(kind, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def add_footer(slide, no):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.05),
                                  Inches(6.6), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    _run(p, "Asymchem · Sandwich Site   |   Continuous Hydrogenation Skid",
         size=9, color=FOOT)
    tb = slide.shapes.add_textbox(Inches(7.2), Inches(7.05),
                                  Inches(4.8), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    _run(p, f"{PERIOD}    ·    {no} / {TOTAL}", size=9, color=FOOT)


def add_header(slide, no, title, sub, color):
    box(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 0.16, fill=color)
    box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 0.42, 1.0, 1.0, fill=color)
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.42),
                                  Inches(1.0), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, no, size=34, bold=True, color="FFFFFF")

    tb = slide.shapes.add_textbox(Inches(1.85), Inches(0.48),
                                  Inches(10.8), Inches(0.95))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    _run(p, title, size=28, bold=True, color=color)
    p2 = tf.add_paragraph()
    p2.space_before = Pt(2)
    _run(p2, sub, font=FONT_ZH, size=14, color=SUB)

    box(slide, MSO_SHAPE.RECTANGLE, 0.6, 1.58, 12.13, 0.02, fill=LINE)


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_bullet_block(slide, y, color, en, cn, *, small=False, subs=None, gap=0.22):
    subs = subs or []
    en_sz = 14 if small else 16.5
    cn_sz = 10.5 if small else 11.5
    dot_sz = 14 if small else 17
    h_box = 0.52 + 0.34 * len(subs)
    tb = slide.shapes.add_textbox(Inches(0.7), Inches(y),
                                  Inches(12.0), Inches(h_box))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    _run(p, "●  ", size=dot_sz, bold=True, color=(SUB if small else color))
    _run(p, en, size=en_sz, bold=(not small),
         color=(BODY if not small else "5A6675"))
    p2 = tf.add_paragraph()
    p2.space_before = Pt(1)
    _run(p2, "     " + cn, font=FONT_ZH, size=cn_sz, color=SUB)
    for se, sc in subs:
        ps = tf.add_paragraph()
        ps.space_before = Pt(3)
        _run(ps, "        –  ", size=12, color="B7C0CF")
        _run(ps, se, size=13, color="3C4A5A")
        pc = tf.add_paragraph()
        _run(pc, "             " + sc, font=FONT_ZH, size=10, color=SUB)
    return y + 0.48 + 0.32 * len(subs) + gap


def add_issue_card(slide, y, h, en, cn, subs):
    box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, y, 12.13, h,
        fill="FFF8EC", line="D97706", lw=1.5)
    tb = slide.shapes.add_textbox(Inches(0.85), Inches(y + 0.10),
                                  Inches(11.65), Inches(h - 0.18))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    _run(p, "● ", size=16, bold=True, color="D97706")
    _run(p, en, size=15, bold=True, color="92400E")
    p2 = tf.add_paragraph()
    p2.space_before = Pt(2)
    _run(p2, cn, font=FONT_ZH, size=11, color="B4763A")
    for se, sc in subs:
        ps = tf.add_paragraph()
        ps.space_before = Pt(3)
        _run(ps, se, size=12, color="7C5518")
        pc = tf.add_paragraph()
        _run(pc, sc, font=FONT_ZH, size=10, color="B4763A")
    return y + h + 0.16


def slide_01(prs):
    color = "2563EB"
    s = blank_slide(prs)
    add_header(s, "01", "Overall Progress", "整体推进", color)

    y = 1.78
    y = add_bullet_block(
        s, y, color,
        "Technical session with Keith on 8 September — two tracker items "
        "closed in the meeting",
        "9 月 8 日与 Keith 召开技术会，两项跟踪条目在会上关闭",
        subs=[
            ("High-pressure gasket accepted as a 316L metallic ring; "
             "tantalum is withdrawn",
             "高压垫片接受 316L 金属环方案，钽材路线取消"),
            ("PIC-028 stays Bronkhorst if the mechanical seal holds at 270 °C",
             "PIC-028 继续用 Bronkhorst，前提是 270 ℃ 时机械密封不外漏"),
        ],
        gap=0.16,
    )
    y = add_bullet_block(
        s, y, color,
        "3D layout is not frozen — the next model follows the direction "
        "agreed on 8 September",
        "3D 布置尚未冻结，下一版模型按 9 月 8 日商定的方向更新",
        subs=[
            ("Compact split-skid, 2.5 m shipping height, withdrawable front module",
             "紧凑分撬、运输高度 2.5 m、前撬可拉出，不留中间通道"),
        ],
        gap=0.16,
    )
    y = add_bullet_block(
        s, y, color,
        "Working programme aligned on 9 September with the Sandwich visit — "
        "layout freeze and HAZOP sit in the same trip",
        "9 月 9 日将工作计划与赴 Sandwich 行程对齐：布置冻结与 HAZOP 安排在同一次现场",
        gap=0.20,
    )
    y = add_bullet_block(
        s, y, color,
        "Draft visit agenda issued — four blocks, eight working days, "
        "proposed 12–22 October",
        "现场议程草案已发出：四块内容、八个工作日，建议 10 月 12–22 日",
        subs=[
            ("Walkdown, 3D, manuals / control / sequence, then HAZOP — "
             "dates for Keith and Claire to confirm",
             "勘察、3D、手册 / 控制 / 顺控，然后 HAZOP——日期待 Keith、Claire 确认"),
        ],
        gap=0.12,
    )
    add_footer(s, "01")


def slide_02(prs):
    color = "16A34A"
    s = blank_slide(prs)
    add_header(s, "02", "Technical Progress", "关键技术进展", color)

    y = 1.76
    y = add_bullet_block(
        s, y, color,
        "Gasket — closed.  316L metallic ring; hardness written into the spec",
        "垫片已关闭：316L 金属环，硬度要求写入规格",
        subs=[
            ("Gasket < 150 HB, flange > 160 HB; mill certificates plus a "
             "vendor re-check before assembly",
             "垫片 < 150 HB，法兰 > 160 HB；厂家硬度报告，组装前复验"),
        ],
        gap=0.12,
    )
    y = add_bullet_block(
        s, y, color,
        "PIC-028 — closed.  Electronics may fail at 270 °C; the seal must not",
        "PIC-028 已关闭：270 ℃ 时电子件允许失效，密封不得外漏",
        subs=[
            ("PSV takes the relief; valve is replaced afterwards.  FFKM on "
             "the datasheet and the order; same check for FCV-004 if needed",
             "泄放由安全阀承担，事后换阀。FFKM 写入数据单与合同；FCV-004 视情况同样处理"),
        ],
        gap=0.12,
    )
    y = add_bullet_block(
        s, y, color,
        "Layout direction for the next 3D revision",
        "下一版 3D 的布置方向",
        subs=[
            ("Ten TCU pipes through a wall plate (China plate / UK slot); "
             "last elbow fitted on site",
             "10 根 TCU 管经穿墙板（中方出板、英方开槽）；最后弯头现场安装"),
            ("Separator may be lowered after the CIP change; new height "
             "goes to process",
             "清洗方案调整后分离器可降低，新高度交工艺核算减压后余压"),
            ("Cabinet set back and rectangular; vents over 2.5 m stripped "
             "for shipping",
             "控制柜后移并改为长方形；超 2.5 m 放散管运输拆下、现场复原"),
        ],
        gap=0.22,
    )
    add_issue_card(
        s, y, 1.12,
        "Still open — model being redrawn; one instrument item not reached "
        "on 8 September",
        "仍开放：模型按新方向重排中；9 月 8 日未轮到一项仪表问题",
        [
            ("Whether the PSV can stay at the DN65 size if the reactor is "
             "later confirmed as DN50.  DN65 remains the working basis "
             "pending the September trial.",
             "若反应器最终确认为 DN50，安全阀能否维持 DN65 口径。"
             "在 9 月试验数据到位前，工作基准仍为 DN65。"),
        ],
    )
    add_footer(s, "02")


def slide_03(prs):
    color = "7C3AED"
    s = blank_slide(prs)
    add_header(s, "03", "Working Programme", "工作计划", color)

    tb = s.shapes.add_textbox(Inches(0.65), Inches(1.70),
                                  Inches(12.1), Inches(0.42))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    _run(p, "Near-term freeze points aligned on 9 September.  Downstream "
            "work keeps the established sequence after the layout is frozen.",
         size=13.5, color="3C4A5A")
    p2 = tf.add_paragraph()
    _run(p2, "9 月 9 日对齐的近期冻结节点。布置冻结之后，后续仍按既定顺序推进。",
         font=FONT_ZH, size=11, color=SUB)

    rows = [
        ("End Sep 2026", "9 月底",
         "Draft FDS / SDS / HDS, control description and operating "
         "instructions as HAZOP input.  Finals after FAT.",
         "FDS / SDS / HDS、控制说明与操作手册出初稿，作为 HAZOP 输入；定稿在 FAT 之后。"),
        ("End Sep 2026", "9 月底",
         "Reactor specification confirmed for enquiry.  Main-equipment "
         "procurement can then start — it does not wait for 3D freeze.",
         "反应柱规格确认后即可发起主设备采购，不必等待 3D 冻结。"),
        ("12–16 Oct 2026", "10/12–16",
         "Sandwich: walkdown, 3D layout, operating manual, control "
         "philosophy and sequence.  Proposed — for confirmation.",
         "赴 Sandwich：现场勘察、3D、操作手册、控制说明与顺控。建议档期，待确认。"),
        ("20–22 Oct 2026", "10/20–22",
         "Sandwich HAZOP on the layout frozen that week.  Proposed — "
         "Keith and Claire to confirm.",
         "在当周冻结的布置上做 HAZOP。建议档期，待 Keith、Claire 确认。"),
        ("Oct W2 2026", "10 月第 2 周",
         "Main-equipment enquiry after the 1–7 October holiday in China, "
         "once the reactor specification is in hand.",
         "中国 10 月 1–7 日假期之后，规格到位即可进入主设备询价。"),
        ("~ Nov W3 2026", "约 11 月第 3 周",
         "Vendor information confirmation — after manufacturer selection, "
         "not before.  About four weeks.",
         "厂家信息确认：在选定厂家之后，而不是之前。约四周闭环。"),
    ]

    y = 2.18
    for when_en, when_zh, en, cn in rows:
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, y, 12.13, 0.72,
            fill="F7F4FF", line="DDD6FE", lw=1.0)
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.72, y + 0.12, 2.15, 0.48,
            fill=color)
        tb = s.shapes.add_textbox(Inches(0.72), Inches(y + 0.12),
                                      Inches(2.15), Inches(0.48))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _run(p, when_en, size=11, bold=True, color="FFFFFF")
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        _run(p2, when_zh, font=FONT_ZH, size=9, color="EDE9FE")

        tb = s.shapes.add_textbox(Inches(3.05), Inches(y + 0.06),
                                      Inches(9.50), Inches(0.60))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(0)
        p = tf.paragraphs[0]
        _run(p, en, size=12, color=BODY)
        p2 = tf.add_paragraph()
        _run(p2, cn, font=FONT_ZH, size=10, color=SUB)
        y += 0.76

    add_footer(s, "03")


def slide_04(prs):
    color = "0E9AA7"
    s = blank_slide(prs)
    add_header(s, "04", "Sandwich Visit", "赴 Sandwich 现场", color)

    box(s, MSO_SHAPE.ROUNDED_RECTANGLE, 0.6, 1.74, 12.13, 0.78,
        fill="E7F6F7", line=color, lw=1.2)
    tb = s.shapes.add_textbox(Inches(0.85), Inches(1.80),
                                  Inches(11.65), Inches(0.66))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    _run(p, "Purpose — freeze the layout against the room, then run HAZOP "
            "once on that frozen basis.",
         size=15, bold=True, color="0F5E67")
    p2 = tf.add_paragraph()
    p2.space_before = Pt(2)
    _run(p2, "目的：对照现场空间把布置冻结，再在该基准上只做一次 HAZOP。",
         font=FONT_ZH, size=12, color="4A8A90")

    cards = [
        ("1", "Site walkdown", "现场勘察", "0.5 day",
         "Skid boundary, stairwell, emergency shower, HTF wall penetration, "
         "goods-lift and delivery route — measured and jointly confirmed."),
        ("2", "3D model", "3D 模型", "2.5 days",
         "Equipment sequence, maintenance access, control-cabinet envelope, "
         "and the split-skid boundaries, including height for the goods lift."),
        ("3", "Manuals & control", "手册与控制", "2 days",
         "Operating manual, control philosophy and sequence walked against "
         "the layout fixed earlier in the week."),
        ("4", "HAZOP", "HAZOP", "3 days",
         "Study on the frozen P&ID and layout.  Actions and owners recorded "
         "in the room; no node left for a follow-up study."),
    ]
    positions = [(0.6, 2.68), (6.72, 2.68), (0.6, 4.48), (6.72, 4.48)]
    for (x, y), (no, en, zh, dur, blurb) in zip(positions, cards):
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, 5.95, 1.68,
            fill="F4FBFB", line="C5E4E7", lw=1.0)
        box(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.14, y + 0.16, 0.42, 0.42,
            fill=color)
        tb = s.shapes.add_textbox(Inches(x + 0.14), Inches(y + 0.16),
                                      Inches(0.42), Inches(0.42))
        tf = tb.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Emu(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _run(p, no, size=16, bold=True, color="FFFFFF")

        tb = s.shapes.add_textbox(Inches(x + 0.68), Inches(y + 0.14),
                                      Inches(5.05), Inches(0.52))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(0)
        p = tf.paragraphs[0]
        _run(p, en, size=15, bold=True, color="0F5E67")
        _run(p, "   " + dur, size=11, color=color)
        p2 = tf.add_paragraph()
        _run(p2, zh, font=FONT_ZH, size=11, color=SUB)

        tb = s.shapes.add_textbox(Inches(x + 0.18), Inches(y + 0.72),
                                      Inches(5.60), Inches(0.86))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(0)
        p = tf.paragraphs[0]
        _run(p, blurb, size=12, color="3C4A5A")

    tb = s.shapes.add_textbox(Inches(0.65), Inches(6.28),
                                  Inches(12.1), Inches(0.64))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    p = tf.paragraphs[0]
    _run(p, "Proposed   ", size=12, bold=True, color=color)
    _run(p, "12–16 Oct (walkdown / 3D / manuals)   ·   20–22 Oct (HAZOP)",
         size=12, color=BODY)
    p2 = tf.add_paragraph()
    p2.space_before = Pt(2)
    _run(p2, "Attendees  ", size=12, bold=True, color=color)
    _run(p2, "Keith · Li Tao · Fan Shuangshuang · Hu Chaoqun · Meng Dezhi",
         size=12, color=BODY)
    _run(p2, "    Agenda: ", size=11, color=SUB)
    _run(p2, "Agenda of the visit to Sandwich (3D & HAZOP).xlsx",
         size=11, italic=True, color="5A6675")

    add_footer(s, "04")


def slide_05(prs):
    color = "0E9AA7"
    s = blank_slide(prs)
    add_header(s, "05", "Next Steps", "下一步", color)

    steps = [
        ("Fan — redraw the 3D model to the 8 September direction; pass the "
         "new separator height to process",
         "范双双：按 9 月 8 日方向重排 3D，并将分离器新高度交给工艺"),
        ("Meng — send Keith the rectangular, reachable cabinet envelope "
         "for enquiry",
         "孟德智：把长方形、够得到的控制柜目标尺寸发给 Keith 询价"),
        ("Zhao / Gao — write gasket hardness and PIC-028 FFKM into the "
         "datasheets (FCV-004 if needed)",
         "赵子亮 / 高宇：垫片硬度与 PIC-028 的 FFKM 写入数据单（必要时含 FCV-004）"),
        ("Issue draft FDS, control description and operating instructions "
         "by the end of September, as HAZOP input",
         "9 月底发出 FDS、控制说明与操作手册初稿，作为 HAZOP 输入"),
        ("Raise the main-equipment enquiry once the reactor specification "
         "is confirmed — do not wait for 3D freeze",
         "反应柱规格确认后即发起主设备询价，不因 3D 未冻结而等待"),
        ("Confirm 12–22 October with Keith and Claire; start visa / travel "
         "once the dates are firm",
         "与 Keith、Claire 确认 10 月 12–22 日，日期确定后启动签证与差旅"),
        ("Close remaining tracker lines — cleaning, manuals, control "
         "description — at the visit or the next session",
         "剩余跟踪条目（清洗、手册、控制说明）在现场或下一次技术会上关闭"),
    ]

    y = 1.78
    for en, cn in steps:
        y = add_bullet_block(s, y, color, en, cn, gap=0.10)

    add_footer(s, "05")


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide_01(prs)
    slide_02(prs)
    slide_03(prs)
    slide_04(prs)
    slide_05(prs)
    prs.save(OUT)
    print("Saved:", OUT, "slides:", len(prs.slides))


if __name__ == "__main__":
    main()
