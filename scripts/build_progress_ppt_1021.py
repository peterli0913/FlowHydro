"""
Build a 7-slide progress report (16:9) for 23 Sep – 21 Oct 2026.

Same chrome as the 0909 / 0923 decks (numbered tile, coloured header bar,
footer), but English only so the file can go to SW as-is.

Sources
-------
- Meeting Minutes Tracking Record 1020.xlsx (sheet 1020 vs 0922)
- 8 Oct technical session with Keith (Meeting_Summary_1008_Technical_Keith.md)
- Keith's 9 Oct email on multipoint RTDs; 14 Oct nozzle-size email
- 撬块及辅助设备投资评估表-2026.xlsx (CFCT re-estimate)
- R&D temperature-rise results (DN40 / DN65)

Budget arithmetic
-----------------
Approved budget GBP 351,591 covers skid & auxiliary equipment plus E&I and
control cabinets.  It was converted at 9.62 RMB/GBP; today's rate is 8.93.
Current estimate RMB 3,172,300 + 540,818.  The overrun at today's rate is
split on that same base: cost change (estimate vs budget at 9.62) and the
exchange-rate effect (budget at 9.62 vs 8.93).

Omitted
-------
Internal positions on the pulsation damper, who selected which item, and
anything not yet agreed with SW.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from PIL import ImageFont

FONT = "Segoe UI"

BODY = "1E2B3A"
SUB = "5B6B80"
FOOT = "9AA6B8"
LINE = "E3E8F0"
UP = "B91C1C"
DOWN = "15803D"

PERIOD = "Progress 2026-09-23 ~ 10-21"
TOTAL = "07"
OUT = "Progress_Report_1021.pptx"

BUDGET_GBP = 351_591
RATE_APPROVED = 9.62
RATE_NOW = 8.93
SKID_RMB = 3_172_300
EI_RMB = 540_818

# (category, quantity change, previous RMB, current RMB)
COST_LINES = [
    ("Main equipment", "10 → 10  (feed buffer tank removed, sampler added)",
     946_000, 1_049_400),
    ("Pumps", "3 → 1  (one three-head metering pump)", 260_000, 250_000),
    ("Flowmeters and controllers", "5 → 5", 315_000, 400_000),
    ("Temperature transmitters", "19 → 11", 136_800, 147_600),
    ("Pressure transmitters", "16 → 13", 112_000, 169_000),
    ("Level instruments", "4 gauges → 3 transmitters", 80_000, 68_500),
    ("Control valves", "5 → 5 pressure controllers / regulators",
     100_000, 240_000),
    ("On/off valves", "12 → 6", 96_000, 48_000),
]
SKID_PREVIOUS_RMB = 2_840_000


def rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _run(p, text, *, size=18, bold=False, color=BODY, italic=False):
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = rgb(color)
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
                                  Inches(7.2), Inches(0.3))
    _run(tb.text_frame.paragraphs[0],
         "Asymchem · Sandwich Site   |   Continuous Hydrogenation Skid",
         size=9, color=FOOT)
    tb = slide.shapes.add_textbox(Inches(7.6), Inches(7.05),
                                  Inches(5.1), Inches(0.3))
    p = tb.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    _run(p, f"{PERIOD}    ·    {no} / {TOTAL}", size=9, color=FOOT)


def add_header(slide, no, title, sub, color):
    box(slide, MSO_SHAPE.RECTANGLE, 0, 0, 13.333, 0.16, fill=color)
    box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, 0.38, 0.92, 0.92, fill=color)
    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.38),
                                  Inches(0.92), Inches(0.92))
    tf = tb.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, no, size=32, bold=True, color="FFFFFF")

    tb = slide.shapes.add_textbox(Inches(1.68), Inches(0.40),
                                  Inches(11.0), Inches(0.88))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    _run(tf.paragraphs[0], title, size=28, bold=True, color=color)
    p2 = tf.add_paragraph()
    p2.space_before = Pt(1)
    _run(p2, sub, size=14, color="8A97AB")

    box(slide, MSO_SHAPE.RECTANGLE, 0.55, 1.46, 12.23, 0.015, fill=LINE)


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


_LATIN = ImageFont.truetype(
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 100)
_LATIN_B = ImageFont.truetype(
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", 100)


def _nlines(text, size, width_in, bold=False):
    """Word-wrap count in PowerPoint points; 0.95 leaves room for Segoe UI."""
    font = _LATIN_B if bold else _LATIN
    width_pt = width_in * 72 * 0.95
    lines = 0
    for para in text.split("\n"):
        lines += 1
        used = 0.0
        for i, word in enumerate(para.split(" ")):
            w = font.getlength(word + " ") * size / 100.0
            if used + w > width_pt and used > 0:
                lines += 1
                used = w
            else:
                used += w
    return max(lines, 1)


def _block_h(head, subs, small):
    head_sz = 14 if small else 16
    sub_sz = 13
    h = 0.12
    h += _nlines(head, head_sz, 11.35, bold=not small) * (head_sz * 1.25) / 72
    for se in subs:
        h += 0.06
        h += _nlines("–  " + se, sub_sz, 11.0) * (sub_sz * 1.25) / 72
    h += 0.14
    return h


def add_block(slide, y, color, head, *, subs=None, small=False, gap=0.12):
    subs = subs or []
    h = _block_h(head, subs, small)
    box(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.55, y, 12.23, h,
        fill="F4F7FB" if not small else "F8FAFC", line=LINE, lw=1.0)
    box(slide, MSO_SHAPE.RECTANGLE, 0.55, y, 0.07, h, fill=color)

    tb = slide.shapes.add_textbox(Inches(0.82), Inches(y + 0.07),
                                  Inches(11.75), Inches(max(h - 0.10, 0.3)))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    _run(tf.paragraphs[0], head, size=14 if small else 16, bold=not small,
         color=BODY if not small else "3C4A5A")
    for se in subs:
        ps = tf.add_paragraph()
        ps.space_before = Pt(4)
        _run(ps, "–  ", size=13, color="94A3B8")
        _run(ps, se, size=13, color=SUB)
    return y + h + gap


def _cell(cell, text, *, size=12, bold=False, color=BODY, fill=None,
          align=PP_ALIGN.LEFT):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    cell.margin_left = cell.margin_right = Inches(0.08)
    cell.margin_top = cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    _run(p, text, size=size, bold=bold, color=color)
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill)
    else:
        cell.fill.background()


def add_table(slide, x, y, widths, row_h, rows, *, header_fill, styles=None):
    """rows[0] is the header.  styles maps row index to a dict of options."""
    styles = styles or {}
    shape = slide.shapes.add_table(len(rows), len(widths), Inches(x),
                                   Inches(y), Inches(sum(widths)),
                                   Inches(row_h * len(rows)))
    tbl = shape.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    style_id = tblPr.find(qn("a:tableStyleId"))
    if style_id is not None:
        style_id.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"
    for i, w in enumerate(widths):
        tbl.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        tbl.rows[r].height = Inches(row_h)
        opt = styles.get(r, {})
        for c, val in enumerate(row):
            align = PP_ALIGN.LEFT if c < opt.get("left_cols", 2) else PP_ALIGN.RIGHT
            if r == 0:
                _cell(tbl.cell(r, c), val, size=12, bold=True, color="FFFFFF",
                      fill=header_fill, align=align)
                continue
            color = opt.get("colors", {}).get(c, BODY)
            _cell(tbl.cell(r, c), val, size=opt.get("size", 12),
                  bold=opt.get("bold", False), color=color,
                  fill=opt.get("fill", "FFFFFF" if r % 2 else "F6F8FB"),
                  align=align)
    return y + row_h * len(rows)


def k(v):
    return f"{v / 1000:,.1f}"


def signed_k(v):
    return ("+" if v > 0 else "−") + f"{abs(v) / 1000:,.1f}"


def slide_01(prs):
    color = "2563EB"
    s = blank_slide(prs)
    add_header(s, "01", "Overall Progress", "23 September – 21 October 2026",
               color)
    y = 1.66
    y = add_block(
        s, y, color,
        "After several review rounds with SW, the positions of the key "
        "equipment are largely settled.",
        subs=[
            "Three modules; the front module pulls out for maintenance.  "
            "BT01, BT04, the separator and the condenser sit forward; the "
            "separator is lowered.",
            "Reviews this period: 8, 15 and 20 October, on models issued "
            "by CFCT.",
        ])
    y = add_block(
        s, y, color,
        "We are now going through the model at true positions: nozzle "
        "orientation, valve positions, pipe routing, and replacing "
        "instruments that take too much space.")
    y = add_block(
        s, y, color,
        "Two decisions in this report: the re-estimated equipment cost "
        "against the approved budget, and DN65 for Reactor 1 (R24-CR01).")
    y = add_block(
        s, y, color,
        "Action register (20 Oct): 63 closed, 47 pending, 28 open.",
        subs=[
            "Three items added: multipoint reactor temperature, metering-pump "
            "pulsation damper, and the position of the split control cabinets.",
        ], gap=0.08)
    add_footer(s, "01")
    return y


def slide_02(prs):
    color = "16A34A"
    s = blank_slide(prs)
    add_header(s, "02", "Layout — Settled and Open Points",
               "From the 8, 15 and 20 October reviews", color)
    y = 1.62
    y = add_block(
        s, y, color,
        "TCU tie-in: the site TCU lines can shift about 150 mm to meet the "
        "skid nozzles (agreed 15 Oct).",
        subs=[
            "Flanged wall plate; the ten lines may stack in two layers after "
            "the wall.  The Reactor 3 bottom HTF now turns toward the wall "
            "riser.",
        ], gap=0.10)
    y = add_block(
        s, y, color,
        "Split control cabinets: three cabinets / junction boxes are placed "
        "in the 3D at their current sizes (20 Oct).",
        subs=[
            "Main IS cabinet: 5 AI, 3 DI, 3 DO; minimum 530 × 450 × 230 mm, "
            "down from 950 × 450 × 350 mm.",
            "New IS cabinet on module 2: 3 RTD, 5 AI, 1 DI; minimum "
            "470 × 450 × 230 mm.",
        ], gap=0.10)
    y = add_block(
        s, y, color,
        "Remote IS I/O panel: SW budget price about £20k, 10–12 weeks.",
        subs=[
            "Module list and Profinet / Profibus to be aligned first; then "
            "the quote is firmed up and the purchase route (UK or China) "
            "agreed.",
        ], gap=0.10)
    y = add_block(
        s, y, "D97706",
        "Open: SE01 / CU01 — wall side or end face of module 3.  Drain and "
        "CIP valves at the front.",
        subs=[
            "The end-face option clashes with the Reactor 3 HTF.  Both "
            "points are in the 20 Oct model for SW review.",
            "Valves within 300–400 mm of the end stay manual; beyond that, "
            "compact actuated valves.",
        ], gap=0.08)
    add_footer(s, "02")
    return y


def slide_03(prs):
    color = "16A34A"
    s = blank_slide(prs)
    add_header(s, "03", "Detailed Review — Pump and Instruments",
               "Smaller, higher-specification items where space is tight",
               color)
    y = 1.62
    y = add_block(
        s, y, color,
        "Metering pump: a three-head pump with inlet and outlet manifolds "
        "and no discharge damper replaces the single head with a pulsation "
        "damper.",
        subs=[
            "Addresses SW's concern: the damper is a dead leg, and a "
            "diaphragm failure could go undetected.",
            "Vendor drawing received 20 Oct and placed in the 3D; pulsation "
            "curves supplied.  20 L/h, 7.5 MPa discharge, GMP.",
            "CE marking is sufficient; no foreign motor brand is required "
            "(15 Oct).",
            "Vendor: flow can be calibrated on site for each material; a "
            "flowmeter is needed only for remote flow indication.  To be "
            "confirmed by process.",
        ], gap=0.12)
    y = add_block(
        s, y, color,
        "Reactor temperatures: SW's multipoint RTD concept is feasible — no "
        "local transmitters or displays; readings on the HMI.",
        subs=[
            "E+H iTherm MultiSens Slim TMS21 on flexible conduit, wired "
            "direct to Siemens RTD module 6ES7134-7SD51-0AB0 in the ET200iSP "
            "panel.",
            "E+H has confirmed supply; Siemens has confirmed the RTD module.  "
            "Parameter review with E+H this week.",
        ], gap=0.12)
    y = add_block(
        s, y, color,
        "Wider instrument selection: level gauges become guided-wave radar "
        "transmitters; on/off valves drop from 12 to 6; pressure control "
        "moves to electronic controllers and regulators.",
        gap=0.08)
    add_footer(s, "03")
    return y


def slide_04(prs):
    color = "DC2626"
    s = blank_slide(prs)
    add_header(s, "04", "Decision 1 — Equipment Cost Re-estimate",
               "Skid and auxiliary equipment, CFCT estimate, RMB thousand",
               color)
    y = 1.62
    tb = s.shapes.add_textbox(Inches(0.55), Inches(y), Inches(12.23),
                              Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    _run(tf.paragraphs[0],
         "Because of the skid space, components were re-selected for better "
         "performance and a smaller size.  Fewer items in several "
         "categories; higher unit prices.",
         size=14, color=SUB)
    y += 0.55

    rows = [["Category", "Quantity", "Previous", "Current", "Change"]]
    styles = {}
    for name, qty, old, new in COST_LINES:
        d = new - old
        rows.append([name, qty, k(old), k(new), signed_k(d)])
        styles[len(rows) - 1] = {"colors": {4: UP if d > 0 else DOWN}}
    old_sum = sum(o for _, _, o, _ in COST_LINES)
    new_sum = sum(n for _, _, _, n in COST_LINES)
    rows.append(["Subtotal, eight categories", "", k(old_sum), k(new_sum),
                 signed_k(new_sum - old_sum)])
    styles[len(rows) - 1] = {"bold": True, "fill": "EEF2F7",
                             "colors": {4: UP}}
    rows.append(["Skid and auxiliary equipment, total", "",
                 "≈ " + k(SKID_PREVIOUS_RMB), k(SKID_RMB),
                 "≈ +" + f"{round((SKID_RMB - SKID_PREVIOUS_RMB) / 10_000) * 10:,}"])
    styles[len(rows) - 1] = {"bold": True, "fill": "FDECEC",
                             "colors": {4: UP}}
    y = add_table(s, 0.55, y, [3.35, 4.68, 1.4, 1.4, 1.4], 0.37, rows,
                  header_fill="7F1D1D", styles=styles)

    tb = s.shapes.add_textbox(Inches(0.55), Inches(y + 0.12), Inches(12.23),
                              Inches(0.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    _run(tf.paragraphs[0],
         "Modular assembly, CE certification and freight are not yet "
         "re-estimated and are carried at the previous figures.",
         size=12, italic=True, color=SUB)
    add_footer(s, "04")
    return y + 0.6


def slide_05(prs):
    color = "DC2626"
    s = blank_slide(prs)
    add_header(s, "05", "Decision 1 — Against the Approved Budget",
               "Skid and auxiliary equipment + E&I and control cabinets",
               color)
    budget_then = BUDGET_GBP * RATE_APPROVED
    budget_now = BUDGET_GBP * RATE_NOW
    estimate = SKID_RMB + EI_RMB
    over = estimate - budget_now
    cost_part = estimate - budget_then
    fx_part = budget_then - budget_now
    gbp_est = estimate / RATE_NOW

    y = 1.62
    rows = [
        ["", "RMB thousand", "GBP"],
        ["Approved budget", "", f"{BUDGET_GBP:,.0f}"],
        [f"  at approval rate, 1 GBP = {RATE_APPROVED} RMB", k(budget_then), ""],
        [f"  at today's rate, 1 GBP = {RATE_NOW} RMB", k(budget_now), ""],
        ["Current estimate — skid and auxiliary equipment", k(SKID_RMB), ""],
        ["Current estimate — E&I and control cabinets", k(EI_RMB), ""],
        ["Current estimate — total", k(estimate), f"{gbp_est:,.0f}"],
        [f"Over budget at today's rate  ({over / budget_now:.1%})",
         k(over), f"{gbp_est - BUDGET_GBP:,.0f}"],
        [f"  of which cost change  ({cost_part / budget_now * 100:.1f} pts)",
         k(cost_part), ""],
        [f"  of which exchange rate  ({fx_part / budget_now * 100:.1f} pts)",
         k(fx_part), ""],
    ]
    styles = {
        6: {"bold": True, "fill": "EEF2F7", "left_cols": 1},
        7: {"bold": True, "fill": "FDECEC", "left_cols": 1,
            "colors": {0: UP, 1: UP, 2: UP}},
    }
    for r in range(1, len(rows)):
        styles.setdefault(r, {})["left_cols"] = 1
    y = add_table(s, 0.55, y, [7.83, 2.2, 2.2], 0.39, rows,
                  header_fill="7F1D1D", styles=styles)

    y += 0.22
    y = add_block(
        s, y, color,
        f"At today's rate the estimate is {over / budget_now:.1%} over the "
        f"approved budget.  {fx_part / budget_now * 100:.1f} points of that "
        f"is the exchange rate.  Had the rate stayed at {RATE_APPROVED}, the "
        f"estimate would be {cost_part / budget_then:.1%} over.",
        subs=[
            "Assembly, certification and freight are still to be "
            "re-estimated, so the total may move.",
        ], gap=0.08)
    add_footer(s, "05")
    return y


def slide_06(prs):
    color = "7C3AED"
    s = blank_slide(prs)
    add_header(s, "06", "Decision 2 — Reactor 1 (R24-CR01) at DN65",
               "Preliminary decision, 15 October", color)
    y = 1.62
    y = add_block(
        s, y, color,
        "R24-CR01 goes to DN65 — the same diameter as Reactors 2 and 3.",
        subs=[
            "Datasheet basis: 3.5 L and 0.23 m² heat-transfer area "
            "(DN80 was 5.1 L and 0.27 m²).",
        ], gap=0.10)
    y = add_block(
        s, y, color,
        "Catalyst charging decides it.",
        subs=[
            "DN50: the largest catalyst charge and discharge nozzle is 3/4\"; "
            "the smallest bottom discharge valve is 3/4\".",
            "SW view (8 Oct): 1/2\" would not take the catalyst; 3/4\" would "
            "be difficult.",
            "DN65 keeps the 1\" charge and discharge nozzles already designed.",
        ], gap=0.10)
    y = add_block(
        s, y, color,
        "Temperature: R&D measured about 5–7 °C rise at DN40 and a "
        "15–18 °C hot spot at DN65.  There is no DN50 data.",
        subs=[
            "At DN65 the hot spot can be controlled by diluting the catalyst "
            "bed with inert material at loading.",
        ], gap=0.10)
    y = add_block(
        s, y, color,
        "Fits the skid height; DN50 would need a longer column for the same "
        "volume.  PFD and mass balance are unchanged; the DN65 relief "
        "calculation already exists.",
        gap=0.08)
    add_footer(s, "06")
    return y


def slide_07(prs):
    color = "0E9AA7"
    s = blank_slide(prs)
    add_header(s, "07", "Next Steps", "Owners and open items", color)
    y = 1.62
    steps = [
        "SW (Keith) — review the 20 Oct model: SE01 / CU01 position, drain "
        "and CIP valves, cabinet positions, three-head pump.",
        "CFCT (Fan Shuangshuang) — freeze nozzle orientation and HTF routes, "
        "then release the main-equipment drawings for enquiry.",
        "CIMT (Meng Dezhi) — align the Siemens IS I/O module list and "
        "Profinet / Profibus; agree UK or China purchase.",
        "Instruments (Gao Yu) — E+H parameter review; apply the multipoint "
        "RTD concept to other temperature points where it suits.",
        "Process (Hu Chaoqun) — confirm feed-flow measurement against the "
        "three-head pulsation curves.",
        "R&D (Dr Feng) — catalyst loading and bed-dilution procedure for "
        "R24-CR01 at DN65.",
        "CFCT — re-estimate assembly, certification and freight; update the "
        "budget comparison.",
    ]
    for text in steps:
        y = add_block(s, y, color, text, small=True, gap=0.09)
    add_footer(s, "07")
    return y


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    ys = [f(prs) for f in (slide_01, slide_02, slide_03, slide_04, slide_05,
                           slide_06, slide_07)]
    prs.save(OUT)
    print("Saved:", OUT, "end-y:", [round(y, 2) for y in ys])
    if any(y > 6.95 for y in ys):
        raise SystemExit("A slide runs into the footer. Shorten copy or spacing.")


if __name__ == "__main__":
    main()
