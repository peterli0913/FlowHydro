"""
Save a one-page Chinese agenda from
`0909-Agenda of the visit to Sandwich (3D & HAZOP).xlsx`.

The source file is not modified.  Table chrome (title / header / day banner /
merged date column) is kept.  Person in Charge is replaced by 现场参与人员
and 线上参与人员; a 备注 column is added at the end.  All headers and body
text are Chinese only.

Clare attends site walkdown plus operability / safety sessions only.
"""

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

SRC = "0909-Agenda of the visit to Sandwich (3D & HAZOP).xlsx"
DST = "0909-赴Sandwich现场议程（3D与HAZOP）.xlsx"
SHEET = "现场议程"

NAVY = "2E75B5"
NAVY_DK = "1F4E79"
BANNER = "BDD7EE"
ROW_A = "FFFFFF"
ROW_B = "F3F8FC"
LUNCH_BG = "FFF8EC"
LINE = "B8C7D9"
WHITE = "FFFFFF"
INK = "1E2B3A"
MUTED = "5B6B7C"

ONSITE = "Keith、Clare\n林浩淦、范双双、胡超群"
ONSITE_NO_CLARE = "Keith、林浩淦\n范双双、胡超群"
ONLINE = "高少峰、孟德智\n高宇、赵子亮"

# Clare: site walk + operability / safety only
CLARE_WALK = "Clare 参加（现场勘察）"
CLARE_OPS = "Clare 参加（操作性）"
CLARE_SAFE = "Clare 参加（安全）"
CLARE_OPS_SAFE = "Clare 参加（操作性、安全）"
CLARE_NO = "Clare 不参加"

HEADERS = [
    "序号", "日期", "时间", "行程",
    "现场参与人员", "线上参与人员",
    "详细议题", "目标", "备注",
]


def font(size=9, bold=False, color=INK, name="微软雅黑"):
    return Font(name=name, size=size, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def thin_border():
    s = Side(style="thin", color=LINE)
    return Border(left=s, right=s, top=s, bottom=s)


def style_cell(cell, *, value=None, fnt=None, fl=None, al=None):
    if value is not None:
        cell.value = value
    cell.font = fnt or font()
    cell.fill = fl or fill(ROW_A)
    cell.alignment = al or align()
    cell.border = thin_border()


def people(clare):
    if clare:
        return ONSITE, ONLINE
    return ONSITE_NO_CLARE, ONLINE


def main():
    if not Path(SRC).exists():
        raise FileNotFoundError(SRC)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET

    banner_fill = fill(BANNER)
    header_fill = fill(NAVY)

    last = len(HEADERS)
    last_letter = get_column_letter(last)

    widths = {
        "A": 5.2, "B": 11.2, "C": 11.5, "D": 16.5,
        "E": 16.0, "F": 15.2, "G": 36.0, "H": 28.0, "I": 16.5,
    }
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ----- title -------------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
    t = ws["A1"]
    t.value = "赴 Sandwich 现场议程  ·  3D 布置与 HAZOP"
    t.font = Font(name="微软雅黑", size=15, bold=True, color=WHITE)
    t.fill = fill(NAVY_DK)
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.border = thin_border()
    for col in range(2, last + 1):
        c = ws.cell(1, col)
        c.fill = fill(NAVY_DK)
        c.border = thin_border()
    ws.row_dimensions[1].height = 24

    # subtitle / legend
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last)
    sub = ws["A2"]
    sub.value = (
        "现场：Keith、Clare、林浩淦、范双双、胡超群"
        "（Clare 仅参加现场勘察及操作性、安全相关场次）"
        "     线上：高少峰、孟德智、高宇、赵子亮"
        "     建议档期：10月12–16日（布置 / 手册 / 控制），10月20–22日（HAZOP）"
    )
    sub.font = Font(name="微软雅黑", size=8.5, color=NAVY_DK)
    sub.fill = fill("E8F1F8")
    sub.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sub.border = thin_border()
    for col in range(2, last + 1):
        c = ws.cell(2, col)
        c.fill = fill("E8F1F8")
        c.border = thin_border()
    ws.row_dimensions[2].height = 20

    # ----- header ------------------------------------------------------
    for col, text in enumerate(HEADERS, start=1):
        c = ws.cell(3, col, value=text)
        c.font = Font(name="微软雅黑", size=9, bold=True, color=WHITE)
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border()
    ws.row_dimensions[3].height = 20

    # (banner_title, date, date_merge_first_content_row_offset, rows)
    # each content row: (time, itinerary, detail, objective, clare_note)
    # clare_note is None for lunch; str for 备注; False means Clare 不参加
    DAYS = [
        ("现场勘察与 3D 模型评审", "10月12日（一）", [
            ("09:00–12:00", "1. 现场勘察\n（0.5 天）",
             "1. 撬块边界线、楼梯井净空与实际可操作面\n"
             "2. 紧急淋浴与洗眼器的位置和尺寸\n"
             "3. HTF 穿墙位置与设备间接口\n"
             "4. 货梯与进场路径尺寸",
             "现场实测约束当场记录并双方确认，作为本周全部布置决策的固定基准。",
             CLARE_WALK, 36),
            ("12:00–13:00", "午餐", None, None, None, 14),
            ("13:00–17:00", "2.1 设备顺序调整\n（0.5 天）",
             "1. U 形重排可释放的宽度\n"
             "2. 预热器 / 冷凝器 / 分离器 / 产品罐位置\n"
             "3. 重排后管路长度与交叉",
             "设备顺序定案，释放宽度以数值双方认可，当日更新模型。",
             CLARE_NO, 34),
        ]),
        ("3D 模型评审", "10月13日（二）", [
            ("09:00–12:00", "2.2 维修通道变更\n（0.5 天）",
             "1. 中间通道与可拉出前撬，选定一条维护路线\n"
             "2. 受限空间内约 100 kg 法兰的人工搬运\n"
             "3. 法兰离墙距离、探身深度与垫片更换空间",
             "维护可达性路线定案，并明确对现场拆装工程量与 SAT 范围的影响。",
             CLARE_OPS, 36),
            ("12:00–13:00", "午餐", None, None, None, 14),
            ("13:00–17:00", "2.3 控制柜布置（分体）\n（第 1 段）",
             "1. 控制柜分柜布置方案\n"
             "2. 本安柜与非本安柜分别设置\n"
             "3. 与紧急淋浴 / 洗眼器的位置关系",
             "控制柜位置方案确认，不再悬置。",
             CLARE_NO, 32),
        ]),
        ("3D 模型评审", "10月14日（三）", [
            ("09:00–12:00", "2.3 控制柜布置（分体）\n（第 2 段）",
             "1. 接线箱拆分与电缆数对齐\n"
             "2. 柜体铰接，免拆已测电缆发运\n"
             "3. 按本段确定的仪表归属核算柜体尺寸",
             "柜体数量、尺寸与安装位置定案，电缆数对齐为双方认可的唯一数值。",
             CLARE_NO, 34),
            ("12:00–13:00", "午餐", None, None, None, 14),
            ("13:00–17:00", "2.4 四个分体撬块方案\n（0.5 天）",
             "1. 分体边界与各模块 I/O 归属\n"
             "2. 吊耳、滑移搬运与现场重组顺序\n"
             "3. 整撬高度适配货梯（顶部可拆、HTF 散件发运）",
             "分体边界与各模块 I/O 冻结，进货梯降高方案确认可行。",
             CLARE_NO, 34),
        ]),
        ("操作手册", "10月15日（四）", [
            ("09:00–12:00", "3.1 操作手册讨论\n（第 1 段）",
             "1. 对照本周已定布置逐节走查操作手册\n"
             "2. 催化剂装料与卸料（含水润催化剂卸料顺序）\n"
             "3. 清洗操作的阀门开关与软管连接",
             "意见当场记录并逐条处置，操作手册修订范围达成一致。",
             CLARE_OPS, 34),
            ("12:00–13:00", "午餐", None, None, None, 14),
            ("13:00–17:00", "3.1 操作手册讨论\n（第 2 段）",
             "1. 开车、正常操作、停车与应急处置\n"
             "2. 结合现场勘察确认阀门与取样点的操作可达性\n"
             "3. 剩余意见逐条关闭",
             "操作手册签认，或每条未决意见均有双方认可的处置与责任人。",
             CLARE_OPS_SAFE, 34),
        ]),
        ("控制说明与顺控", "10月16日（五）", [
            ("09:00–12:00", "3.2 控制说明讨论\n（0.5 天）",
             "1. 控制回路、联锁与故障位\n"
             "2. 硬联锁与远程 I/O 的划分\n"
             "3. 仪表信号形式（Modbus 或 4–20 mA、本安或隔爆）",
             "控制说明达成一致，仪表选型条目关闭，I/O 表可据此发布。",
             CLARE_SAFE, 34),
            ("12:00–13:00", "午餐", None, None, None, 14),
            ("13:00–17:00", "3.3 顺控方案讨论\n（0.5 天）",
             "1. 顺控步骤、允许条件与保持点\n"
             "2. 清洗顺控（含惰化）及清洗模式下高液位联锁旁通\n"
             "3. 操作界面与报警处理",
             "顺控方案达成一致并作为 HAZOP 输入，进入分析前不留未定的顺控问题。",
             CLARE_OPS_SAFE, 34),
        ]),
        ("HAZOP 分析", "10月20日（二）", [
            ("09:00–17:00", "4. HAZOP 分析\n（第 1 天）",
             "确认分析范围、节点划分及所用 P&ID / 布置版本。\n"
             "节点：进料、缓冲罐 BT01 / BT04 与计量泵；"
             "预热器与反应器 CR01–CR03（冷却失效、热点、氢气供给）。",
             "分析基准锁定为本周定案布置；当日节点完成，行动项与责任人会上记录。",
             CLARE_SAFE, 32),
        ]),
        ("HAZOP 分析", "10月21日（三）", [
            ("09:00–17:00", "4. HAZOP 分析\n（第 2 天）",
             "节点：冷凝器、气液分离器 SE01、安全泄放与放空；"
             "产品缓冲罐、精滤器与产品出料；结合现场勘察复核阀门可达性与取样。",
             "节点当场关闭；由此引出的布置或阀门变更，对照实体空间当场解决。",
             CLARE_SAFE, 28),
        ]),
        ("HAZOP 分析", "10月22日（四）", [
            ("09:00–12:00", "4. HAZOP 分析\n（第 3 天，上午）",
             "节点：CIP 与清洗回路、惰化、公用工程、控制与联锁。\n"
             "汇总行动项、责任人与时间并双方签认；确认需升版的文件及发布日期。",
             "全部节点完成；离场前形成经双方签认的行动项与文件升版清单。",
             CLARE_SAFE, 28),
        ]),
    ]

    row = 4
    no = 1
    stripe = 0
    for banner, date, items in DAYS:
        # banner
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=last)
        b = ws.cell(row, 1, value=banner)
        b.font = Font(name="微软雅黑", size=9, bold=True, color=NAVY_DK)
        b.fill = banner_fill
        b.alignment = Alignment(horizontal="center", vertical="center")
        b.border = thin_border()
        for col in range(2, last + 1):
            c = ws.cell(row, col)
            c.fill = fill(BANNER)
            c.border = thin_border()
        ws.row_dimensions[row].height = 16
        row += 1

        first = row
        for time, itinerary, detail, objective, note, height in items:
            lunch = detail is None
            bg = fill(LUNCH_BG) if lunch else fill(ROW_B if stripe % 2 else ROW_A)
            if not lunch:
                stripe += 1

            if lunch:
                onsite = online = remark = ""
            elif note == CLARE_NO:
                onsite, online = people(False)
                remark = note
            else:
                onsite, online = people(True)
                remark = note

            vals = [no, date if row == first else None, time, itinerary,
                    onsite, online, detail or "", objective or "", remark]
            for col, val in enumerate(vals, start=1):
                c = ws.cell(row, col, value=val)
                center_cols = {1, 2, 3, 5, 6}
                c.font = font(8.5 if col in (5, 6, 9) else 9,
                              bold=(col == 4 and not lunch),
                              color=MUTED if lunch else INK)
                c.fill = bg
                c.alignment = align(
                    "center" if col in center_cols else "left",
                    "center", True)
                c.border = thin_border()
            ws.row_dimensions[row].height = height
            no += 1
            row += 1

        if row - 1 > first:
            ws.merge_cells(start_row=first, start_column=2,
                           end_row=row - 1, end_column=2)

    # ----- total -------------------------------------------------------
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row, end_column=last)
    tot = ws.cell(
        row, 1,
        value="合计 7.5 个工作日：现场勘察 0.5 ＋ 3D 模型 2.5 ＋ 操作手册与控制说明 2 ＋ HAZOP 2.5",
    )
    tot.font = Font(name="微软雅黑", size=9, bold=True, color=WHITE)
    tot.fill = fill(NAVY)
    tot.alignment = Alignment(horizontal="center", vertical="center")
    tot.border = thin_border()
    for col in range(2, last + 1):
        c = ws.cell(row, col)
        c.fill = fill(NAVY)
        c.border = thin_border()
    ws.row_dimensions[row].height = 18

    ws.freeze_panes = "A4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.35, bottom=0.35,
        header=0.15, footer=0.15,
    )
    ws.print_title_rows = "1:3"
    ws.sheet_view.showGridLines = False
    ws.oddFooter.left.text = "Asymchem  ·  Sandwich 连续氢化撬块"
    ws.oddFooter.right.text = "纯中文版  ·  一页展示"

    wb.save(DST)
    print(f"Saved: {DST}  rows={row}  cols={last}")


if __name__ == "__main__":
    main()
