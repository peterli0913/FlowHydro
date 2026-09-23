"""
English counterpart of the latest hand-edited Chinese agenda
`0909-赴Sandwich现场议程（3D与HAZOP）.xlsx`.

The Chinese file is copied, not modified.  Layout, merges, fills, row heights
and the empty Remarks column are kept.  Only visible text is translated.
"""

import shutil
from pathlib import Path

import openpyxl
from openpyxl.worksheet.page import PageMargins

SRC = "0909-赴Sandwich现场议程（3D与HAZOP）.xlsx"
DST = "0909-Agenda of the visit to Sandwich (3D & HAZOP)_EN.xlsx"
SHEET = "Visit Agenda"

# Exact source strings -> English.  Names stay as used with SW (pinyin).
TEXT = {
    "赴 Sandwich 现场议程  ·  3D 布置与 HAZOP":
        "Agenda of the visit to Sandwich  ·  3D layout & HAZOP",

    "现场：Keith、Clare、林浩淦、范双双、胡超群（Clare 参加现场勘察及操作性、安全相关场次）"
    "     线上：高少峰、孟德智、高宇、赵子亮"
    "     建议时间安排：10月12–16日（布置 / 手册 / 控制），10月20–22日（HAZOP）":
        "On site: Keith, Clare, Lin Haogan, Fan Shuangshuang, Hu Chaoqun"
        " (Clare joins the site walkdown and the operability / safety sessions)\n"
        "Online: Gao Shaofeng, Meng Dezhi, Gao Yu, Zhao Ziliang\n"
        "Proposed dates: 12–16 Oct (layout / manuals / control), 20–22 Oct (HAZOP)",

    "序号": "No.",
    "日期": "Date",
    "时间": "Time",
    "行程": "Itinerary",
    "现场参与人员": "On-site attendees",
    "线上参与人员": "Online attendees",
    "详细议题": "Topics",
    "目标": "Objective",
    "备注": "Remarks",

    "现场勘察与 3D 模型评审": "Site walkdown & 3D model review",
    "3D 模型评审": "3D model review",
    "操作手册": "Operating manual",
    "控制说明与顺控": "Control philosophy & sequence",
    "HAZOP 分析": "HAZOP study",

    "10月12日（一）": "12 Oct (Mon)",
    "10月13日（二）": "13 Oct (Tue)",
    "10月14日（三）": "14 Oct (Wed)",
    "10月15日（四）": "15 Oct (Thu)",
    "10月16日（五）": "16 Oct (Fri)",
    "10月20日（二）": "20 Oct (Tue)",
    "10月21日（三）": "21 Oct (Wed)",
    "10月22日（四）": "22 Oct (Thu)",

    "1. 现场勘察\n（0.5 天）": "1. Site walkdown\n(0.5 day)",
    "2.1 设备顺序调整\n（0.5 天）": "2.1 Equipment sequence\n(0.5 day)",
    "2.2 维修通道变更\n（0.5 天）": "2.2 Maintenance access\n(0.5 day)",
    "2.3 控制柜布置（分体）\n（第 1 段）": "2.3 Control-cabinet layout (split)\n(Part 1)",
    "2.3 控制柜布置（分体）\n（第 2 段）": "2.3 Control-cabinet layout (split)\n(Part 2)",
    "2.4 三分体撬块方案\n（0.5 天）": "2.4 Three-module split\n(0.5 day)",
    "3.1 操作手册讨论\n（第 1 段）": "3.1 Operating manual\n(Part 1)",
    "3.1 操作手册讨论\n（第 2 段）": "3.1 Operating manual\n(Part 2)",
    "3.2 控制说明讨论\n（0.5 天）": "3.2 Control philosophy\n(0.5 day)",
    "3.3 顺控方案讨论\n（0.5 天）": "3.3 Sequence control\n(0.5 day)",
    "4. HAZOP 分析\n（第 1 天）": "4. HAZOP study\n(Day 1)",
    "4. HAZOP 分析\n（第 2 天）": "4. HAZOP study\n(Day 2)",
    "4. HAZOP 分析\n（第 3 天，上午）": "4. HAZOP study\n(Day 3, morning)",

    "Keith、Clare\n林浩淦、范双双、胡超群":
        "Keith, Clare\nLin Haogan, Fan Shuangshuang, Hu Chaoqun",
    "Keith、林浩淦\n范双双、胡超群":
        "Keith, Lin Haogan\nFan Shuangshuang, Hu Chaoqun",
    "SW team\n林浩淦、范双双、胡超群":
        "SW team\nLin Haogan, Fan Shuangshuang, Hu Chaoqun",
    "高少峰、孟德智\n高宇、赵子亮":
        "Gao Shaofeng, Meng Dezhi\nGao Yu, Zhao Ziliang",

    "1. 撬块边界线、楼梯井净空与实际可操作面\n"
    "2. HTF 穿墙位置与设备间接口\n"
    "3. 货梯与进场路径尺寸":
        "1. Skid boundary, stairwell clearance and usable operating faces\n"
        "2. HTF wall penetration and equipment-room interface\n"
        "3. Goods-lift and delivery-route dimensions",

    "现场实测，作为本周全部布置决策的固定基准。":
        "Site measurements taken as the fixed basis for every layout "
        "decision this week.",

    "1. U 形重排可释放的宽度\n"
    "2. 预热器 / 冷凝器 / 分离器 / 产品罐位置\n"
    "3. 重排后管路长度与交叉":
        "1. Width released by the U-shape re-sequence\n"
        "2. Positions of preheater / condenser / separator / product tank\n"
        "3. Pipe length and crossings after re-sequencing",

    "设备顺序定案，当日更新模型。":
        "Equipment sequence frozen; model updated the same day.",

    "1. 设计中间通道与前撬，选定一条维护路线\n"
    "2. 实测法兰离墙距离、探身深度与垫片更换空间":
        "1. Design the centre aisle and front skid; pick one maintenance route\n"
        "2. Flange-to-wall distance, reach-in depth and gasket-change space",

    "维护可达性路线定案，并明确对现场拆装工程量与 SAT 范围的影响。":
        "Maintenance-access route decided; impact on site strip-out "
        "and SAT scope stated.",

    "1. 控制柜分柜布置方案\n"
    "2. 本安柜与非本安柜分别设置\n"
    "3. 与紧急淋浴 / 洗眼器的位置关系":
        "1. Split-cabinet layout\n"
        "2. IS and non-IS panels as separate enclosures\n"
        "3. Position relative to the emergency shower / eyewash",

    "控制柜位置方案确认。":
        "Control-cabinet position confirmed.",

    "1. 接线箱拆分与电缆数对齐\n"
    "2. 柜体铰接方案设计，免拆已测电缆发运\n"
    "3. 按本段确定的仪表归属核算柜体尺寸":
        "1. Junction-box split; align cable count\n"
        "2. Hinged cabinet — ship without disconnecting tested cables\n"
        "3. Cabinet size from the instrument allocation agreed here",

    "柜体数量、尺寸与安装位置定案，电缆数对齐。":
        "Cabinet quantity, size and mounting position frozen; cable count "
        "aligned.",

    "1. 分体边界与各模块 I/O 归属\n"
    "2. 吊耳、滑移搬运与现场重组顺序\n"
    "3. 整撬高度适配货梯":
        "1. Split boundaries and I/O allocation per module\n"
        "2. Lifting lugs, skate handling and on-site re-assembly sequence\n"
        "3. Overall height to fit the goods lift",

    "分体边界与各模块 I/O 冻结，进货梯降高方案确认可行。":
        "Split boundaries and per-module I/O frozen; goods-lift height "
        "solution confirmed feasible.",

    "1. 对照本周已定布置逐节走查操作手册\n"
    "2. 催化剂装料与卸料（含水润催化剂卸料顺序）\n"
    "3. 清洗操作的阀门开关与软管连接":
        "1. Walk the manual against this week's frozen layout\n"
        "2. Catalyst charge / discharge (incl. water-wet sequence)\n"
        "3. Cleaning: valve line-up and hose connections",

    "意见当场记录并逐条处置，操作手册修订范围达成一致。":
        "Comments recorded and dispositioned in the session; revision "
        "scope of the manual agreed.",

    "1. 开车、正常操作、停车与应急处置\n"
    "2. 结合现场勘察确认阀门与取样点的操作可达性\n"
    "3. 剩余意见逐条关闭":
        "1. Start-up, normal operation, shutdown and emergency response\n"
        "2. Valve and sample-point access, checked against the walkdown\n"
        "3. Remaining comments closed one by one",

    "操作手册签认，或每条未决意见均有双方认可的处置与责任人。":
        "Operating manual signed off, or every open comment has an agreed "
        "disposition and owner.",

    "1. 控制回路、联锁与故障位\n"
    "2. 硬联锁与远程 I/O 的划分\n"
    "3. 仪表信号形式确认":
        "1. Control loops, interlocks and fail positions\n"
        "2. Split between hardwired interlocks and remote I/O\n"
        "3. Confirm instrument signal types",

    "控制说明达成一致，仪表选型条目关闭，I/O 表可据此发布。":
        "Control philosophy agreed; instrument items closed so the I/O "
        "schedule can be issued.",

    "1. 顺控步骤、允许条件与保持点\n"
    "2. 清洗顺控及清洗模式下高液位联锁旁通\n"
    "3. 操作界面与报警处理":
        "1. Sequence steps, permissives and hold points\n"
        "2. Cleaning sequence and high-level interlock bypass in cleaning "
        "mode\n"
        "3. Operator interface and alarm handling",

    "顺控方案达成一致并作为 HAZOP 输入，进入分析前不留未定的顺控问题。":
        "Sequence agreed as HAZOP input; no open sequence question "
        "left going into the study.",

    "当日节点完成。":
        "Nodes for the day completed.",

    "全部节点完成；形成双方确认的行动项与文件升版清单。":
        "All nodes completed; agreed action list and document-revision "
        "list issued.",

    "合计 7.5 个工作日：现场勘察 0.5 ＋ 3D 模型 2.5 ＋ 操作手册与控制说明 2 ＋ HAZOP 2.5":
        "Total 7.5 working days: site walkdown 0.5 + 3D model 2.5 + "
        "operating manual & control 2 + HAZOP 2.5",
}


def main():
    if not Path(SRC).exists():
        raise FileNotFoundError(SRC)

    shutil.copyfile(SRC, DST)
    wb = openpyxl.load_workbook(DST)
    ws = wb.active
    ws.title = SHEET

    missing = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            if isinstance(cell.value, (int, float)):
                continue
            src = cell.value
            if src in TEXT:
                cell.value = TEXT[src]
            elif any("\u4e00" <= ch <= "\u9fff" for ch in str(src)):
                missing.append((cell.coordinate, src))

    if missing:
        raise SystemExit(
            "Untranslated cells:\n" +
            "\n".join(f"  {c}: {v!r}" for c, v in missing)
        )

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.4, right=0.4, top=0.35, bottom=0.35,
        header=0.15, footer=0.15,
    )
    ws.oddFooter.left.text = "Asymchem  ·  Sandwich continuous hydrogenation skid"
    ws.oddFooter.right.text = "English  ·  one page"

    # English runs longer than Chinese; keep one-page A3 fit but give wrap room.
    ws.row_dimensions[2].height = 36
    for r, h in (
        (5, 46), (6, 46), (8, 44), (9, 46), (11, 48), (12, 46),
        (14, 48), (15, 46), (17, 46), (18, 46), (20, 42), (21, 42), (22, 46),
    ):
        ws.row_dimensions[r].height = h
    ws.column_dimensions["G"].width = 38
    ws.column_dimensions["H"].width = 30

    wb.save(DST)
    print(f"Saved: {DST}")


if __name__ == "__main__":
    main()
