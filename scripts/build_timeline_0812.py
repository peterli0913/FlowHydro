"""
Generate a presentation-grade HTML Gantt view of the 12 Aug 2026 revision of
the Sandwich continuous hydrogenation master schedule.

Week columns run from 2026-08 W1 to 2027-12 W4 (17 months x 4 weeks).
Bar positions are taken from the "Timeline 0812" sheet of
`0812-Sandwich Continuous Hydrogen Timeline.xlsx`.
"""

from pathlib import Path

START_YEAR, START_MONTH = 2026, 8
N_MONTHS = 17
N_COLS = N_MONTHS * 4

TODAY = (2026, 8, 2)          # 12 Aug 2026 -> week 2 of Aug


def col(year, month, week):
    idx = (year - START_YEAR) * 12 + (month - START_MONTH)
    return idx * 4 + week


def months():
    out = []
    y, m = START_YEAR, START_MONTH
    for _ in range(N_MONTHS):
        out.append((y, m))
        m += 1
        if m == 13:
            m = 1
            y += 1
    return out


# kind: done | active | plan | ship | site | milestone
PHASES = [
    ("Design Finalisation", "设计定稿", "#2563EB", [
        ("Design documents re-issued (P&ID, datasheets, interlock logic)",
         "设计文件升版（P&ID、设备/仪表数据单、联锁表）",
         (2026, 8, 1), (2026, 8, 2), "done"),
        ("Cleaning philosophy · control philosophy · operating manual",
         "清洗说明 · 控制说明 · 操作手册",
         (2026, 8, 1), (2026, 8, 2), "active"),
        ("FDS / SDS / HDS",
         "功能/软件/硬件设计说明",
         (2026, 8, 1), (2026, 8, 4), "active"),
        ("PFD &amp; mass balance — final on R&amp;D reactor data",
         "PFD 与物料平衡 —— 待研发反应器数据定版",
         (2026, 8, 1), (2026, 9, 2), "plan"),
        ("Reactor column specification — DN65 / DN50 decision",
         "反应柱规格书 —— DN65 / DN50 定案",
         (2026, 8, 1), (2026, 9, 2), "plan"),
        ("HAZOP",
         "HAZOP 分析",
         (2026, 9, 1), (2026, 9, 2), "milestone"),
        ("Process &amp; instrument documents revised after HAZOP",
         "HAZOP 后工艺与仪表文件升版",
         (2026, 9, 3), (2026, 9, 4), "plan"),
        ("Cable schedule",
         "电缆清单",
         (2026, 8, 1), (2026, 10, 2), "plan"),
        ("PLC hardware drawing · I/O allocation",
         "PLC 硬件图 · I/O 分配表",
         (2026, 10, 1), (2026, 11, 1), "plan"),
    ]),
    ("Procurement &amp; Fabrication", "采购与制造", "#7C3AED", [
        ("Engineering procurement — valves, instruments, static equipment",
         "工程采购 —— 阀门、仪表、静设备",
         (2026, 9, 1), (2027, 2, 1), "plan"),
        ("Procurement information confirmation",
         "采购信息确认",
         (2026, 10, 3), (2026, 11, 2), "plan"),
        ("3D model frozen",
         "3D 模型定版",
         (2026, 10, 2), (2026, 12, 1), "plan"),
        ("BOM",
         "材料表",
         (2026, 12, 2), (2026, 12, 3), "plan"),
        ("Fabrication drawings",
         "制造图",
         (2026, 12, 4), (2027, 1, 1), "plan"),
        ("Control system procurement &amp; configuration",
         "控制系统采购与组态",
         (2026, 9, 2), (2027, 1, 2), "plan"),
        ("Equipment fabrication, assembly &amp; debugging",
         "设备制造、组装与调试",
         (2027, 1, 3), (2027, 3, 1), "plan"),
    ]),
    ("FAT, Certification &amp; Shipping", "工厂验收、认证与运输", "#B45309", [
        ("Integrated factory acceptance test (FAT)",
         "整橇工厂验收测试（FAT）",
         (2027, 3, 2), (2027, 4, 1), "plan"),
        ("CE certification — document review phase",
         "CE 认证 —— 资料评审阶段",
         (2026, 9, 1), (2026, 11, 4), "plan"),
        ("CE certification — equipment &amp; whole-skid assessment",
         "CE 认证 —— 单体设备与整橇评定",
         (2027, 2, 1), (2027, 5, 1), "plan"),
        ("Export formalities · transport inspection",
         "出口手续 · 运输检验",
         (2027, 5, 2), (2027, 5, 3), "ship"),
        ("Sea freight to UK",
         "海运至英国",
         (2027, 5, 4), (2027, 8, 3), "ship"),
    ]),
    ("Site Installation &amp; Qualification", "现场安装与确认", "#0E9AA7", [
        ("Skid installation",
         "撬块就位安装",
         (2027, 8, 4), (2027, 9, 1), "site"),
        ("Mechanical, electrical, instrumentation &amp; piping",
         "机、电、仪、管安装",
         (2027, 9, 2), (2027, 10, 3), "site"),
        ("Qualification &amp; acceptance testing (DQ / IQ / OQ)",
         "确认与验收测试（DQ / IQ / OQ）",
         (2027, 10, 4), (2027, 11, 2), "site"),
        ("Commissioning and start-up — beneficial use",
         "试车与开车 —— 可投入使用",
         (2027, 11, 3), (2027, 11, 3), "milestone"),
    ]),
]

COMPARE = [
    ("HAZOP", "HAZOP 分析", "Sep 2026 W1", "Sep 2026 W1", "no change"),
    ("Reactor size decision (DN65 / DN50)", "反应器尺寸定案",
     "Jun 2026 W4", "Sep 2026 W2", "+2.5 months"),
    ("Procurement information confirmation", "采购信息确认",
     "Aug 2026 W1", "Nov 2026 W2", "+3 months"),
    ("Equipment fabrication complete", "设备制造完成",
     "Jan 2027 W1", "Mar 2027 W1", "+2 months"),
    ("Integrated FAT complete", "整橇 FAT 完成",
     "Feb 2027 W1", "Apr 2027 W1", "+2 months"),
    ("CE certification complete", "CE 认证完成",
     "Mar 2027 W1", "May 2027 W1", "+2 months"),
    ("Arrival in UK", "抵达英国",
     "Jun 2027 W3", "Aug 2027 W3", "+2 months"),
    ("Commissioning / beneficial use", "投用（Beneficial use）",
     "Sep 2027 W3", "Nov 2027 W3", "+2 months"),
]

DRIVERS = [
    ("Reactor sizing depends on R&amp;D data due early September",
     "反应器尺寸取决于 9 月初才产出的研发数据",
     "The high-temperature point found on DN80 requires a change to DN65 or "
     "DN50. Until the size is fixed the main equipment cannot be released "
     "for purchase, and re-issuing after enquiry would invalidate the "
     "quotations already obtained.",
     "DN80 出现高温点，需改为 DN65 或 DN50。尺寸未定则主设备无法提交采购；"
     "若先行询价后再变更，已取得的比价将作废并需重走流程。"),
    ("Procurement sequence corrected to reflect real dependencies",
     "采购顺序按真实依赖关系重排",
     "The previous schedule allowed procurement, BOM and FAT to overlap the "
     "design freeze. The revision restores the correct chain: design freeze "
     "→ HAZOP → document revision → 3D freeze → procurement confirmation → "
     "vendor returns → BOM and fabrication drawings → build → FAT → CE.",
     "原计划中采购、材料表与 FAT 与设计定版存在重叠。本次恢复真实链条："
     "设计定版 → HAZOP → 文件升版 → 3D 定版 → 采购信息确认 → 厂家返资料 → "
     "材料表与制造图 → 制造 → FAT → CE 认证。"),
    ("Vendor feedback loop was previously under-counted",
     "厂家返资料的往复周期此前计入不足",
     "Enquiry, comparison, vendor document return and our review is a real "
     "cycle of several weeks. Cable schedule, PLC hardware drawings and I/O "
     "allocation cannot be frozen until instrument vendor selection returns.",
     "询价、比价、厂家返资料与我方审查是数周量级的实际往复。电缆清单、"
     "PLC 硬件图与 I/O 分配表须待仪表厂家最终选型返回后方可定版。"),
    ("CE notified body not yet appointed",
     "CE 认证机构尚未选定",
     "Document preparation starts once the 3D model is frozen; the body can "
     "then be appointed and the review phase begun.",
     "资料准备待 3D 模型定版后启动，随后确定机构并进入资料评审阶段。"),
]


def bar_html(row, ncols):
    label_en, label_cn, s, e, kind = row
    c1 = col(*s)
    c2 = col(*e)
    span = max(1, c2 - c1 + 1)
    if kind == "milestone":
        inner = f'<span class="dia"></span>'
        cls = "bar ms"
    else:
        inner = ""
        cls = f"bar {kind}"
    return (f'<div class="{cls}" style="grid-column:{c1} / span {span}">'
            f'{inner}</div>')


def build():
    ms = months()

    # month header
    head = []
    for i, (y, m) in enumerate(ms):
        first = (m == 1) or (i == 0)
        yr = f'<span class="yr">{y}</span>' if first else ""
        head.append(
            f'<div class="mo" style="grid-column:{i*4+1} / span 4">'
            f'{yr}{m:02d}</div>')
    head_html = "\n        ".join(head)

    today_col = col(*TODAY)

    rows_html = []
    for pname_en, pname_cn, pcolor, rows in PHASES:
        rows_html.append(
            f'''      <div class="phase" style="--pc:{pcolor}">
        <div class="phase-label">{pname_en}<span>{pname_cn}</span></div>
        <div class="phase-grid"><div class="today" '''
            f'''style="grid-column:{today_col}"></div></div>
      </div>''')
        for r in rows:
            label_en, label_cn, s, e, kind = r
            date_txt = ""
            if kind == "milestone":
                y, m, w = s
                date_txt = f'{MONTH_ABBR[m]} {y} W{w}'
            else:
                y1, m1, w1 = s
                y2, m2, w2 = e
                date_txt = (f'{MONTH_ABBR[m1]} {y1} W{w1} – '
                            f'{MONTH_ABBR[m2]} {y2} W{w2}')
            star = ' ★' if 'beneficial use' in label_en else ''
            rows_html.append(f'''      <div class="row" style="--pc:{pcolor}">
        <div class="label"><b>{label_en}{star}</b><span>{label_cn}</span></div>
        <div class="track">
          <div class="today" style="grid-column:{today_col}"></div>
          {bar_html(r, N_COLS)}
        </div>
        <div class="when">{date_txt}</div>
      </div>''')
    rows_html = "\n".join(rows_html)

    cmp_rows = []
    for en, cn, old, new, delta in COMPARE:
        cls = "same" if delta == "no change" else "slip"
        cmp_rows.append(f'''        <tr>
          <td class="ci"><b>{en}</b><span>{cn}</span></td>
          <td class="old">{old}</td>
          <td class="new">{new}</td>
          <td class="{cls}">{delta}</td>
        </tr>''')
    cmp_html = "\n".join(cmp_rows)

    drv = []
    for en, cn, den, dcn in DRIVERS:
        drv.append(f'''        <div class="drv">
          <b>{en}</b><span class="cn">{cn}</span>
          <p>{den}</p><p class="cn">{dcn}</p>
        </div>''')
    drv_html = "\n".join(drv)

    return TEMPLATE.format(head=head_html, rows=rows_html, cmp=cmp_html,
                           drv=drv_html, ncols=N_COLS)


MONTH_ABBR = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
              7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}


TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Sandwich Continuous Hydrogenation Skid — Master Schedule (Rev. 12 Aug 2026)</title>
<style>
  :root {{
    --ink:#16202e; --muted:#64748b; --line:#e2e8f0; --paper:#f4f7fb;
    --navy:#0f3460;
  }}
  * {{ box-sizing:border-box; }}
  body {{
    margin:0; padding:26px 22px 44px; color:var(--ink);
    background:linear-gradient(135deg,#f5f8fc 0%,#eaf0f8 100%);
    font-family:'Segoe UI','Helvetica Neue','PingFang SC','Microsoft YaHei',Arial,sans-serif;
    line-height:1.45;
  }}
  .wrap {{ max-width:1560px; margin:0 auto; }}

  header {{
    background:linear-gradient(120deg,#0f3460,#16213e);
    color:#fff; border-radius:14px; padding:22px 28px;
    box-shadow:0 8px 24px rgba(15,52,96,.22); margin-bottom:16px;
    display:flex; justify-content:space-between; align-items:center;
    gap:24px; flex-wrap:wrap;
  }}
  header h1 {{ margin:0; font-size:22px; letter-spacing:.2px; }}
  header .sub {{ display:block; font-size:13px; color:#c3d2e4; margin-top:6px; font-weight:400; }}
  .keybox {{
    background:rgba(255,255,255,.09); border:1px solid rgba(255,255,255,.22);
    border-radius:12px; padding:12px 20px; text-align:center; min-width:250px;
  }}
  .keybox .cap {{ font-size:11px; letter-spacing:.12em; text-transform:uppercase; color:#f0d3a2; }}
  .keybox .val {{ font-size:26px; font-weight:700; margin-top:4px; }}
  .keybox .note {{ font-size:11.5px; color:#c3d2e4; margin-top:3px; }}

  .legend {{
    display:flex; gap:16px; flex-wrap:wrap; align-items:center;
    background:#fff; border:1px solid var(--line); border-radius:10px;
    padding:9px 16px; margin-bottom:12px; font-size:12px; color:var(--muted);
  }}
  .legend i {{ display:inline-block; width:22px; height:10px; border-radius:3px;
               vertical-align:middle; margin-right:6px; }}
  .legend .dia-l {{ display:inline-block; width:11px; height:11px; background:#DC2626;
                    transform:rotate(45deg); margin-right:8px; vertical-align:middle; }}

  .chart {{
    background:#fff; border:1px solid var(--line); border-radius:14px;
    box-shadow:0 4px 16px rgba(30,50,80,.07); padding:14px 18px 18px;
    margin-bottom:18px; overflow-x:auto;
  }}
  .grid-head {{
    display:grid; grid-template-columns:340px 1fr 150px; align-items:end;
    gap:0; border-bottom:2px solid #cfdae7; padding-bottom:6px; margin-bottom:6px;
    min-width:1180px;
  }}
  .grid-head .sp {{ font-size:11px; color:var(--muted); letter-spacing:.08em;
                    text-transform:uppercase; }}
  .months {{ display:grid; grid-template-columns:repeat({ncols},1fr); }}
  .mo {{
    text-align:center; font-size:11px; color:#48607d; font-weight:600;
    border-left:1px solid #eef2f7; padding:2px 0; position:relative;
  }}
  .mo .yr {{ display:block; font-size:10px; color:#94a3b8; font-weight:700; }}

  .phase {{ display:grid; grid-template-columns:340px 1fr 150px;
            margin:12px 0 4px; align-items:center; min-width:1180px; }}
  .phase-label {{
    font-size:13.5px; font-weight:700; color:var(--pc);
    border-left:5px solid var(--pc); padding-left:10px;
  }}
  .phase-label span {{ display:block; font-size:10.5px; color:#94a3b8; font-weight:500; }}
  .phase-grid {{ display:grid; grid-template-columns:repeat({ncols},1fr);
                 height:14px; position:relative; }}

  .row {{ display:grid; grid-template-columns:340px 1fr 150px;
          align-items:center; min-height:30px; min-width:1180px; }}
  .row:hover {{ background:#f8fbff; }}
  .label {{ padding:3px 12px 3px 22px; }}
  .label b {{ font-size:12.5px; font-weight:600; color:#22303f; display:block; }}
  .label span {{ font-size:10.5px; color:#95a3b4; }}
  .track {{
    display:grid; grid-template-columns:repeat({ncols},1fr);
    height:26px; align-items:center; position:relative;
    background-image:repeating-linear-gradient(to right,#f1f5f9 0 1px,transparent 1px calc(100%/{ncols}));
  }}
  .when {{ font-size:10.5px; color:#8496a9; text-align:right; padding-right:4px;
           white-space:nowrap; }}

  .bar {{ height:13px; border-radius:7px; position:relative; }}
  .bar.done   {{ background:linear-gradient(90deg,#16A34A,#22c55e); }}
  .bar.active {{ background:linear-gradient(90deg,#eab308,#facc15); }}
  .bar.plan   {{ background:linear-gradient(90deg,var(--pc),var(--pc)); opacity:.78; }}
  .bar.ship   {{ background:linear-gradient(90deg,#B45309,#f59e0b); }}
  .bar.site   {{ background:linear-gradient(90deg,#0E9AA7,#22c3d0); }}
  .bar.ms {{ background:transparent; display:flex; align-items:center;
             justify-content:center; }}
  .bar.ms .dia {{ width:14px; height:14px; background:#DC2626; transform:rotate(45deg);
                  box-shadow:0 0 0 3px rgba(220,38,38,.16); }}

  .today {{ grid-row:1; align-self:stretch; border-left:2px dashed #dc2626;
            opacity:.45; }}

  .panels {{ display:grid; grid-template-columns:1.05fr 1fr; gap:16px; }}
  .panel {{
    background:#fff; border:1px solid var(--line); border-radius:14px;
    box-shadow:0 4px 16px rgba(30,50,80,.07); padding:16px 20px 18px;
  }}
  .panel h2 {{ margin:0 0 4px; font-size:16px; color:var(--navy); }}
  .panel h2 span {{ display:block; font-size:11.5px; color:#94a3b8; font-weight:400; margin-top:2px; }}

  table {{ width:100%; border-collapse:collapse; font-size:12.5px; margin-top:10px; }}
  th, td {{ border-bottom:1px solid #eef2f7; padding:7px 8px; text-align:left; vertical-align:middle; }}
  th {{ font-size:10.5px; color:#8496a9; text-transform:uppercase; letter-spacing:.06em;
        border-bottom:2px solid #dbe4ee; }}
  td.ci b {{ display:block; font-size:12.5px; color:#22303f; }}
  td.ci span {{ font-size:10.5px; color:#95a3b4; }}
  td.old {{ color:#9aa6b8; width:110px; text-decoration:line-through; }}
  td.new {{ color:#0f3460; font-weight:700; width:110px; }}
  td.slip {{ color:#b91c1c; font-weight:700; width:96px; text-align:right; }}
  td.same {{ color:#16a34a; font-weight:600; width:96px; text-align:right; }}

  .drv {{ border-left:3px solid #d97706; background:#fffbf3; border-radius:0 8px 8px 0;
          padding:9px 13px; margin-top:10px; }}
  .drv b {{ font-size:12.5px; color:#92400e; }}
  .drv .cn {{ display:block; font-size:10.5px; color:#a97b45; }}
  .drv p {{ margin:5px 0 0; font-size:11.5px; color:#4a5768; line-height:1.55; }}
  .drv p.cn {{ color:#8496a9; font-size:10.5px; margin-top:3px; }}

  footer {{ text-align:center; font-size:11px; color:#9aa6b8; margin-top:20px; }}

  @media (max-width:1100px) {{ .panels {{ grid-template-columns:1fr; }} }}
  @media print {{
    body {{ background:#fff; padding:0; }}
    .chart, .panel {{ box-shadow:none; }}
  }}
</style>
</head>
<body>
<div class="wrap">

<header>
  <div>
    <h1>Sandwich Continuous Hydrogenation Skid — Master Schedule
      <span class="sub">Rev. 12 August 2026 ｜ 连续氢化撬块总进度计划（2026 年 8 月 12 日版）</span>
    </h1>
  </div>
  <div class="keybox">
    <div class="cap">Beneficial use 可投用</div>
    <div class="val">Nov 2027</div>
    <div class="note">Commissioning &amp; start-up complete · W3</div>
  </div>
</header>

<div class="legend">
  <span><i style="background:linear-gradient(90deg,#16A34A,#22c55e)"></i>Completed 已完成</span>
  <span><i style="background:linear-gradient(90deg,#eab308,#facc15)"></i>Issued / under review 已出版会审</span>
  <span><i style="background:#2563EB;opacity:.78"></i>Planned 计划</span>
  <span><i style="background:linear-gradient(90deg,#B45309,#f59e0b)"></i>Export &amp; shipping 出口运输</span>
  <span><i style="background:linear-gradient(90deg,#0E9AA7,#22c3d0)"></i>Site works 现场施工</span>
  <span><span class="dia-l"></span>Milestone 里程碑</span>
  <span style="margin-left:auto;color:#dc2626">┆ today 今日</span>
</div>

<div class="chart">
  <div class="grid-head">
    <div class="sp">Work item ｜ 工作事项</div>
    <div class="months">
        {head}
    </div>
    <div class="sp" style="text-align:right">Window ｜ 区间</div>
  </div>

{rows}
</div>

<div class="panels">
  <div class="panel">
    <h2>Change against the 29 July revision
      <span>与 7 月 29 日版本的差异</span>
    </h2>
    <table>
      <tr><th>Milestone ｜ 里程碑</th><th>29 Jul</th><th>12 Aug</th><th>Δ</th></tr>
{cmp}
    </table>
  </div>

  <div class="panel">
    <h2>What drove the change
      <span>本次调整的原因</span>
    </h2>
{drv}
  </div>
</div>

<footer>Asymchem · Sandwich Site ｜ IEPE · CFCT · CIMT ｜ Master schedule rev. 12 Aug 2026</footer>

</div>
</body>
</html>
"""


def main():
    html = build()
    out = Path("timeline_0812_master_schedule.html")
    out.write_text(html, encoding="utf-8")
    print("Saved:", out, len(html), "chars")


if __name__ == "__main__":
    main()
