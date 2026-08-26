"""
Build a presentation-quality mind map (脑图) summarising the continuous
hydrogenation project progress for 2026-07-01 ~ 2026-07-15, for the biweekly
leadership + SW review meeting.

English is the primary language; Chinese is shown as small secondary text.
Output: an HTML file with inline SVG connectors + styled cards, plus a
high-resolution PNG rendered via headless Chrome.
"""

import subprocess
import os

W, H = 1680, 1050

CX, CY = 840, 530
CENTER = dict(left=690, top=460, w=300, h=140)

# Each bullet: (en, cn, small, subs)  subs = list of (en, cn)
CARDS = {
    "overall": dict(
        left=55, top=90, w=560, color="#2563EB",
        title="Overall Progress", sub="整体推进",
        anchor=("right", 250),
        bullets=[
            ("84 merged risks tracked one-by-one to closure",
             "84 项风险清单合并统一跟踪，逐条闭环", False, []),
            ("Key design deliverables — current status:",
             "关键设计文件更新", False, [
                 ("Cleaning plan in progress (Chaoqun; Xiantao ref.)",
                  "清洗方案编制中（Chaoqun 主责，Xiantao 提供参考）"),
                 ("3D Layout ready — pending joint review",
                  "3D 布置已完成，待综合讨论"),
                 ("Control Philosophy draft ready — pending review",
                  "控制策略已有初版，待综合讨论"),
                 ("Catalyst charging plan under CN discussion",
                  "催化剂加料方案 国内讨论中"),
             ]),
        ]),
    "tech": dict(
        left=55, top=650, w=560, color="#16A34A",
        title="Key Technical Discussions", sub="关键技术讨论",
        anchor=("right", 790),
        bullets=[
            ("SE01 separator — direction preliminarily set",
             "SE01 气液分离器重点讨论，初步确定思路方向", False, [
                 ("Remove side branch · enlarge vessel · flat plate to "
                  "carry level transmitter / spray ball / relief",
                  "去旁路 · 加大分离器 · 用 flat plate 承载液位计/喷淋球/泄放"),
             ]),
            ("Design temperature / flange class discussed",
             "设计温度 / 法兰等级讨论", False, []),
            ("Leak-test strategy discussed",
             "检漏策略讨论", False, []),
        ]),
    "decisions": dict(
        left=1085, top=115, w=545, color="#DC2626",
        title="⚠ Decisions Needed from Leadership", sub="需领导层决策",
        anchor=("left", 300), highlight=True,
        bullets=[
            ("Capacity: CR01 DN80→DN50 may fall below URS 30 kg/day "
             "→ keep capacity or revise URS?",
             "产能：CR01 缩径可能低于 URS 30kg/day → 保产能 or 调整 URS？",
             False, []),
            ("Transport: whole-unit lift vs modular skids → scheme "
             "decision (align with roof replacement?)",
             "运输：整体吊装 vs 分撬模块化（是否配合屋顶更换）",
             True, []),
            ("CE: on-site assembly needs CE-mark rep in UK for 3–4 wks "
             "→ certification boundary & cost",
             "认证：现场组装需 CE-mark 人员赴英 3–4 周 → 边界与成本",
             True, []),
        ]),
    "next": dict(
        left=1085, top=640, w=545, color="#0E9AA7",
        title="Next Steps", sub="下一步",
        anchor=("left", 780),
        bullets=[
            ("Size up SE01 & consult Rosemount on the probe",
             "SE01 做大方案测算 + 与 Rosemount 沟通探针", False, []),
            ("Cleaning plan next week → align acceptance criteria",
             "清洗方案下周出 → 对齐验收标准", False, []),
            ("Control philosophy & interlock schedule next week",
             "控制策略 / 联锁清单下周发出", False, []),
        ]),
}


def anchor_point(card):
    side, y = card["anchor"]
    if side == "right":
        return card["left"] + card["w"], y
    return card["left"], y


def _clamp(y):
    lo = CENTER["top"] + 20
    hi = CENTER["top"] + CENTER["h"] - 20
    return max(lo, min(hi, y))


def center_edge(card):
    ax, ay = anchor_point(card)
    if ax < CX:
        return CENTER["left"], _clamp(ay)
    return CENTER["left"] + CENTER["w"], _clamp(ay)


def connector_path(card):
    sx, sy = center_edge(card)
    ex, ey = anchor_point(card)
    dx = (ex - sx) * 0.5
    return f"M {sx},{sy} C {sx+dx},{sy} {ex-dx},{ey} {ex},{ey}"


def build_bullet(color, en, cn, small, subs):
    scls = " small" if small else ""
    subs_html = ""
    if subs:
        rows = "".join(
            f'<li class="sub"><span class="sdash">–</span>'
            f'<div class="txt"><div class="en">{se}</div>'
            f'<div class="cn">{sc}</div></div></li>'
            for se, sc in subs)
        subs_html = f'<ul class="sublist">{rows}</ul>'
    return (f'<li class="blt{scls}">'
            f'<span class="dot" style="background:{color}"></span>'
            f'<div class="txt"><div class="en">{en}</div>'
            f'<div class="cn">{cn}</div>{subs_html}</div></li>')


def build_card_html(card):
    hl = card.get("highlight")
    body = "".join(build_bullet(card["color"], *b) for b in card["bullets"])
    cls = "card highlight" if hl else "card"
    return f'''
    <div class="{cls}" style="left:{card['left']}px;top:{card['top']}px;
         width:{card['w']}px; border-top:6px solid {card['color']};">
      <div class="card-head">
        <span class="ttl" style="color:{card['color']}">{card['title']}</span>
        <span class="sub">{card['sub']}</span>
      </div>
      <ul class="bullets">{body}</ul>
    </div>'''


paths = "\n".join(
    f'<path d="{connector_path(c)}" stroke="{c["color"]}" '
    f'stroke-width="3" fill="none" opacity="0.55"/>'
    for c in CARDS.values())
nubs = "\n".join(
    f'<circle cx="{anchor_point(c)[0]}" cy="{anchor_point(c)[1]}" r="5" '
    f'fill="{c["color"]}"/>'
    for c in CARDS.values())
cards_html = "\n".join(build_card_html(c) for c in CARDS.values())

html = f'''<!doctype html><html lang="en"><head><meta charset="utf-8">
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  html,body {{ width:{W}px; height:{H}px;
     font-family:"Segoe UI","Microsoft YaHei","Noto Sans CJK SC",sans-serif;
     background:linear-gradient(135deg,#f4f7fb 0%,#eaf0f8 100%); }}
  .stage {{ position:relative; width:{W}px; height:{H}px; }}
  .title {{ position:absolute; left:0; top:22px; width:100%;
     text-align:center; }}
  .title h1 {{ font-size:29px; color:#00366A; letter-spacing:.5px; }}
  .title .rng {{ font-size:14px; color:#5b6b83; margin-top:5px;
     font-weight:400; }}
  svg {{ position:absolute; left:0; top:0; }}
  .center {{ position:absolute; left:{CENTER['left']}px;
     top:{CENTER['top']}px; width:{CENTER['w']}px; height:{CENTER['h']}px;
     border-radius:18px;
     background:linear-gradient(135deg,#00366A 0%,#12539e 100%);
     color:#fff; display:flex; flex-direction:column; align-items:center;
     justify-content:center; box-shadow:0 10px 26px rgba(0,54,106,.35);
     text-align:center; padding:10px; }}
  .center .big {{ font-size:22px; font-weight:700; line-height:1.25; }}
  .center .cn {{ font-size:13px; opacity:.9; margin-top:6px; }}
  .center .tag {{ font-size:12px; margin-top:9px;
     background:rgba(255,255,255,.16); padding:3px 12px; border-radius:20px; }}
  .card {{ position:absolute; background:#fff; border-radius:14px;
     box-shadow:0 6px 20px rgba(30,50,80,.12); padding:13px 18px 15px; }}
  .card.highlight {{ box-shadow:0 8px 26px rgba(220,38,38,.28);
     background:#fff7f7; }}
  .card-head {{ display:flex; align-items:baseline; gap:10px;
     border-bottom:1px solid #eef1f6; padding-bottom:8px; margin-bottom:8px; }}
  .card-head .ttl {{ font-size:19px; font-weight:700; }}
  .card-head .sub {{ font-size:13px; color:#8a97ab; }}
  ul.bullets {{ list-style:none; }}
  li.blt {{ display:flex; align-items:flex-start; gap:9px; margin:8px 0; }}
  .dot {{ width:8px; height:8px; border-radius:50%; margin-top:6px;
     flex:0 0 auto; }}
  .txt .en {{ font-size:15px; color:#1e2b3a; line-height:1.34;
     font-weight:600; }}
  .txt .cn {{ font-size:11.5px; color:#8a97ab; line-height:1.3;
     margin-top:1px; font-weight:400; }}
  li.blt.small .txt .en {{ font-size:12px; font-weight:500; color:#5a6675; }}
  li.blt.small .txt .cn {{ font-size:10px; }}
  li.blt.small .dot {{ width:6px; height:6px; margin-top:5px; }}
  ul.sublist {{ list-style:none; margin:4px 0 2px 2px; }}
  li.sub {{ display:flex; align-items:flex-start; gap:7px; margin:3px 0; }}
  li.sub .sdash {{ color:#b7c0cf; font-size:12px; margin-top:1px; }}
  li.sub .txt .en {{ font-size:12.5px; font-weight:500; color:#3c4a5a; }}
  li.sub .txt .cn {{ font-size:10px; }}
  .card.highlight li.blt:first-child .txt .en {{ font-size:16px;
     font-weight:700; color:#991b1b; }}
  .card.highlight li.blt:first-child .txt .cn {{ font-size:12px;
     color:#b45c5c; }}
  .foot {{ position:absolute; left:0; bottom:14px; width:100%;
     text-align:center; font-size:11.5px; color:#9aa6b8; }}
</style></head>
<body>
  <div class="stage">
    <div class="title">
      <h1>Continuous Hydrogenation Skid · Bi-weekly Progress</h1>
      <div class="rng">连续氢化建设项目 · 双周进展汇报 &nbsp;|&nbsp;
        2026-07-01 ~ 07-15</div>
    </div>
    <svg width="{W}" height="{H}">{paths}{nubs}</svg>
    <div class="center">
      <div class="big">Progress Snapshot<br>0701 – 0715</div>
      <div class="cn">连续氢化项目进展</div>
      <div class="tag">On track · 整体节奏可控</div>
    </div>
    {cards_html}
    <div class="foot">Asymchem · Sandwich Site &nbsp;|&nbsp;
      IEPE · CFCT · UK Sandwich</div>
  </div>
</body></html>'''

out_html = "progress_mindmap_0715.html"
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html)
print("Saved:", out_html)

out_png = "progress_mindmap_0715.png"
abspath = os.path.abspath(out_html)
cmd = [
    "google-chrome", "--headless=new", "--no-sandbox", "--disable-gpu",
    "--disable-dev-shm-usage", "--hide-scrollbars",
    "--force-device-scale-factor=2", "--virtual-time-budget=8000",
    "--run-all-compositor-stages-before-draw",
    f"--window-size={W},{H}",
    f"--screenshot={os.path.abspath(out_png)}",
    f"file://{abspath}",
]
res = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
print("chrome rc:", res.returncode)
print("Saved:", out_png)
