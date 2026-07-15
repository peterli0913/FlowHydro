"""
Build a presentation-quality mind map (脑图) summarising the continuous
hydrogenation project progress for 2026-07-01 ~ 2026-07-15, for the biweekly
leadership + SW review meeting.

Focus: overall progress + decisions needed from leadership (light on tech
detail).  Output: an HTML file with inline SVG connectors and styled cards,
plus a high-resolution PNG rendered via headless Chrome.
"""

import subprocess
import os

W, H = 1680, 1050

# central node
CX, CY = 840, 525
CENTER = dict(left=690, top=460, w=300, h=130)

# card definitions: key -> geometry, colour, title, subtitle, bullets
CARDS = {
    "overall": dict(
        left=60, top=95, w=520, h=270, color="#2563EB",
        title="整体推进", sub="Overall Progress",
        anchor=("right", 235),
        bullets=[
            ("92 项风险清单合并统一跟踪，逐条闭环推进",
             "92 merged risks tracked to closure"),
            ("关键设计文件更新：P&ID · URS · 设备/仪表数据单",
             "Key deliverables updated"),
            ("3D 模型迭代优化，已上传 SharePoint 共享",
             "3D model iterated & shared on SharePoint"),
            ("HAZOP 前对齐持续推进，节奏可控",
             "Pre-HAZOP alignment on track"),
        ]),
    "resolved": dict(
        left=60, top=690, w=520, h=270, color="#16A34A",
        title="技术难点突破", sub="Key Issues Resolved",
        anchor=("right", 815),
        bullets=[
            ("SE01 气液分离器方案定型（清洗死角问题解决）",
             "SE01 separator design finalised"),
            ("设计温度 / 法兰等级确认，无需升级（降泄放压力即可）",
             "Design temp & flange class confirmed — no upgrade"),
            ("催化剂粉尘处理方案达成一致",
             "Catalyst dust handling agreed"),
            ("检漏策略明确（国内气密 + 现场氦检）",
             "Leak-test strategy agreed"),
        ]),
    "decisions": dict(
        left=1100, top=60, w=530, h=370, color="#DC2626",
        title="⚠ 需领导层决策", sub="Decisions Needed from Leadership",
        anchor=("left", 245), highlight=True,
        bullets=[
            ("产能 Capacity：CR01 DN80→DN50 或低于 URS 30 kg/day "
             "→ 保产能 or 调整 URS？",
             ""),
            ("工况 Duty：高/低压通用性 + 液气比(L:G) 未定 "
             "→ 明确工况优先级",
             ""),
            ("运输 Transport：整体吊装 vs 分撬模块化 → 方案决策"
             "（是否配合屋顶更换）",
             ""),
            ("认证 CE：现场组装需 CE-mark 人员赴英 3–4 周 "
             "→ 认证边界与成本",
             ""),
        ]),
    "scope": dict(
        left=1100, top=470, w=530, h=200, color="#EA7317",
        title="分工与协作", sub="Scope & Collaboration",
        anchor=("left", 545),
        bullets=[
            ("SW 主导：爆炸分区 · EA 排放计算 · 场地起重/照明",
             "SW-led: HAC · emissions · site lifting/lighting"),
            ("中方主导：撬块本体 · CE-MD/PED · 加料步梯",
             "CN-led: skid package · CE-MD/PED · charge steps"),
            ("双周技术会 + 双语 Meeting Track 稳定运行",
             "Biweekly review + bilingual tracker running"),
        ]),
    "next": dict(
        left=1100, top=710, w=530, h=250, color="#0E9AA7",
        title="下一步", sub="Next Steps",
        anchor=("left", 815),
        bullets=[
            ("冯博士向 SW 高层解释 CR01 产能（明日会议）",
             "Dr Feng to brief SW on CR01 capacity (tomorrow)"),
            ("SE01 做大方案测算 + 与 Rosemount 沟通探针",
             "Size up SE01 & consult Rosemount"),
            ("清洗方案下周出 → 对齐验收标准",
             "Cleaning plan next week → align criteria"),
            ("控制策略 / 联锁清单下周发出",
             "Control philosophy & interlock schedule next week"),
        ]),
}


def anchor_point(card):
    side, y = card["anchor"]
    if side == "right":
        return card["left"] + card["w"], y
    return card["left"], y


def center_edge(card):
    """Which point on the centre node to start the connector from."""
    ax, ay = anchor_point(card)
    if ax < CX:                       # card on the left
        return CENTER["left"], _clamp(ay)
    return CENTER["left"] + CENTER["w"], _clamp(ay)


def _clamp(y):
    lo = CENTER["top"] + 20
    hi = CENTER["top"] + CENTER["h"] - 20
    return max(lo, min(hi, y))


def connector_path(card):
    sx, sy = center_edge(card)
    ex, ey = anchor_point(card)
    dx = (ex - sx) * 0.5
    return f"M {sx},{sy} C {sx+dx},{sy} {ex-dx},{ey} {ex},{ey}"


def build_card_html(key, card):
    hl = card.get("highlight")
    body_rows = []
    for cn, en in card["bullets"]:
        en_html = f'<div class="en">{en}</div>' if en else ""
        body_rows.append(
            f'<li><span class="dot" style="background:{card["color"]}"></span>'
            f'<div class="txt"><div class="cn">{cn}</div>{en_html}</div></li>'
        )
    body = "\n".join(body_rows)
    cls = "card highlight" if hl else "card"
    return f'''
    <div class="{cls}" style="left:{card['left']}px;top:{card['top']}px;
         width:{card['w']}px;min-height:{card['h']}px;
         border-top:6px solid {card['color']};">
      <div class="card-head">
        <span class="ttl" style="color:{card['color']}">{card['title']}</span>
        <span class="sub">{card['sub']}</span>
      </div>
      <ul class="bullets">{body}</ul>
    </div>'''


paths = "\n".join(
    f'<path d="{connector_path(c)}" stroke="{c["color"]}" '
    f'stroke-width="3" fill="none" opacity="0.55"/>'
    for c in CARDS.values()
)
# small nubs at anchor points
nubs = "\n".join(
    f'<circle cx="{anchor_point(c)[0]}" cy="{anchor_point(c)[1]}" r="5" '
    f'fill="{c["color"]}"/>'
    for c in CARDS.values()
)

cards_html = "\n".join(build_card_html(k, c) for k, c in CARDS.items())

html = f'''<!doctype html><html lang="zh"><head><meta charset="utf-8">
<style>
  * {{ box-sizing:border-box; margin:0; padding:0; }}
  html,body {{ width:{W}px; height:{H}px;
     font-family:"Microsoft YaHei","Noto Sans CJK SC",sans-serif;
     background:linear-gradient(135deg,#f4f7fb 0%,#eaf0f8 100%); }}
  .stage {{ position:relative; width:{W}px; height:{H}px; }}
  .title {{ position:absolute; left:0; top:22px; width:100%;
     text-align:center; }}
  .title h1 {{ font-size:30px; color:#00366A; letter-spacing:1px; }}
  .title .rng {{ font-size:15px; color:#5b6b83; margin-top:4px;
     font-weight:400; }}
  svg {{ position:absolute; left:0; top:0; }}
  .center {{ position:absolute; left:{CENTER['left']}px;
     top:{CENTER['top']}px; width:{CENTER['w']}px; height:{CENTER['h']}px;
     border-radius:18px;
     background:linear-gradient(135deg,#00366A 0%,#12539e 100%);
     color:#fff; display:flex; flex-direction:column; align-items:center;
     justify-content:center; box-shadow:0 10px 26px rgba(0,54,106,.35);
     text-align:center; padding:10px; }}
  .center .big {{ font-size:23px; font-weight:700; line-height:1.25; }}
  .center .en {{ font-size:12.5px; opacity:.85; margin-top:6px;
     letter-spacing:.5px; }}
  .center .tag {{ font-size:12px; margin-top:8px;
     background:rgba(255,255,255,.16); padding:3px 12px; border-radius:20px; }}
  .card {{ position:absolute; background:#fff; border-radius:14px;
     box-shadow:0 6px 20px rgba(30,50,80,.12); padding:14px 18px 16px; }}
  .card.highlight {{ box-shadow:0 8px 26px rgba(220,38,38,.28);
     background:#fff7f7; }}
  .card-head {{ display:flex; align-items:baseline; gap:10px;
     border-bottom:1px solid #eef1f6; padding-bottom:8px; margin-bottom:8px; }}
  .card-head .ttl {{ font-size:19px; font-weight:700; }}
  .card-head .sub {{ font-size:12px; color:#8a97ab; }}
  ul.bullets {{ list-style:none; }}
  ul.bullets li {{ display:flex; align-items:flex-start; gap:9px;
     margin:7px 0; }}
  .dot {{ width:8px; height:8px; border-radius:50%; margin-top:6px;
     flex:0 0 auto; }}
  .txt .cn {{ font-size:14.5px; color:#22303f; line-height:1.35; }}
  .txt .en {{ font-size:11.5px; color:#8a97ab; line-height:1.3;
     margin-top:1px; }}
  .card.highlight .txt .cn {{ font-weight:600; color:#7f1d1d; }}
  .foot {{ position:absolute; left:0; bottom:14px; width:100%;
     text-align:center; font-size:11.5px; color:#9aa6b8; }}
</style></head>
<body>
  <div class="stage">
    <div class="title">
      <h1>连续氢化建设项目 · 双周进展汇报</h1>
      <div class="rng">Continuous Hydrogenation Skid — Bi-weekly Progress
        &nbsp;|&nbsp; 2026-07-01 ~ 07-15</div>
    </div>
    <svg width="{W}" height="{H}">{paths}{nubs}</svg>
    <div class="center">
      <div class="big">连续氢化项目<br>0701–0715 进展</div>
      <div class="en">Progress Snapshot</div>
      <div class="tag">整体节奏可控 · On track</div>
    </div>
    {cards_html}
    <div class="foot">凯莱英 Asymchem · Sandwich Site &nbsp;|&nbsp;
      IEPE · CFCT · UK Sandwich</div>
  </div>
</body></html>'''

out_html = "progress_mindmap_0715.html"
with open(out_html, "w", encoding="utf-8") as f:
    f.write(html)
print("Saved:", out_html)

# render to PNG with headless chrome
out_png = "progress_mindmap_0715.png"
abspath = os.path.abspath(out_html)
cmd = [
    "google-chrome", "--headless", "--no-sandbox", "--disable-gpu",
    "--hide-scrollbars", "--force-device-scale-factor=2",
    f"--window-size={W},{H}",
    f"--screenshot={os.path.abspath(out_png)}",
    f"file://{abspath}",
]
res = subprocess.run(cmd, capture_output=True, text=True)
print("chrome rc:", res.returncode)
if res.returncode != 0:
    print(res.stderr[-800:])
print("Saved:", out_png)
