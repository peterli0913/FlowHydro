# 今天与 Keith 交流内容 · Talking Points for Today's Meeting with Keith
**2026-07-23 · Continuous Hydrogenation Skid (T-0002)**

> 会议原则（内部准备会共识）｜ Meeting principle (from internal prep):
> 快速过，只讨论"需要 Keith 拍板 / 澄清 / 我方推进不下去"的条目；清洗方案初版内部仍在核对，**暂不发完整版**；7/1–7/2 的风险清单待我方全部关闭后，**统一给 Keith 过一遍确认关闭**。
> Keep it quick — only raise items that need Keith's decision / clarification, or that we cannot progress on our own. The cleaning plan (draft) is still under internal review and will **not** be shared yet. The 7/1–7/2 risk-register items will be **batch-reviewed with Keith for closure** once we have closed them internally.

---

## A. 需 Keith 决策 / 确认 · Decisions & Confirmations

### A1 ⭐ 设计温度 vs 泄放温度 / 法兰等级（今日重点技术议题）
**Design temperature vs relief temperature / flange class (key technical topic today)**
- 现状：为保 **CLASS 900**，拟将安全阀起跳压力与分离器设计压力调整为 **7.5 MPaG**（R42/R57, 07/14）。
- 数据显示：起跳压力降到 7.5 MPa 时，泄放温度只能降到 **~267 °C，降不到 250 °C**；若要降到 250 °C，压力需降到约 4 MPa 级，与需求相差太大。316L、CLASS 900 在 300 °C（甚至 375 °C）下 GR2.3 许用约 **78 bar**，略低于 **80 bar** 需求。
- **请 Keith 明确**：泄放（异常瞬态）工况是否必须被"正常设计温度"覆盖？（国内做法：泄放异常工况不与正常设计温度绑定；英方倾向更严格考虑。）我方今日**现场展示不同温度对应的安全阀起跳压力数据**供讨论。
- EN: To keep CLASS 900 we propose setting the PSV set pressure & separator design pressure to **7.5 MPaG**. However, at 7.5 MPa the relief temperature only falls to **~267 °C (not 250 °C)**; reaching 250 °C would need ~4 MPa, far below requirement. 316L CLASS 900 at 300 °C gives ~**78 bar** allowable, just under the 80 bar target. **Please confirm** whether the (transient) relief case must be bounded by the normal design temperature, or can be treated as a short-duration excursion. We will **show the temperature-vs-set-pressure data live** today.

### A2 ⭐ 反应器 CR01 管径 → 产能 · CR01 diameter → capacity (R120)
- 07/15 暂定 **DN80 → DN65**；待 **9 月**研发提供 DN65 反应器实验数据后，最终确认为 **DN65 或 DN50**。
- 请 SW / 领导层**知会并认可**这一"先按 DN65、9 月用数据定案"的路线（关系 URS 30 kg/day 产能）。
- EN: Tentatively **DN80 → DN65** (07/15); final size (**DN65 or DN50**) to be confirmed after R&D provides DN65 test data in **September**. Please **acknowledge** this data-driven plan (impacts the URS 30 kg/day capacity).

### A3 SE01 气液分离器方案 · SE01 separator design (R42)
- 已定方向：液位计**顶装内伸**、分离器采用**平顶**以尽量多布管口，管口不够时**适当加大直径**；安全阀起跳压力/设计压力调 **7.5 MPaG** 以保 CLASS 900。
- 请 Keith 确认该方向可接受。
- EN: Direction set — **top-mounted internal** level transmitter, **flat-top** vessel to fit more nozzles, **enlarge diameter** if needed; set/design pressure at **7.5 MPaG** for CLASS 900. Please confirm acceptable.

---

## B. 需 Keith 澄清 / 给输入 · Clarifications & Inputs Needed

### B1 操作模式与故障位 / 模式切换保护 · Operating modes, fail positions & mode-transition safeguards (R52/R53)
- 分离器浸泡清洗易触发高液位联锁 → 拟**增大排净管径实现淋洗**，并在**清洗模式下绕开高液位联锁**。
- 请 Keith 确认需求：原则是**尽量减少可被绕过的联锁**；若清洗必须绕过，采用严格程序 / 钥匙开关控制。
- EN: Soak cleaning of SE01 trips the high-level interlock → we plan to **enlarge the drain line for shower rinsing** and **bypass the high-level interlock only in cleaning mode**. Please confirm the requirement (minimise defeatable interlocks; if unavoidable, strict procedural / key-switch control).

### B2 模块化（分撬）设计 · Modular / split-skid design (R70)
- 现方案按**进口电梯尺寸**拆成 **4 个撬块**运输组装；**SW 建议**考虑**从屋顶直接整体吊装**（可配合屋顶更换）→ 需在周三会 / 领导层决策。
- EN: Current plan splits into **4 skids** sized to the imported lift; **SW suggests** evaluating a **direct whole-unit lift through the roof** (aligning with roof replacement) → needs Wednesday / leadership decision.

### B3 现场组装 vs CE / UKCA 符合性声明 · Site assembly vs CE/UKCA DoC (R71)
- 需与 **CE 认证机构（Notified Body）**确认细节：认证机构是否需**到 UK 现场验证后再发证**。
- EN: To confirm with the **Notified Body** whether on-site verification in the UK is required before issuing the certificate.

### B4 系统检漏 / 氦检漏 · Leak / helium test (R67)
- 国内做**气密性试验**；整撬分体运英、现场重装后需**再做泄漏试验**（现场可做氦检漏）。将与 **CE 认证机构确认具体方案**。
- EN: Pneumatic test in China; after split shipment & UK reassembly a **re-test** is needed (site can do helium leak test). We will **confirm the plan with the Notified Body**.

### B5 精滤器 FTR9832 堵塞 · Polishing filter blockage (R80)
- 目前"待定"→ 请 Keith 明确期望（差压监测已加联锁；溶剂清洗覆盖粉尘）。
- EN: Currently "TBD" → please clarify expectation (DP monitoring interlock added; solvent cleaning covers dust).

### B6 清洗验证 / 确认标准 · Cleaning validation / acceptance criteria (R95/R37)
- 清洗方案**初版完成**，内部核对后再发；**验收标准与检测方法请 SW（Keith）侧给出期望**（反应柱/分离器高压无法配视镜，靠取样分析残留浓度）。
- EN: Cleaning plan **draft done**, to be issued after internal review; please have **SW define the acceptance criteria & verification method** (no sight glass on high-pressure vessels — verify by residue sampling analysis).

### B7 减压 / 背压控制器 · Pressure reducer / back-pressure controllers
- **R113 PIC030**（分离器液相出口）：目前选**机械式减压阀**，无 SIL 认证 → 请确认能否满足 SIL / 风险评估假设的独立性。
- **R112 PIC028**：阀前背压 **0.9–6 MPaG**，量程比不宜过大。
- EN: **PIC030** now a **mechanical regulator** (no SIL cert) → confirm SIL/independence acceptability. **PIC028** pre-valve back-pressure 0.9–6 MPaG, turndown limited.

### B8 催化剂 OEB3 偏离 URS · Catalyst OEB3 deviation from URS (R63)
- 设计基于无粉尘颗粒催化剂、**OEB3 密闭**（相对 URS 存在偏离）→ 请 Keith 确认**接受该偏离**。
- EN: Design based on dust-free pellets with **OEB3 containment (a deviation from URS)** → please confirm acceptance.

---

## C. 新增待澄清（07/23）· New Items to Raise (07/23)

### C1 SE01 压力循环次数 · SE01 pressure cycles (R121)
- 气液分离器 SE01 从**常压 ↔ 60 barG** 的压力循环次数是多少？（用于疲劳设计）→ 请 Keith 提供预期循环次数。
- EN: How many **atmospheric ↔ 60 barG** pressure cycles will SE01 see? (for fatigue design) → please provide the expected cycle count.

### C2 压力表根部阀 · Root valves on pressure gauges (R122)
- P&ID 中压力表无根部阀，是否需要增加根部阀？
- EN: Pressure gauges on the P&ID have no root valves — are root valves required?

### C3 PCV028 卸压速率 · PCV028 depressurisation rate (R123)
- PCV028 孔径小，系统卸压时即使全开也很慢；**建议将 PCV028 旁通手阀改为开关阀**。
- EN: PCV028 orifice is small — depressurisation is slow even fully open; **propose changing the PCV028 bypass manual valve to an on-off valve**.

---

## D. 分工 / 边界（基本已确认，简要过）· Scope / Boundaries (largely confirmed)
- **通道 / 起重 / 照明 (R64/65/66)**：我方 3D 定稿后由 **SW 推进**（已确认）。/ Access, lifting aids, lighting — SW to proceed after our 3D is finalised (confirmed).
- **催化剂装卸 (R75/76/77/78/91/92)**：**SW 负责**装卸方案（07/15 确认）。/ Catalyst charge/discharge scheme — SW responsible (confirmed 07/15).
- **排放量 / EA 许可 (R118)**：**SW 负责**计算，我方提供到汇集管前的泄放数据。/ Emissions / EA permit — SW-led; we provide relief data up to the header.
- **法规合规 (R119)**：SW 内部提醒事项；**CFCT 核实加料步梯是否包含**；**SW 提供步梯/平台尺寸**。/ Regulatory — SW internal reminders; CFCT to verify feed staircase scope; SW to provide staircase/platform dimensions.
- **文档 (R60/R117)**：其他设计文件已完成，**仅剩操作手册（O&M）编制中**。/ Docs — all deliverables done except the **O&M manual (in progress)**.

---

## E. 收尾 · Wrap-up
- 7/1–7/2 已答复的风险项，待我方内部全部关闭后**统一提交 Keith 复核关闭**（统一标注颜色）。
- 今日执行不下去、需 Keith 输入的项（B / C 各条）请当场给方向或约定回复时间。
- EN: The answered 7/1–7/2 risk items will be **submitted to Keith for batch closure review** once internally closed. For items in B / C that we cannot progress alone, please give direction today or agree a response date.
