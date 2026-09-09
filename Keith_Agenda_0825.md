# 08-25 与 Keith 会议 · 议程与交流要点
Agenda & Talking Points — Session with Keith, 25 Aug 2026
Sandwich Continuous Hydrogenation Skid

---

## 会前必做：跟踪表补录

跟踪表 `0812` 页最新记录只到 **08/13**，**08/18 技术会的结论一条都还没录入**。上次（08/13）过表时就出现过反复确认"这条是不是已经过了"的情况，会前补录能省下会上时间。

需要补录的 08/18 结论：

| 条目 | 08/18 结论 |
|---|---|
| BT01 排净阀 | 取消一个排净阀，25→10 变径移到排净阀下游，排净阀保持 25 mm；**须确认位于系统最低点** |
| V9858 泵入口隔离阀 | 移到变径下游、改 **DN10**（解决 25 mm 法兰阀占空间问题）；工艺无影响 |
| 计量泵安全阀 | 倾向改用**泵内置液压侧安全阀**（消除工艺管线死腔），待泵厂家确认设定值与可靠性 |
| 整撬循环清洗 | 限制是**流量非压力**；我方按实际系统做水力计算后回复可达流量 |
| 喷淋球 | 三台**全部改固定式**，取消可旋转与可伸缩 |
| 撬块宽度 | Keith 明确**不能加宽**（紧贴楼梯井 + 二层楼梯消防要求） |
| 绿色接线箱 | Keith 建议拆成 **3 个小箱**分装各撬块，可预接线 |
| 分离器高度 | 若循环清洗经排净阀排放、不需自流排空，**或可降低安装高度** |

---

## 时间分配建议（总 60 分钟）

| 段 | 议题 | 时长 | 为什么排这个位置 |
|---|---|---|---|
| 0 | 开场：本次目标与时间分配 | 2 min | 前几次都超时且讲不完，开场先把边界立住 |
| **1** | **PFD 与反应器尺寸（含 TCU）** | **15 min** | 冯先涛只参加这一段，放最前面确保不被挤掉，讲完他可退出 |
| **2** | **高压垫片 / 密封** | **10 min** | 在设备下单关键路径上；本段主要是**向 SW 要三个输入**，短而明确 |
| **3** | **3D 优化 + 控制柜尺寸** | **28 min** | 本次主体、最易发散；控制柜是 3D 的子集，合并讲不单列 |
| 4 | 文件反馈机制（操作手册 / 控制说明 / 清洗方案） | 5 min | 他还没看，不现场展开内容，只定反馈窗口与方式 |

**贯穿全场的说服主线**：R130 要求 **3D 布置与控制说明必须在 9 月初 HAZOP 前与 Keith 达成一致**，目标是 **HAZOP 只做一次**。任何需要他表态的事项都挂在这条上。

---

# 段 1 ｜ PFD 与反应器尺寸（冯先涛在场）

## 要说什么

**开场定位（中）**
这一段请研发冯博士一起参加，因为 PFD 与物料平衡的**计算基准必须由研发确定**，工程侧无法独立完成。我们想在今天把基准和边界定下来，9 月初数据一到就能直接出图。

**English**
We have asked Dr Feng from R&D to join this item, because the **calculation basis for the PFD and mass balance has to come from R&D** — engineering cannot settle it alone. We would like to agree the basis and the boundaries today, so that the moment the data lands in early September we can issue directly.

## 要回答 Keith 08/12 提的三个问题

他当时的原话是 "as you make your reactor smaller, your surface area goes down, and it becomes harder to take away the same kilowatts of heat"，并追问 "what is the limiting factor? I think your heat-transfer area is probably the limiting factor."

| # | 他的问题 | 我方要给的答复口径 |
|---|---|---|
| 1 | 缩径后反应器的**预期热负荷是否下降**，还是处理量不变、仍要求同等移热能力？ | 处理量随缩径下降；缩径的**本意正是强化传热、消除中心高温点**——产热量正比于容积（DN80→DN65 降约 34%，→DN50 降约 61%），而换热面积只降 19% / 37.5%，**产热下降快于移热能力下降**，温度更可控 |
| 2 | 现有换热面积做得到吗？限制因素是不是换热面积？ | 需研发给出基准后核算；**方向上是单位容积传热能力提升、绝对可移除千瓦数下降** |
| 3 | 新的操作参数包络记在哪份文件？ | **写进 PFD**；按 DN65 / DN50 分别出两版 |

## TCU 连带影响（他明确要求 HAZOP 前要有方案或评估）

**要说的事实**
- 为反应器 1 配置的两台**大 TCU 已经采购**
- 缩径后该 TCU 可能选型过大，会长期在约 **10% 负荷**下运行，控制品质下降（调节阀贴座、可调比恶化、易震荡）——反而破坏缩径想解决的温度稳定性

**我方的初步判断（高少峰）**
若 TCU **循环泵为变频**、且机组本身设有**旁路 / 回流管**，低负荷通常在设计考虑范围内，问题不大；缺点是此阶段**温差难以确定**，可能比完全匹配的机组略高。

**要请 Keith 做的第一步**
> Could you confirm whether the circulation pump on the purchased TCU is **variable-frequency**, and whether the unit has an **internal bypass / return line**? That determines whether low-load operation is acceptable before we consider replacing it.

**两条出路摆出来，但不在今天定**
- 加旁路人为制造热负荷，让 TCU 落回设计工作区
- 换一台小 TCU（他估约 **£38,000**），空间上也更好安装，代价是多一台闲置大 TCU

## 一个要主动做的动作（回应他对流程的批评）

08/12 他和 Claire 指出：反应器尺寸变更**没有被当作正式变更登记、也没有配套影响评估**。建议这次主动说：

> We are treating the reactor size change as a **formal change with an impact assessment**, tracked under item R124. The TCU duty is one line on that register, not a separate ad-hoc issue. Dr Feng is participating in the assessment so that we identify the knock-on effects now rather than six or nine months later.

**中**：把 TCU 作为 **R124 变更影响清单**里的一个条目呈现，而不是孤立的技术问题——这样正好回应他对流程的意见。

## 本段要拿到的结论
1. 研发确认 PFD / 物料平衡的计算基准（按哪个产能、体积、组成）
2. Keith 确认已采购 TCU 是否变频、是否有内部旁路
3. 双方确认时间路径：**9 月初研发数据 → 定尺寸 → 出两版 PFD/物料平衡 → HAZOP 前完成 TCU 评估**

---

# 段 2 ｜ 高压垫片 / 密封（跟踪表 R121）

## 现状要如实说

SE01 设计温度 **270 ℃**、设计压力 **7.5 MPaG**。原拟 316L 金属环垫，问题是**垫片与法兰同为 316L、硬度无差**，高压涉氢下难以形成可靠密封，而**氦检漏灵敏度极高，微小泄漏即被检出**。

**已排除的路径（要一次讲完，避免他重复建议）**

| 材料 / 方案 | 排除原因 |
|---|---|
| 铜 | 原料呈酸性、且溶于乙酸乙酯，存在腐蚀风险 |
| 镍、铝、蒙乃尔 | 对乙酸乙酯均有腐蚀问题 |
| PTFE（316L 缠绕垫 / 包覆） | 仅可到 200 ℃，低于 270 ℃ 泄放工况 |
| PEEK | 硬度远低于 316L、属软质密封，高压下塑性变形，不构成金属对金属硬密封 |
| 哈氏合金法兰 + 软垫 | 技术可行但**成本过高**（法兰与管道、设备连接数量大） |
| Kalrez（他自己提的） | 耐腐蚀但不耐高温高压 |

**影响范围要说清**：不只是 SE01，**所有高压法兰连接**（反应柱、液位计接口、进出口法兰盖）都是同一问题，且处于**设备下单关键路径**上。

## 本段的核心：向 SW 要三个输入

不要把这一段讲成"我们还没找到"，而要转成"我方需要三个输入才能收敛"。

**输入一：原料的酸是什么、浓度多少**
> To narrow the material selection we need to know **which acid is present in the feed and at what concentration**. At the moment we only know the feed is acidic, and that is what is blocking the screening.

**中**：目前只知道"原料酸性"，这是材料筛选卡住的根因。酸的种类与浓度不明确，任何耐腐蚀材料都无法排除或选定。

**输入二：SE01 的压力循环次数（R121 的原始问题，SW 至今未反馈）**
> R121 originally asked how many pressure cycles SE01 sees from atmospheric to 60 barg. We still do not have that number, and it **defines the fatigue duty on the gasket**. Could SW confirm it?

**中**：这个数决定垫片的疲劳工况，7/23 就提出、至今未回。

**输入三：SW 侧的密封供应商资源**
> You said last time we are at the limit of conventional materials and need a specialist sealing company. Does SW have a **specialist sealing supplier you have used** that we can approach jointly? You mentioned Kalrez — are there other materials your sites have used in hydrogen service at this temperature?

## 可以摆出来一起讨论的备选路径

| 路径 | 说明 |
|---|---|
| **局部升级法兰材质** | 只在关键几处用哈氏合金 / 因科镍法兰配软垫，而非全部升级，控制成本 |
| **减少高压法兰数量** | 改焊接结构、把液位计接口改为其他形式，从源头减少密封点 |
| **重新审视 270 ℃ 的来源** | 270 ℃ 是**泄放工况**温度。请说明泄放工况的**持续时间与频次**——若为极短时、极低频，能否评估在该瞬态下允许可控泄漏（泄放本身即向安全阀排放）？<br>⚠ 他 08/06 曾明确表示"不能假设 270 ℃ 泄放而允许垫片提前失效"，所以**换成"持续时间与频次"来问**，不要重复问"能不能降设计温度" |

---

# 段 3 ｜ 3D 优化 + 控制柜尺寸（本次主体）

## 3.1 先确认图纸是否互换到位

**我方应已发出**：最新版 P&ID（供他画圈划分模块）、3D 模型 PDF（供他批注）
**Keith 应回**：模块划分圈图、U 形重排布局草图、紧急淋浴准确尺寸、HTF 穿墙接口位置

开场先确认这四项到位情况，缺哪项当场定日期。

## 3.2 讲法：先给核算结果，再看图

08/18 那场之所以发散，是因为双方都在口头描述布局。这次建议**先把两个数报出来**，再展开图纸。

**数一：U 形重排到底能省多少宽度**
他估计能省 **500–600 mm**。我方据其草图核算的实际值是 ______。

> Based on your sketch, we have calculated the width that the U-shape re-sequencing actually releases. It comes to ____ mm against your estimate of 500 to 600.

**数二：反应器模块整体移出的拆装工程量**
不只是断管，量化如下：
- **6 根 HTF + 2 根工艺管 = 8 根管**
- 反应器周围传感器电缆：多支温度、压力、TCU 温度
- 顶部控制部件需拆下、降下
- **电缆桥架要拆，所有相关电缆都要断开**
- 估计工时 ______ 人天，风险点 ______

> Detaching the reactor module is not only a piping exercise. It is 8 pipes, the cable tray, all the sensor cabling around the reactors, and the control components on top have to come down. Our estimate is ____ man-days each time. That is the cost we are trading against the space it releases.

## 3.3 我方必须守住的立场：两层布置 + 中间通道

三条理由，成体系讲一遍，不要逐条否定他的方案：

1. **内层设备本身需要接近** —— 冷凝器、预热器、容器在内层，顶部有大量法兰连接
2. **法兰需要不定期接近**（频次不高但确实需要）—— 气体泄漏排查、垫片失效更换、接头松动紧固；**在 300 mm 空间里换垫片非常困难**
3. **分体重组后的装配验证** —— 现场重新组装后，若没有通道，**无法保证这些连接装配到位**；这一点直接关系到 SAT 能否通过

**同时要承认的事实**：管路无法全焊接——分撬必然有法兰，与设备连接处也必须法兰；且若要焊接，人进不去也无法施焊。

## 3.4 我方可以接受、且应主动提出的部分

把这些主动摆出来，会议氛围会从"互相否定"转成"共同收敛"：

| 采纳项 | 价值 |
|---|---|
| **HTF 水平走后部 + 竖直上墙 + 穿墙板** | 腾出中间空间；法兰离墙 150–200 mm，人探身约 400 mm 可操作 |
| **HART 远程组态 / 标定** | 从需求侧缩小"必须可达"的范围，是缓解空间矛盾最现实的手段 |
| **后部只放低维护频次设备**（冷凝器、预热器） | 维护周期可放宽 |
| **非安全 I/O 走本地接线箱 + 传统 I/O** | 直接压缩现场拆接与重测量，也就是压缩 SAT 工作量 |
| **评估降低分离器安装高度** | 循环清洗经排净阀排放、不需自流排空；同时利好整撬降高与集装箱适配 |

**分离器降高这一条建议主动提**，因为它同时服务三个目标（维护空间、整撬降高、集装箱适配），是本次最容易达成一致的议题。

## 3.5 控制柜尺寸变更（并入本段讲）

**要汇报的结论**
| 柜体 | 处置 |
|---|---|
| **棕色大柜**（remote IS / remote I/O） | 位于淋浴喷头上方不冲突；需**上移**并由正方形改**长方形**，确保完全在淋浴上方。变更后尺寸为 ______ |
| **绿色小箱** | 当前位置占用淋浴使用空间，必须处理。Keith 建议**拆成 3 个小箱**分装各撬块、就近接线，可**预接线（pre-wire）**——这正好压缩他担心的"重测 80 个 I/O 点" |

**必须向 Keith 确认的一件事（08/13 内部会提出，至今未问）**
> When you said the lower panel is "in the way" — is that a **physical clash**, or is it that **water from the safety shower would spray onto the panel**? If it is spray, moving the panel up may not help much, and we would need shielding or a different location instead.

**中**：如果是溅水问题，上移作用有限，需要的是隔离或改位置。这个必须问清，否则改了也白改。

**要请 Keith 表态的一件事（最省事的解法）**
> Could the **eyewash / safety shower be relocated to just outside the door** on the right-hand side? Placing it outside the door is normally acceptable, and having it this close to the skid may not be ideal for safety either. If it can move, the panel location problem disappears.

**要澄清的数据差异**
绿色箱电缆数：孟工口径是 **16 根**（PIC-028 / PIC-030 / PIC-049 + 两个电磁阀）；Keith 数的是约 **21 根**（含回路 20、急停 21）。**先对齐这个数**，否则拆箱方案没法定。

**要说明的规范约束（影响柜体尺寸）**
本安柜目前下进线，但线数多、单侧不够；**上进线可能不合规**（规范一般要求侧进或下进，上进有漏水进柜风险）；改侧进则柜内需加桥架、**柜体宽度还要变大**。最终尺寸待厂家确认开孔位置后定。

## 3.6 顺带在本段关闭的两个仪表条目（R72）

- 流量控制器电缆 —— **Modbus 还是 4–20 mA？**
- 硬联锁开关阀回路 —— **本安还是隔爆（Ex d）？**

这两项 7/23 就提出待确认，孟工的 cable schedule、PLC 硬件图、I/O 分配表都卡在这里。

## 3.7 本段要拿到的结论
1. U 形重排是否采纳（据核算数据定）
2. 中间通道是否保留（我方立场 + 他的可移出模块方案，二选一或折中）
3. 模块划分边界确认（他的圈图）
4. 控制柜"碍事"性质澄清 + 洗眼器能否外移 + 绿色箱拆分是否可行
5. R72 两个仪表选型确认

---

# 段 4 ｜ 文件反馈机制（5 分钟，不展开内容）

## 讲法：不要开放式征求意见

Keith 08/18 说清洗方案**尚未阅读**、操作手册也未反馈（他们在调试分析仪、有部件损坏）。所以这一段**不要问"有什么意见"**，而要用"我方已解决 N 项 / 待你输入 M 项 / 请给反馈窗口"的结构。

## 我方已按 08/13 意见修改的（报关闭）

| 项 | 处理 |
|---|---|
| **催化剂卸料顺序** | 采纳你的经验，改为**溶剂洗 → 水洗 → 卸水润催化剂 → 溶剂洗**（原 4.2.1 与 4.2.2 交换）。理由记录为：绝不卸载带溶剂的催化剂，避免催化剂着火 |
| **喷淋球** | 三台全部改**固定式**，取消可旋转与可伸缩 |
| **排净阀与变径** | 取消一个排净阀、变径下移、V9858 改 DN10 |
| **精滤器** | 加旁通，清洗时打开旁通便于卸压（R80） |
| **SE01 清洁度验证** | 不增加视镜；冯博士提供国内清洁度验收限度的**方法**供 SW 参考，具体限值由 SW 属地 QA 定（R38） |

## 仍需 Keith 输入或我方论证的（点出来，不现场解决）

| 项 | 状态 |
|---|---|
| **惰化（inerting）** | 你提出喷淋溶剂有静电点火风险；你邮件里的方案是 CIP 泵吸入侧 manifold 接 **0.8 barg 氮气**、系统压控 **1 barg**、超压由 **PIC028** 排放。我方正在核算：① PIC028 在 1 barg 低压端能否可靠动作（其阀前背压范围为 0.9–6 MPaG）② 0.8 对 1 barg 仅 0.2 bar 余量是否足够 ③ 经 BT04 泄压路径是否通 |
| **低流速循环清洗的有效性** | 你指出 40 L/h 循环可能"只是把污染物在系统内搬运"。我方按实际系统（多为 10 mm 卫生管）做水力计算后给出可达流量与判据 |
| **R39 喷淋球流量与充液时间** | 最小可选流量约 **20 L/min**，SE01**在 25 秒内**即达高液位跳停，机械喷淋窗口极短。请说明 SE01 预期的清洗时间窗口。<br>注：结合 08/18 两项决定——喷淋球改固定式、整撬循环为**浸泡式**——该条的答复口径需一并给出 |
| **R37 清洁度验证数值标准** | 方法已定（取样分析残留浓度），**数值限值待 SW 属地 QA 给出** |
| **接头形式** | 会上你说 **cam lock**，邮件里写 **triclamp**，以哪个为准？分离器侧此前明确要求 **CLASS 900 法兰**，triclamp 是否只用于低压侧（BT01 / BT04）？ |
| **隔离阀与软管接头可接近性** | 你邮件最后一段专门提到这点，我方已作为正式输入纳入 3D 布置调整 |

## 要约定的机制

> Could we agree a **feedback window**: comments on the cleaning philosophy, control philosophy and operating manual back by ____, either by email or in a dedicated session? The cleaning items are numerous — on 13 August we only got through the first line of the tracker — so a **dedicated cleaning session plus email** would use our weekly slot better than working through them in the routine meeting.

**中**：清洗条目多，8/13 那次只过了第一条。建议**清洗剩余部分走专题会 + 邮件**，把例会时间留给影响进度的决策项。

---

# 收尾 ｜ 两句话

**中**
今天最需要拿到的是两件事：一是 PFD 的计算基准与 TCU 的处理路径，二是 3D 布置的方向定案。这两项都在 HAZOP 前置路径上——3D 与控制说明必须在 9 月初之前谈定，我们的目标是 HAZOP 只做一次。

**English**
The two things we most need out of today are the calculation basis for the PFD together with the route for the TCU, and a decision on the direction of the 3D layout. Both sit on the critical path into HAZOP — the 3D layout and the control philosophy have to be agreed before early September, because our objective is to run HAZOP once and only once.

---

## 附：本周跟踪表更新一览（0812 页，相对上一版）

| 行 | 内容 | 变化 |
|---|---|---|
| R38 | SE01 清洁度验证 | 新增 08/13：不加视镜，冯博士提供国内验收限度方法供 SW 参考；负责人加入 Xiantao feng |
| **R39** | 喷淋球流量 vs 充液时间 | **数值更正：约 2 L/min → 约 20 L/min**；SE01 在 25 秒内达高液位跳停 |
| R63 | 催化剂规格 / OEB 3 | 负责人改为 **SW**（由 SW 关闭，国内无需动作），蓝标取消 |
| R80 | 精滤器 FTR9832 | 新增 08/06：加旁通，清洗时打开便于卸压；蓝标取消 |
| R82 | 差压监测精度 | 状态 **NO → YES**（已关闭），蓝标取消 |
| R121 | SE01 垫片材质 | 新增 08/06：原料酸性、PEEK 不适用，继续寻找合适垫片；**仍标蓝** |
| R123 | PCV028 旁通手阀 | 从"单独发邮件沟通"改为"按现有 P&ID 配置" |
| R125 | DN65/DN50 与 PFD | **新增蓝色标记**（本次讨论项） |
| R47 / R53 / R72 / R119 / R128 | — | 蓝标取消（08/13 已过），但**状态仍为 NO**，尚未关闭 |

**注意**：R47、R53、R72、R119、R128 蓝标虽已取消，但状态还是 NO。R72 的两个仪表选型（Modbus/4–20 mA、本安/隔爆）建议在段 3 顺带关闭。
