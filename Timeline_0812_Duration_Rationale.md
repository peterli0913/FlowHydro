# Duration Rationale — Procurement, Shipping, Certification, Installation
# 工期依据说明 · 采购 / 运输 / 认证验收 / 安装确认

配套材料：`timeline_0812_master_schedule.html` · `Timeline_0812_Talking_Points.md`
用途：SW 追问"这一段为什么要这么久"时的逐项应答口径。

---

## 讲解总原则 ｜ How to frame every answer

**中文**
每一段都用同一句式回答：**这段的长度不是执行速度问题，而是取决于它必须等谁**。先说前置依赖，再说内部构成，最后说能不能压缩。避免只给一个天数，那会让人觉得是拍脑袋定的。

**English**
Answer every one of these the same way: **the length of a phase is not a question of how fast we work, it is a question of what it has to wait for.** State the dependency first, then what the window is made of, then whether it can be compressed. Giving only a number invites the question of whether it was estimated at all.

---

# A. 采购段 ｜ Procurement

## A1. 采购信息确认 — 4 周（2026/10 W3 – 11 W2）
**Procurement information confirmation — 4 weeks**

**为什么是这个时长 ｜ Why this long**

**中文**
这不是一次性动作，而是一个四步闭环：设备数据单反馈 → 厂家出图纸 → 与 3D 模型合并会审 → 确认无误后才算信息冻结。会上明确了它的起点不能早于 10 月中旬，因为它的上游是 9 月初 HAZOP、之后两周的工艺与仪表文件升版、以及 9 月底才到位的仪表厂家反馈。

**English**
This is not a single action but a four-step loop: datasheet comments returned → vendor issues drawings → joint review against the 3D model → information frozen only once that review closes. Its start cannot be earlier than mid-October, because upstream of it sit the early-September HAZOP, the two weeks of process and instrument document revision that follow it, and instrument vendor feedback which only lands at the end of September.

**能否压缩 ｜ Compression**

**中文**
可以压缩的部分只有会审这一轮。如果英方能在收到数据单后按约定窗口反馈，可省掉一轮往返。

**English**
The only compressible element is the review round itself. If comments come back within an agreed window, one cycle can be avoided.

---

## A2. 工程采购（阀门、仪表、静设备）— 约 5 个月（2026/09 W1 – 2027/02 W1）
**Engineering procurement — about 5 months**

**为什么是这个时长 ｜ Why this long**

**中文**
这 5 个月不是一个连续的采购动作，而是**分批次滚动**的。会上确认的批次逻辑是：

- **第一批（9 月起）**：反应器以外的设备、阀门、仪表。这些规格已定，不必等反应器，可以先发询价。
- **第二批（9 月底起）**：反应器。等研发数据定案 DN65 或 DN50 后再提交，因为规格一变价格就变，比价流程要整个重走。
- **第三批（采购信息确认后）**：管件、垫片、手动阀等，须等主设备与自控阀选型返回、详细设计完成后才能提料。

每一批内部都包含询价 → 比价 → 定厂家 → 厂家返资料 → 我方审查这一完整循环，单批就是数周量级。

**English**
The five months are not one continuous purchasing action — procurement is **released in batches**:

- **Batch 1, from September:** all equipment, valves and instruments other than the reactor. Their specifications are settled, so they do not need to wait for the reactor.
- **Batch 2, from late September:** the reactor, once the R&D data fixes DN65 or DN50. Releasing it earlier would change the price when the size changes and force the whole comparison process to be repeated.
- **Batch 3, after procurement information is confirmed:** pipe fittings, gaskets and manual valves, which can only be taken off once main equipment and control valve selections have returned and detailed design is complete.

Each batch contains a full cycle of enquiry → comparison → vendor selection → vendor document return → our review, and a single batch is a matter of several weeks.

**关键点 ｜ The point to land**

**中文**
之所以拉到 5 个月，是因为批次之间是串联的：后一批的料表来自前一批的返回结果。这不是效率问题，是信息依赖。

**English**
It runs to five months because the batches are serial: the bill of material for a later batch is derived from what came back from the earlier one. This is an information dependency, not an efficiency issue.

---

## A3. 3D 模型定版 → BOM → 制造图 — 各约 2 周，串联至 2027/01 W1
**3D freeze → BOM → fabrication drawings — about 2 weeks each, in series**

**为什么必须串联 ｜ Why these must be serial**

**中文**
会上专门澄清过这条链，因为上一版把它们排成了并行：

1. 3D 模型定版**之后仍有详细设计**——自控阀等实物尺寸由厂家返资料确定，模型要据此更新并出工程图。
2. BOM 必须在**采购信息返回、详细设计完成之后**才能生成。范工的原话是：加工用的材料表要等采购信息返回来之后才能出。
3. 制造图在 BOM 之后。

**English**
This chain was specifically clarified in the meeting, because the previous revision had it running in parallel:

1. **Detailed design continues after the 3D model is frozen** — actual dimensions of control valves and similar items come from vendor returns, and the model must be updated and engineering drawings produced accordingly.
2. The BOM can only be generated **after procurement information has returned and detailed design is complete**. The fabrication material list cannot be produced before vendor information is in.
3. Fabrication drawings follow the BOM.

**关键点 ｜ The point to land**

**中文**
这三项各自只有两周，短的不是它们，是它们必须排在厂家返资料之后。

**English**
Each of these is only two weeks — they are not the long items. What positions them late is that they must sit behind the vendor returns.

---

## A4. 设备制造、组装与调试 — 约 7 周（2027/01 W3 – 03 W1）
**Fabrication, assembly and debugging — about 7 weeks**

**中文**
包含单体制造、撬块框架组装、管道预制与安装、电气仪表接线、以及出厂前的功能调试。这个时长本身没有变化——与 7 月 29 日版本相同，只是整体后移。

**English**
This covers component manufacture, skid frame assembly, piping prefabrication and installation, electrical and instrument wiring, and pre-delivery functional debugging. The duration itself is unchanged from the 29 July revision — it has simply moved.

---

# B. 运输段 ｜ Export and Shipping

## B1. 出口手续与运输检验 — 2 周（2027/05 W2 – W3）
**Export formalities and transport inspection — 2 weeks**

**中文**
包括报关资料准备、商检 / 运输检验、包装与加固方案确认。之所以放在 CE 认证完成之后，是因为报关与检验需要引用最终的认证文件。

**English**
Covers customs documentation, inspection for transport, and confirmation of packing and securing arrangements. It sits after CE certification because the customs and inspection submissions reference the final certification documents.

---

## B2. 海运至英国 — 约 12 周（2027/05 W4 – 08 W3）
**Sea freight to the UK — about 12 weeks**

**为什么不是"一趟船的时间" ｜ Why this is not just sailing time**

**中文**
这 12 周是**门到门**的窗口，不是航程本身。构成是：订舱与排期 → 集港与装箱加固 → 海运航程 → 英国口岸清关 → 内陆运输至 Sandwich 现场。其中订舱排期和口岸清关都不是我们能完全控制的环节。

**English**
The twelve weeks is a **door-to-door** window, not the sailing time. It is made up of booking and sailing schedule, consolidation and securing at the load port, the ocean leg, customs clearance on arrival, and inland haulage to the Sandwich site. Booking availability and port clearance are not fully within our control.

**额外的不确定性 ｜ The added uncertainty this time**

**中文**
本次还叠加了**超尺寸问题**。撬块最高处约 3.6–3.7 m，而普通集装箱内高 2.35 m、高柜约 2.5–2.69 m，立装装不下。开顶框架箱可以装，但不密封、暴露在空气中有腐蚀风险，不建议直接采用；同时整橇运输还要满足陆运高度限制。目前正在协调进出口部核实超尺寸承运方案，方案确定前这一段按常规窗口预留。

**English**
There is an additional **oversized-load issue** this time. The tallest point of the skid is about 3.6–3.7 m, against 2.35 m internal height for a standard container and about 2.5–2.69 m for a high cube — it cannot ship upright. An open-top flat rack can take it, but it is unsealed and exposed, which brings a corrosion risk, so it is not recommended as-is; whole-skid transport also has to satisfy road height limits. We are working with the import and export department to confirm oversized carrier options, and until that is settled this leg is held at a conventional window.

**这里是可以谈的地方 ｜ Where the conversation should go**

**中文**
运输方案的选择同时影响这一段和现场段。如果整橇可行，现场无需拆装与重复 IQ/OQ，可复用约 90% 的 FAT 结果——压缩空间在现场段，不在海运段。

**English**
The transport decision affects both this leg and the site phase. If whole-skid shipment is feasible, the site avoids disassembly and repeat IQ/OQ and can reuse roughly 90% of the FAT results — the recoverable time is in the site phase, not in the ocean leg.

---

# C. 认证与工厂验收段 ｜ Certification and FAT

## C1. CE 认证 — 分两段，合计约 6 个月
**CE certification — two phases, about 6 months in total**

**第一段：资料评审（2026/09 W1 – 11 W4，约 13 周）**
**Phase 1: document review**

**中文**
从 3D 模型定版起同步准备认证资料，随后确定认证机构并进入评审。目前**认证机构尚未选定**，这是这一段的起点条件。这一段与设计收尾、采购并行，不占用串联工期。

**English**
Certification documents are assembled from the 3D freeze onwards, after which the notified body is appointed and the review begins. The **notified body has not yet been appointed**, and that is the gating condition for this phase. It runs in parallel with design close-out and procurement, so it does not consume serial float.

**第二段：单体设备与整橇评定（2027/02 W1 – 05 W1，约 13 周）**
**Phase 2: equipment and whole-skid assessment**

**中文**
整橇的 CE 评定**必须在撬块组装完成之后**——这是本次修订纠正的一个顺序错误，上一版把它排得偏前。这一段包含设计审核、生产复验、测试见证、整改与发证。整改往返是其中不可省的部分。

**English**
Whole-skid CE assessment **can only follow completion of skid assembly** — correcting this ordering was one of the fixes in this revision, as the previous version had it too early. The phase covers design review, production re-inspection, witnessed testing, corrective actions and certificate issue. The corrective-action loop is the part that cannot be removed.

---

## C2. 整橇工厂验收 FAT — 约 4 周（2027/03 W2 – 04 W1）
**Integrated FAT — about 4 weeks**

**为什么在制造之后，且需要 4 周 ｜ Why it follows fabrication, and why four weeks**

**中文**
会上明确：**FAT 一定是制造完成之后才能开展**，上一版把它排在制造期内是不成立的。4 周的构成是：静态检查与文件核对 → 压力与密封试验 → 仪表回路与联锁测试 → 控制系统功能测试 → 缺陷整改与复测。

其中**整改与复测**是必须预留的——FAT 的价值恰恰在于把问题暴露在工厂而不是现场。

**English**
The meeting confirmed that **FAT can only start once fabrication is complete**; scheduling it inside the build period, as the previous revision did, is not valid. The four weeks comprise static inspection and document verification, pressure and leak testing, instrument loop and interlock testing, control system functional testing, and then defect correction and re-test.

The **correction and re-test** element has to be allowed for — the entire value of FAT is that problems surface in the factory rather than on site.

**与英方直接相关的一点 ｜ Directly relevant to the UK side**

**中文**
FAT 做得越完整，现场 IQ/OQ 可抵扣的部分越多。按 Keith 的说法，约 90% 的测试可以在撬块未安装状态下完成。前提是 FAT 方案需由 UK QA 事先审批签字，否则现场无法据此减免测试。

**English**
The more complete the FAT, the more of the site IQ/OQ it can offset — around 90% of the tests can be performed on the skid before installation. The condition is that the FAT protocol must be pre-approved and signed off by UK QA, otherwise it cannot be used to reduce site testing.

---

# D. 现场安装与确认段 ｜ Site Installation and Qualification

> 这一段的时长直接取决于运输方案，讲的时候要把这个因果关系点明。
> The duration of this phase is a direct function of the transport decision — make that link explicit.

## D1. 撬块就位安装 — 2 周（2027/08 W4 – 09 W1）
**Skid installation — 2 weeks**

**中文**
包含卸车、吊装就位、基础固定与找平。若采用四撬块分体方案，此处还需增加撬块间的重新组装工作量。

**English**
Offloading, lifting into position, fixing to foundations and levelling. Under the four-module split option, reassembly between modules has to be added here.

---

## D2. 机、电、仪、管安装 — 约 6 周（2027/09 W2 – 10 W3）
**Mechanical, electrical, instrumentation and piping — about 6 weeks**

**中文**
包含撬块与上下游管线连接、公用工程接入、电气与控制接线、保温。分体方案下还要叠加撬块间管道与电缆的重新连接、重新布线。

**English**
Connection to upstream and downstream pipework, utility tie-ins, electrical and control wiring, and insulation. Under the split option, re-connection of inter-module pipework and re-routing of cabling is added on top.

---

## D3. 确认与验收测试 DQ / IQ / OQ — 约 3 周（2027/10 W4 – 11 W2）
**Qualification and acceptance testing — about 3 weeks**

**为什么只有 3 周 ｜ Why only three weeks**

**中文**
之所以能压到 3 周，前提是 FAT 结果可被复用。如果整橇运输、现场不拆装，这里主要是复核公用工程参数与接口，工作量有限。反之，若分体运输后现场重新组装，则**大量 IQ/OQ 需要重做**，这一段会显著拉长——这正是运输方案对总工期影响最大的地方。

**English**
Three weeks is achievable only because the FAT results can be reused. With whole-skid shipment and no site disassembly, this is largely a matter of verifying utility parameters and interfaces, which is limited work. If the unit ships in modules and is reassembled on site, **a substantial amount of IQ/OQ has to be repeated** and this phase lengthens considerably — which is exactly where the transport decision has its largest effect on the overall programme.

---

## D4. 试车与开车 — 1 周（2027/11 W3）→ 可投用
**Commissioning and start-up — 1 week → beneficial use**

**中文**
公用工程投用、系统吹扫置换、带介质试车，完成即达到可投用状态。

**English**
Utilities in service, purging and inerting, commissioning with process fluid; on completion the unit reaches beneficial use.

---

# E. 一页速查 ｜ One-page summary

| 事项 ｜ Item | 工期 ｜ Duration | 决定长度的因素 ｜ What sets the length |
|---|---|---|
| 采购信息确认 Procurement info confirmation | 4 周 | 数据单反馈 → 厂家出图 → 与 3D 合并会审的闭环 |
| 工程采购 Engineering procurement | ~5 个月 | 三个批次串联，后批料表来自前批返回结果 |
| 3D 定版 / BOM / 制造图 | 各 ~2 周 | 本身很短，位置取决于厂家返资料 |
| 设备制造组装调试 Fabrication | ~7 周 | 工期未变，仅整体后移 |
| 出口手续与检验 Export formalities | 2 周 | 需引用最终认证文件 |
| 海运 Sea freight | ~12 周 | 门到门窗口；叠加超尺寸承运待确认 |
| CE 资料评审 CE document review | ~13 周 | 与设计采购并行；机构尚未选定 |
| CE 整橇评定 CE whole-skid | ~13 周 | 必须在组装完成之后；含整改往返 |
| 整橇 FAT | 4 周 | 必须在制造完成之后；含整改复测 |
| 撬块就位 Skid installation | 2 周 | 分体方案需增加重新组装 |
| 机电仪管安装 MEIP | ~6 周 | 分体方案需增加重连与重新布线 |
| DQ / IQ / OQ | ~3 周 | 以 FAT 可复用为前提；分体则显著拉长 |
| 试车开车 Commissioning | 1 周 | — |

---

# F. 三句收束 ｜ Three closing lines

**中文**
1. 这些时长的长度，主要由**依赖关系**决定，不是由执行速度决定。
2. 唯一能实质压缩总工期的杠杆是**运输方案**——它同时决定现场是否需要重做 IQ/OQ。
3. 我们已经把并行能做的都并行了：CE 资料与设计收尾并行、非反应器设备提前采购、文件提前发英方预览。剩下的串联部分是真实依赖。

**English**
1. These durations are set by **dependencies**, not by how fast the work is done.
2. The one lever that materially shortens the programme is the **transport decision**, because it determines whether IQ/OQ has to be repeated on site.
3. Everything that can run in parallel already does: CE documentation alongside design close-out, non-reactor equipment released early, and documents shared with the UK ahead of formal issue. What remains in series is genuine dependency.
