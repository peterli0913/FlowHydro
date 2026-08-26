# FlowHydro

凯莱英 UK Sandwich site 中试工厂连续氢化撬块项目的资料汇总、进度跟踪与可视化。

本分支为**综合工作分支**：此前分散在多个分支的会议转写、纪要、跟踪表、时间线、汇报材料与生成脚本已全部合并到这里。写进展、做汇报、准备与 Keith 的讨论时，直接在本分支查阅与新增即可。

## 目录导航

### 事件跟踪表（Meeting Minutes Tracking Record）
按日期滚动升版，每个文件含历史 sheet。**最新为 `0812`**。
`0414` / `0421` / `0512` / `0519` / `0701` / `0703` / `0706` / `0723` / `0729` / `0730` / `0806` / `0812`

> `0414` 与 `0421` 原为同名文件的两个不同快照（分别为 0414、0430 页），合并时按内容拆分保留。

### 会议纪要（已整理，Markdown）
| 文件 | 会议 |
|---|---|
| `Meeting_Summary_0805_Internal.md` | 08-05 内部会：清洗方案与项目进展 |
| `Meeting_Summary_0806_Technical.md` | 08-06 技术会（Keith）：撬块运输与仪表设计 |
| `Meeting_Summary_0812_Briefing.md` | 08-12 汇报会：撬块分体与现场安装规划 |
| `Meeting_Summary_0813_Internal_Prep.md` | 08-13 内部预备会 |
| `Meeting_Summary_0813_Technical_Keith.md` | 08-13 技术会（Keith）：清洗方案与设备设计 |
| `Meeting_Summary_0818_Technical_Keith.md` | 08-18 技术会（Keith）：撬块设计与清洗方案 |

### 会议转写（原始录音转文字）
`05-09` / `05-12` / `07-01` / `07-06` / `07-14` / `08-05` / `08-06` / `08-12` / `08-13` / `08-18` 各场 `.docx`，以及 `2026-05-09`、`2026-06-09`、`2026-07-29`、`2026-08-13` 录音转写。

### 与 Keith 的会议材料
| 文件 | 用途 |
|---|---|
| `Keith_Agenda_0825.md` | 08-25 议程与交流要点（含时间分配与说服主线） |
| `Feng_PFD_Brief_0825.md` | 给冯博士的 PFD / 物料平衡会前说明 |
| `Keith_Talking_Points_0723.md` / `Keith_Talking_Points_0806.md` | 交流要点 |
| `keith_agenda_0806.html` / `keith_meeting_0806_en.html` / `keith_agenda_0806_mindmap.html` | 08-06 议程展示页与思维导图 |
| `outstanding_questions_for_keith.html` | 待澄清问题汇总 |

### 时间线（Timeline）
| 文件 | 说明 |
|---|---|
| `0812-Sandwich Continuous Hydrogen Timeline.xlsx` | **最新版**时间表 |
| `timeline_0812_master_schedule.html` | 0812 版总进度展示页 |
| `Timeline_0812_Talking_Points.md` | 讲解词（中英），含预判问答 |
| `Timeline_0812_Duration_Rationale.md` | 各阶段工期依据（应对质询用） |
| `timeline_view.html` / `timeline_view_en.html` | 中文内部版 / 英文 UK 展示版甘特图 |
| `0114` / `0509` / `0701` 版 `.xlsx`、`timeline_compare.png` | 历史版本与对比 |

### 隔周进度汇报
`Progress_Report_0715.pptx` / `0729` / `0811`，思维导图 `progress_mindmap_0715.html`、`progress_mindmap_0811.html`，另有 `SW_Sandwich_Progress_0710_Briefing.pptx`。

### CE 认证
`CE相关/`（供应商资料与项目要点）、`ce_mark_practice_experience.html`（经验汇报）、`ce_mark_practice_talking_points.md`（讲解词）。

### 技术专题与风险
`SE01_demister_vs_degasser.html`、`relief_valve_methodology_cn_vs_uk.html`、`internal_pre_hazop_meeting_agenda.md`、`internal_HAZOP_template.xlsx`、`Risk_Register_Bilingual.docx`、`T-2502 Project Risk Register v0.1.docx`、`Internal_Priority_Warning_Risks.docx`。

### SW 侧来文
`A_SW最新文档/`：URS、PLC I/O 表、仪表 / 管线 / 阀门清单、图纸 PDF。

## 生成脚本

脚本均**从仓库根目录**运行，读取根目录源文件并把产物写回根目录。

```bash
pip install openpyxl python-pptx python-docx
python3 scripts/build_timeline_view.py        # 中文内部版甘特图
python3 scripts/build_timeline_view_en.py     # 英文 UK 展示版
python3 scripts/build_timeline_0812.py        # 0812 版总进度页
python3 scripts/build_progress_ppt_0811.py    # 隔周进度汇报 PPT
```

其余脚本见 `scripts/`：跟踪表升版、议程生成、HAZOP 模板、思维导图等。
