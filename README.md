# FlowHydro

凯莱英 UK Sandwich site 中试工厂连续氢化撬块项目的进度跟踪与可视化。

## 当前文件

| 文件 | 说明 |
|---|---|
| `0509-Sandwich Continuous Hydrogen Timeline.xlsx` | 原始时间表（含交付物 / 设备投资 / 认证费用等多个 sheet） |
| `05-09 内部会议. 控制策略与防爆分区讨论.docx` | 2026-05-09 内部会议（上半场）转写 |
| `2026-05-09 15.47 录音.docx` | 2026-05-09 内部会议（下半场，采购流程与 3D 模型）转写 |
| `timeline_view.html` | **中文版可视化页面**：以 Timeline sheet 为底，叠加 5/9 会议结论 / 决议 / 行动项 / 风险，含 Gantt + 缩略图 + 编辑功能。 |
| `timeline_view_en.html` | **English Gantt-only view for UK colleagues**：纯英文、只含甘特图 + 图例 + 缩略图，不含内部会议面板。|
| `scripts/build_timeline_view.py` | 生成 `timeline_view.html` 的脚本。 |
| `scripts/build_timeline_view_en.py` | 生成 `timeline_view_en.html`（UK 英文版）的脚本，复用前者的数据加载逻辑。 |

## 重新生成可视化

```bash
pip install openpyxl python-docx
python3 scripts/build_timeline_view.py        # 中文内部版
python3 scripts/build_timeline_view_en.py     # 英文 UK 展示版
```

生成完直接双击 `timeline_view.html` / `timeline_view_en.html` 在浏览器查看。
