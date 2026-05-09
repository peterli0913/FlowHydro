# FlowHydro

凯莱英 UK Sandwich site 中试工厂连续氢化撬块项目的进度跟踪与可视化。

## 当前文件

| 文件 | 说明 |
|---|---|
| `0509-Sandwich Continuous Hydrogen Timeline.xlsx` | 原始时间表（含交付物 / 设备投资 / 认证费用等多个 sheet） |
| `05-09 内部会议. 控制策略与防爆分区讨论.docx` | 2026-05-09 内部会议（上半场）转写 |
| `2026-05-09 15.47 录音.docx` | 2026-05-09 内部会议（下半场，采购流程与 3D 模型）转写 |
| `timeline_view.html` | **可视化页面**：以 Timeline sheet 为底，叠加 5/9 会议结论 / 决议 / 行动项 / 风险。直接在浏览器打开。|
| `scripts/build_timeline_view.py` | 生成 `timeline_view.html` 的脚本（每次更新基线后重跑即可）。 |

## 重新生成可视化

```bash
pip install openpyxl python-docx
python3 scripts/build_timeline_view.py
```

生成完直接双击 `timeline_view.html` 在浏览器查看。
