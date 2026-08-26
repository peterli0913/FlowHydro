# FlowHydro

凯莱英 UK Sandwich site 中试工厂连续氢化撬块项目的进度跟踪与可视化仓库。

## Cursor Cloud specific instructions

这是一个**文档 + 静态可视化生成仓库**，不是常驻服务/后端应用。没有 server、数据库、端口、测试、lint 或 CI。核心产物是自包含的静态 HTML（Gantt 时间线）和一个 `.pptx` 简报。

- **构建/运行命令**（见 `README.md`）：从仓库根目录运行 `python3 scripts/build_timeline_view.py`（中文版 → `timeline_view.html`）与 `python3 scripts/build_timeline_view_en.py`（英文 UK 版 → `timeline_view_en.html`）。另有 `python3 scripts/build_sw_briefing_slide.py` 生成 `SW_Sandwich_Progress_0710_Briefing.pptx`。
- **必须从仓库根目录运行**：脚本用相对路径读取根目录下的 `.xlsx`/`.docx` 源文件并把输出写回根目录。
- `build_timeline_view_en.py` 通过 `from build_timeline_view import ...` 复用中文脚本逻辑；用 `python3 scripts/build_timeline_view_en.py` 调用时可正常导入（Python 会把脚本所在的 `scripts/` 目录加入 `sys.path`）。
- **依赖**：`openpyxl`（所有脚本）、`python-pptx`（仅简报脚本，`README` 未列出）、`python-docx`（`README` 列出，但当前脚本未实际 import）。这些由 startup 更新脚本安装。
- **查看产物**：HTML 是自包含的，可直接用浏览器打开 `file://`，或从仓库根起 `python3 -m http.server 8000` 后访问 `http://localhost:8000/timeline_view.html`。Gantt 时间线跨 2026–2027，很宽，色条按各自日期列稀疏分布——大片空白是正常的，不是渲染 bug；右下角缩略图可看到全部色条。
- **产物变更提示**：HTML 输出是确定性的（内容不变则字节相同）；`.pptx` 因内嵌时间戳，每次运行都会产生 diff，即使内容未变——测试后如不想提交可 `git checkout --` 还原。
