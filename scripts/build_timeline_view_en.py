"""English Gantt-only view for UK colleagues.

This is a slimmed, English-only counterpart of `build_timeline_view.py`.
It only outputs the Gantt chart + legend + draggable mini-map — no
meeting decisions / actions / risks panels (those are internal-team
material in the Chinese version).

The data extraction (xlsx -> bars) is delegated to `build_timeline_view`
so we never get out of sync with the source spreadsheet.

Generates `timeline_view_en.html` at the repo root.
"""

from __future__ import annotations

import html
import json
import os
from typing import Dict, List, Tuple

# Pull data-loading + helpers from the Chinese version so the two stay
# in sync. Everything we import is language-neutral.
from build_timeline_view import (
    XLSX_PATH,
    TODAY,
    PALETTE,
    DEPENDENCIES,
    Cell,
    Item,
    load_items,
    col_to_month_week,
    date_to_col,
    slugify,
)

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
HTML_PATH = os.path.join(REPO_ROOT, "timeline_view_en.html")

LEGEND_EN: List[Tuple[str, str, str]] = [
    ("original", "Light blue", "Original plan (previous baseline)"),
    ("revised",  "Orange",     "Current revised plan"),
    ("update",   "Purple",     "Time node needing update"),
    ("cert",     "Green",      "CE notified body involvement"),
]

OWNER_LABEL_EN: Dict[str, str] = {
    "UK":              "UK",
    "SW":              "UK Sandwich (SW)",
    "IEPE":            "IEPE",
    "CFCT":            "CFCT",
    "CIMT":            "CIMT",
    "IEPE/UK":         "IEPE + UK",
    "IEPE/CFCT":       "IEPE + CFCT",
    "IEPE/CFCT/UK":    "IEPE + CFCT + UK",
    "IEPE/CFCT/CIMT":  "IEPE + CFCT + CIMT",
    "CFCT/IEPE/UK":    "CFCT + IEPE + UK",
    "CFCT/CIMT":       "CFCT + CIMT",
    "UK/CFCT/IEPE":    "UK + CFCT + IEPE",
    "UK/IEPE":         "UK + IEPE",
    "Import/Export":   "Import / Export",
}


def render_en(items: List[Item], col_start: int, col_end: int) -> str:
    # Month / week headers ----------------------------------------------------
    month_headers: List[Tuple[int, int, int]] = []
    cur_y, cur_m, _ = col_to_month_week(col_start)
    cur_count = 0
    cur_ym = (cur_y, cur_m)
    for c in range(col_start, col_end + 1):
        y, m, _ = col_to_month_week(c)
        if (y, m) != cur_ym and cur_count:
            month_headers.append((cur_ym[0], cur_ym[1], cur_count))
            cur_ym = (y, m)
            cur_count = 0
        cur_count += 1
    if cur_count:
        month_headers.append((cur_ym[0], cur_ym[1], cur_count))

    today_col = date_to_col(TODAY)
    n_cols = col_end - col_start + 1
    week_w = 18
    label_w = 380
    track_w = n_cols * week_w
    today_left_px = label_w + (today_col - col_start) * week_w

    # Stable keys -------------------------------------------------------------
    sections: Dict[str, List[Item]] = {}
    for it in items:
        sections.setdefault(it.section or "Other", []).append(it)

    item_key: Dict[str, str] = {}
    used = set()
    for it in items:
        base = slugify(it.name_en)
        k = base; i = 1
        while k in used:
            i += 1; k = f"{base}-{i}"
        used.add(k)
        item_key[it.name_en] = k

    bucket_label_en = {
        "original": "Original",
        "revised":  "Revised",
        "update":   "Update needed",
        "cert":     "CE involvement",
    }

    # Build rows --------------------------------------------------------------
    rows_html: List[str] = []
    for sec_key, sec_items in sections.items():
        rows_html.append(
            f'<div class="section-row">'
            f'<div class="section-label">{html.escape(sec_key)}</div>'
            f'<div class="section-bar" style="width:{track_w}px"></div></div>'
        )
        for it in sec_items:
            label_en   = it.name_en
            owner_lbl  = OWNER_LABEL_EN.get(it.owner, it.owner or "")
            key        = item_key[it.name_en]
            deps       = DEPENDENCIES.get(it.name_en, [])

            # Coalesce cells into bars
            cells_by_col: Dict[int, Cell] = {}
            for c in it.cells:
                ex = cells_by_col.get(c.col)
                if not ex or (ex.bucket == "note" and c.bucket != "note"):
                    cells_by_col[c.col] = c
            bars: List[Tuple[int, int, str, List[str]]] = []
            cur_start = None
            cur_bucket = None
            cur_notes: List[str] = []
            for c in range(col_start, col_end + 2):
                cell = cells_by_col.get(c)
                bucket = cell.bucket if cell else None
                if bucket == "note":
                    bucket = None
                if bucket != cur_bucket:
                    if cur_bucket and cur_start is not None:
                        bars.append((cur_start, c - 1, cur_bucket, cur_notes))
                    cur_bucket = bucket
                    cur_start = c
                    cur_notes = []
                if cell and cell.note:
                    cur_notes.append(cell.note)

            # Bar elements
            bar_html_parts = []
            first_bar_left = None
            for (s, e, bucket, notes) in bars:
                left = (s - col_start) * week_w
                width = (e - s + 1) * week_w
                if first_bar_left is None:
                    first_bar_left = left
                ys, ms, ws_ = col_to_month_week(s)
                ye, me, we_ = col_to_month_week(e)
                date_lbl = (
                    f"{ys}/{ms:02d} W{ws_} -> {ye}/{me:02d} W{we_}"
                    if (ys, ms, ws_) != (ye, me, we_) else f"{ys}/{ms:02d} W{ws_}"
                )
                tip_lines = [
                    f"[ITEM] {label_en}",
                    f"[TIME] {date_lbl} - {bucket_label_en.get(bucket, bucket)}",
                    f"[OWNER] {owner_lbl}",
                ]
                if deps:
                    tip_lines.append("[DEPENDS ON] " + " / ".join(deps))
                # Note: per-cell Chinese annotations from the xlsx are
                # intentionally NOT surfaced in the UK-facing view.
                tip = "&#10;".join(html.escape(t) for t in tip_lines)
                color = PALETTE.get(bucket, "#888")
                bar_id = f"xlsx-{s}-{e}-{bucket}"
                bar_html_parts.append(
                    f'<div class="bar bar-{bucket}" '
                    f'style="left:{left}px;width:{width}px;background:{color}" '
                    f'title="{tip}" data-bucket="{bucket_label_en.get(bucket, "")}" '
                    f'data-item="{key}" data-bar-id="{bar_id}" '
                    f'data-start-col="{s}" data-end-col="{e}" data-bucket-key="{bucket}" '
                    f'data-source="xlsx"></div>'
                )

            dep_keys = [item_key[d] for d in deps if d in item_key]
            note_html = (
                f'<div class="row-note" contenteditable="true" '
                f'data-item="{key}" data-default="" '
                f'data-placeholder="(click to add a note — saved in your browser)">'
                f'</div>'
            )
            first_bar_attr = (
                f' data-first-bar="{first_bar_left}"' if first_bar_left is not None else ""
            )
            rows_html.append(
                f'<div class="item-row" id="row-{key}" data-item="{key}"'
                f' data-default-deps="{html.escape(",".join(dep_keys))}"{first_bar_attr}>'
                f'  <div class="item-label" data-item="{key}" title="Click empty area to scroll to this row\'s time period and highlight its dependencies">'
                f'    <div class="item-en">{html.escape(label_en)}</div>'
                f'    <div class="item-owner">Owner: {html.escape(owner_lbl)}</div>'
                f'    <div class="row-deps">'
                f'      <span class="rel-label dep-label">↳ Depends on:</span>'
                f'      <span class="rel-list deps-list" data-kind="deps"></span>'
                f'      <button class="rel-edit-btn" data-kind="deps" title="Edit dependencies">✏️</button>'
                f'      <button class="rel-edit-btn" data-kind="bars" title="Add / modify / delete time periods">🕓</button>'
                f'    </div>'
                f'    <div class="row-influence">'
                f'      <span class="rel-label inf-label">↗ Impacts:</span>'
                f'      <span class="rel-list inf-list"></span>'
                f'    </div>'
                f'    {note_html}'
                f'  </div>'
                f'  <div class="item-track" style="width:{track_w}px" data-item="{key}">'
                f'    {"".join(bar_html_parts)}'
                f'  </div>'
                f'</div>'
            )

    # Header rows -------------------------------------------------------------
    month_header_html = []
    week_header_html  = []
    for (y, m, span) in month_headers:
        cls = "mh-cur" if (y, m) == (TODAY.year, TODAY.month) else ""
        month_header_html.append(
            f'<div class="month-cell {cls}" style="width:{span*week_w}px">'
            f'{y}/<span class="m">{m:02d}</span></div>'
        )
        for w in range(1, span + 1):
            week_header_html.append(
                f'<div class="week-cell" style="width:{week_w}px">w{w}</div>'
            )

    # Legend ------------------------------------------------------------------
    legend_html = []
    for (lkey, color_en, desc) in LEGEND_EN:
        legend_html.append(
            f'<span class="leg"><span class="leg-swatch" '
            f'style="background:{PALETTE[lkey]}"></span>{color_en} — {desc}</span>'
        )
    minimap_legend_html = []
    for (lkey, color_en, _desc) in LEGEND_EN:
        minimap_legend_html.append(
            f'<span><i style="background:{PALETTE[lkey]}"></i>{color_en}</span>'
        )

    # JSON metadata for JS ----------------------------------------------------
    items_meta = []
    default_deps_map: Dict[str, List[str]] = {}
    for it in items:
        k = item_key[it.name_en]
        items_meta.append({
            "key": k,
            "en":  it.name_en,
            "zh":  it.name_en,                        # JS expects 'zh' key; use English
            "section": it.section or "",
        })
        default_deps_map[k] = [
            item_key[d] for d in DEPENDENCIES.get(it.name_en, []) if d in item_key
        ]
    items_meta_json   = json.dumps(items_meta,        ensure_ascii=False)
    default_deps_json = json.dumps(default_deps_map,  ensure_ascii=False)

    time_slots = []
    for c in range(col_start, col_end + 1):
        y, m, w = col_to_month_week(c)
        time_slots.append({"col": c, "y": y, "m": m, "w": w,
                            "label": f"{y}/{m:02d} W{w}"})
    time_slots_json = json.dumps(time_slots, ensure_ascii=False)

    palette_json = json.dumps(PALETTE)

    # ----- CSS ---------------------------------------------------------------
    css = """
    :root { --label-w: %(label_w)dpx; --week-w: %(week_w)dpx; }
    * { box-sizing: border-box; }
    body { font-family: 'Helvetica Neue', Arial, sans-serif;
           margin: 0; color: #1f2937; background: #f5f7fa; }
    h1 { font-size: 18px; margin: 0; }
    .container { max-width: 1480px; margin: 0 auto; padding: 18px; }
    .page-head { background: linear-gradient(120deg, #0f3460, #16213e);
                  color: #fff; padding: 14px 18px; border-radius: 10px;
                  margin-bottom: 14px; }
    .page-head .sub { font-size: 12px; color: #cbd5e1; margin-top: 4px; }
    .panel { background: #fff; border-radius: 10px; padding: 14px;
             box-shadow: 0 1px 3px rgba(15,52,96,0.06); margin-bottom: 14px; }
    h2 { font-size: 14px; margin: 0 0 8px 0; color: #0f3460; }
    .legend { display: flex; gap: 16px; flex-wrap: wrap; font-size: 12px;
              color: #475569; align-items: center; }
    .leg { display: flex; align-items: center; gap: 6px; }
    .leg-swatch { width: 14px; height: 14px; border-radius: 3px;
                  display: inline-block; }

    .gantt-toolbar { display: flex; gap: 12px; flex-wrap: wrap;
                     align-items: center; justify-content: space-between;
                     padding: 4px 6px 10px 6px;
                     border-bottom: 1px solid #e2e8f0; margin-bottom: 8px; }
    .gantt-toolbar .left  { display: flex; gap: 16px; flex-wrap: wrap;
                            align-items: center; font-size: 12px; color: #475569; }
    .gantt-toolbar .right { display: flex; gap: 6px; }
    .toolbar-btn { background: #0f3460; color: #fff; border: none;
                   padding: 5px 10px; border-radius: 6px; cursor: pointer;
                   font-size: 12px; }
    .toolbar-btn:hover { background: #16213e; }
    .toolbar-btn.secondary { background: #fff; color: #0f3460;
                              border: 1px solid #cbd5e1; }
    .toolbar-btn.secondary:hover { background: #f1f5f9; }
    .toolbar-hint { font-size: 11px; color: #94a3b8; margin-top: 4px; }

    .gantt { background: #fff; border-radius: 10px; overflow: hidden;
             box-shadow: 0 1px 3px rgba(15,52,96,0.06); }
    .gantt-scroll { overflow: auto; max-height: 78vh; position: relative; }
    .gantt-grid { position: relative; }

    .month-row, .week-row { display: flex; position: sticky; z-index: 8;
                            background: #f1f5f9; }
    .month-row { top: 0; font-weight: 600; color: #0f3460; }
    .week-row  { top: 28px; color: #64748b; font-size: 11px; }
    .header-spacer { width: var(--label-w); flex-shrink: 0;
                      background: #e2e8f0; border-right: 1px solid #cbd5e1;
                      position: sticky; left: 0; z-index: 10; }
    .month-cell { padding: 6px 8px; border-right: 1px solid #cbd5e1;
                  text-align: center; flex-shrink: 0; }
    .month-cell.mh-cur { background: #fde68a; color: #92400e; }
    .month-cell .m { font-size: 13px; }
    .week-cell { border-right: 1px dashed #e2e8f0; padding: 2px 0;
                 text-align: center; flex-shrink: 0; }

    .section-row { display: flex; align-items: stretch; background: #eef2f7; }
    .section-label { width: var(--label-w); padding: 6px 10px;
                     font-weight: 600; color: #0f3460; flex-shrink: 0;
                     border-right: 1px solid #cbd5e1;
                     position: sticky; left: 0; z-index: 7;
                     background: #eef2f7; }
    .section-bar { flex-shrink: 0; }

    .item-row { display: flex; border-bottom: 1px solid #f1f5f9;
                min-height: 40px; position: relative; }
    .item-row:hover .item-label { background: #f1f5f9; }
    .item-row.row-selected { background: #fff7ed; }
    .item-row.row-selected .item-label { background: #ffedd5;
                                          box-shadow: inset 3px 0 0 #d97706; }
    .item-row.row-dep { background: #fef9c3; }
    .item-row.row-dep .item-label { background: #fef08a;
                                     box-shadow: inset 3px 0 0 #facc15; }
    .item-label { width: var(--label-w); padding: 8px 10px;
                  border-right: 1px solid #e2e8f0; flex-shrink: 0;
                  position: sticky; left: 0; z-index: 6;
                  background: #fff; cursor: pointer;
                  transition: background 0.15s ease; }
    .item-en { font-size: 13px; font-weight: 600; color: #0f172a;
                line-height: 1.35; word-break: break-word; }
    .item-owner { font-size: 11px; color: #0f766e; margin-top: 3px; }

    .row-deps, .row-influence { font-size: 11px; color: #6b7280;
                                 margin-top: 3px; line-height: 1.7;
                                 display: flex; flex-wrap: wrap; align-items: center;
                                 gap: 4px; }
    .rel-label { color: #6b7280; flex-shrink: 0; }
    .dep-label { color: #0369a1; }
    .inf-label { color: #be185d; }
    .rel-list  { display: inline-flex; flex-wrap: wrap; gap: 4px; flex: 1; }
    .chip { display: inline-flex; align-items: center; gap: 2px;
            background: #e0f2fe; color: #075985; padding: 1px 7px;
            border-radius: 9px; cursor: pointer; font-size: 11px;
            border: 1px solid transparent; }
    .chip:hover { background: #bae6fd; border-color: #0284c7; }
    .chip-inf { background: #fce7f3; color: #9d174d; }
    .chip-inf:hover { background: #fbcfe8; border-color: #db2777; }
    .chip-empty { color: #cbd5e1; font-style: italic; cursor: default; }
    .chip-empty:hover { background: transparent; border-color: transparent; }
    .chip-remove { color: #94a3b8; font-size: 13px; line-height: 1;
                   padding: 0 0 0 2px; }
    .chip-remove:hover { color: #dc2626; }
    .rel-edit-btn { background: none; border: 1px dashed #cbd5e1;
                    border-radius: 9px; padding: 1px 6px; cursor: pointer;
                    font-size: 11px; color: #64748b; }
    .rel-edit-btn:hover { background: #f1f5f9; border-color: #94a3b8;
                          color: #0f3460; }
    .rel-edit-btn.editing { background: #fef3c7; border-color: #d97706;
                            color: #92400e; }

    .row-note { font-size: 11px; color: #475569; background: #fefce8;
                padding: 4px 6px; border-left: 2px solid #facc15;
                margin-top: 4px; line-height: 1.55;
                border-radius: 0 4px 4px 0; outline: none;
                white-space: pre-wrap; min-height: 20px; }
    .row-note:focus { background: #fef3c7; border-left-color: #d97706;
                      box-shadow: 0 0 0 2px rgba(217,119,6,0.15); }
    .row-note.edited { background: #fef3c7; border-left-color: #ea580c; }
    .row-note:empty::before { content: attr(data-placeholder);
                               color: #cbd5e1; font-style: italic; }

    .item-track { position: relative; flex-shrink: 0; }
    .bar { position: absolute; top: 8px; height: 22px; border-radius: 3px;
           opacity: 0.92; cursor: help;
           transition: transform 0.1s ease, box-shadow 0.1s ease; }
    .bar:hover { transform: translateY(-1px);
                  box-shadow: 0 4px 8px rgba(0,0,0,0.15); z-index: 3; }
    .bar.bar-revised  { box-shadow: 0 0 0 1px rgba(0,0,0,0.08); }
    .bar.bar-original { opacity: 0.5; height: 7px; top: 29px; border-radius: 2px; }
    .bar.bar-update   { width: 7px !important; height: 38px !important; top:0;
                          border-radius: 2px; box-shadow: 0 0 0 1px #4c1d95; }
    .bar.bar-cert     { box-shadow: 0 0 0 1px #166534 inset; }
    .bar.bar-user     { outline: 1px dashed rgba(0,0,0,0.35); outline-offset: -1px; }
    .row-selected .bar.bar-revised, .row-selected .bar.bar-update,
    .row-selected .bar.bar-cert {
        box-shadow: 0 0 0 2px #d97706; transform: translateY(-1px);
    }

    .today-marker {
       position: absolute; top: 0; bottom: 0; width: 2px; background: #dc2626;
       left: %(today_left_px).1fpx;
       z-index: 4; pointer-events: none;
    }
    .today-label { position: absolute; top: 30px; transform: translateX(-50%%);
                   background: #dc2626; color: #fff; font-size: 11px;
                   padding: 1px 6px; border-radius: 3px; white-space: nowrap; }

    /* ---- popovers (deps + bars editor share the .dep-editor base) ---- */
    .dep-editor { position: absolute; z-index: 50; background: #fff;
                  border: 1px solid #cbd5e1; border-radius: 8px;
                  box-shadow: 0 8px 24px rgba(15,52,96,0.18);
                  padding: 10px 12px; width: 360px; font-size: 12px;
                  color: #1f2937; }
    .dep-editor .de-title { font-weight: 600; color: #0f3460;
                            margin-bottom: 6px; font-size: 12px; }
    .dep-editor .de-current { display: flex; flex-wrap: wrap; gap: 4px;
                              min-height: 22px; margin-bottom: 8px;
                              padding-bottom: 6px;
                              border-bottom: 1px dashed #e2e8f0; }
    .dep-editor .de-add-row { display: flex; gap: 6px; }
    .dep-editor select { flex: 1; padding: 4px 6px; border: 1px solid #cbd5e1;
                         border-radius: 4px; font-size: 12px; min-width: 0; }
    .dep-editor .de-actions { display: flex; justify-content: space-between;
                              margin-top: 8px; gap: 6px; }
    .dep-editor button.de-btn { background: #0f3460; color: #fff;
                                 border: none; padding: 4px 10px;
                                 border-radius: 4px; cursor: pointer;
                                 font-size: 12px; }
    .dep-editor button.de-btn.secondary { background: #fff; color: #0f3460;
                                           border: 1px solid #cbd5e1; }
    .dep-editor button.de-btn.danger    { background: #fff; color: #b91c1c;
                                           border: 1px solid #fecaca; }
    .dep-editor .de-empty { color: #94a3b8; font-style: italic; font-size: 11px; }

    .bar-editor .be-list { max-height: 240px; overflow-y: auto;
                            padding-right: 4px; margin-bottom: 8px;
                            padding-bottom: 6px;
                            border-bottom: 1px dashed #e2e8f0; }
    .bar-editor .be-row { display: flex; align-items: center; gap: 4px;
                           margin-bottom: 4px; font-size: 11px;
                           flex-wrap: nowrap; }
    .bar-editor .be-row select { padding: 3px 4px; border: 1px solid #cbd5e1;
                                  border-radius: 4px; font-size: 11px;
                                  min-width: 0; }
    .bar-editor .be-row .be-bucket, .bar-editor .be-row .be-row-bucket { width: 110px; flex-shrink: 0; }
    .bar-editor .be-row .be-start, .bar-editor .be-row .be-row-start { flex: 1; }
    .bar-editor .be-row .be-end,   .bar-editor .be-row .be-row-end   { flex: 1; }
    .bar-editor .be-arrow { color: #94a3b8; font-size: 10px; flex-shrink: 0; }
    .bar-editor .be-del-btn { padding: 2px 6px !important; flex-shrink: 0; font-size: 11px; }
    .bar-editor .be-add-btn { padding: 3px 8px !important; flex-shrink: 0; font-size: 11px; }
    .bar-editor .be-add { padding-top: 6px; border-top: 1px dashed #e2e8f0; }
    .bar-editor .be-src-tag { display: inline-block; font-size: 9px;
                               padding: 1px 4px; border-radius: 3px;
                               flex-shrink: 0; }
    .bar-editor .be-src-tag.tag-xlsx { background: #dbeafe; color: #1d4ed8; }
    .bar-editor .be-src-tag.tag-user { background: #fef3c7; color: #b45309; }

    /* ---- minimap ---- */
    .minimap { position: fixed; right: 20px; bottom: 20px;
               width: 360px; background: #fff;
               border: 1px solid #cbd5e1; border-radius: 8px;
               box-shadow: 0 10px 30px rgba(15,52,96,0.20);
               z-index: 200; user-select: none; }
    .minimap.collapsed .mm-body { display: none; }
    .minimap .mm-header { display: flex; align-items: center;
                           justify-content: space-between;
                           background: #0f3460; color: #fff;
                           padding: 6px 10px; border-radius: 8px 8px 0 0;
                           cursor: grab; font-size: 12px; }
    .minimap.dragging .mm-header { cursor: grabbing; }
    .minimap.collapsed .mm-header { border-radius: 8px; }
    .minimap .mm-title { display: flex; align-items: center; gap: 6px;
                          font-weight: 600; }
    .minimap .mm-icons { display: flex; gap: 4px; }
    .minimap .mm-icons button {
      background: rgba(255,255,255,0.15); color: #fff; border: none;
      padding: 2px 8px; border-radius: 4px; cursor: pointer;
      font-size: 11px;
    }
    .minimap .mm-icons button:hover { background: rgba(255,255,255,0.3); }
    .minimap .mm-body { padding: 8px; }
    .minimap .mm-canvas { position: relative; background: #f8fafc;
                           border: 1px solid #e2e8f0; border-radius: 4px;
                           overflow: hidden; }
    .minimap .mm-canvas svg { display: block; width: 100%%; height: 100%%; }
    .minimap .mm-viewport { position: absolute; border: 1.5px solid #dc2626;
                             background: rgba(220,38,38,0.10);
                             pointer-events: none; }
    .minimap .mm-canvas.draggable { cursor: crosshair; }
    .minimap .mm-legend { display: flex; flex-wrap: wrap; gap: 8px;
                           font-size: 10px; color: #475569; margin-top: 6px; }
    .minimap .mm-legend span { display: inline-flex; align-items: center; gap: 3px; }
    .minimap .mm-legend i { display: inline-block; width: 9px; height: 9px;
                             border-radius: 2px; }

    .toast { position: fixed; bottom: 20px; left: 50%%;
             transform: translateX(-50%%) translateY(20px);
             background: #0f3460; color: #fff; padding: 8px 16px;
             border-radius: 6px; font-size: 13px; opacity: 0;
             transition: opacity 0.2s ease, transform 0.2s ease;
             pointer-events: none; z-index: 1000; }
    .toast.show { opacity: 1; transform: translateX(-50%%) translateY(0); }
    """ % dict(label_w=label_w, week_w=week_w, today_left_px=today_left_px)

    # Inline JS — pure English UI strings, otherwise same logic as the
    # Chinese build (so editing parity is kept for UK colleagues).
    js = """
(function() {
  var LS_NOTE = "sandwich-en-note:";
  var LS_DEPS = "sandwich-en-deps:";
  var LS_BARS = "sandwich-en-bars:";
  var scroller = document.getElementById('ganttScroll');
  var grid     = document.getElementById('ganttGrid');
  var labelW   = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--label-w'), 10) || 380;
  var weekW    = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--week-w'), 10) || 18;
  var todayLeftPx = __TODAY_LEFT__;
  var trackW   = __TRACK_W__;
  var colStart = __COL_START__;
  var nCols    = __N_COLS__;
  var ITEMS     = JSON.parse(document.getElementById('items-meta').textContent);
  var DEFAULTS  = JSON.parse(document.getElementById('default-deps').textContent);
  var TIMESLOTS = JSON.parse(document.getElementById('time-slots').textContent);
  var PALETTE   = JSON.parse(document.getElementById('bucket-palette').textContent);
  var itemByKey = {}; ITEMS.forEach(function(it) { itemByKey[it.key] = it; });
  var slotByCol = {}; TIMESLOTS.forEach(function(s) { slotByCol[s.col] = s; });
  var BUCKETS = [
    {key: 'original', en: 'Original plan'},
    {key: 'revised',  en: 'Revised plan'},
    {key: 'update',   en: 'Update node'},
    {key: 'cert',     en: 'CE involvement'},
  ];
  var bucketByKey = {}; BUCKETS.forEach(function(b) { bucketByKey[b.key] = b; });

  function toast(msg) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(function() { t.classList.remove('show'); }, 1800);
  }

  // ===== Deps state =======================================================
  var effDeps = {};
  var infMap  = {};
  function loadDeps() {
    ITEMS.forEach(function(it) {
      var saved = null;
      try { saved = localStorage.getItem(LS_DEPS + it.key); } catch(e) {}
      if (saved !== null) {
        try { effDeps[it.key] = JSON.parse(saved); }
        catch(e) { effDeps[it.key] = (DEFAULTS[it.key] || []).slice(); }
      } else {
        effDeps[it.key] = (DEFAULTS[it.key] || []).slice();
      }
    });
    recomputeInfluence();
  }
  function recomputeInfluence() {
    infMap = {};
    Object.keys(effDeps).forEach(function(k) {
      effDeps[k].forEach(function(d) {
        if (!infMap[d]) infMap[d] = [];
        if (infMap[d].indexOf(k) < 0) infMap[d].push(k);
      });
    });
  }
  function saveDeps(key) {
    var def = (DEFAULTS[key] || []).slice().sort().join(',');
    var cur = (effDeps[key] || []).slice().sort().join(',');
    try {
      if (def === cur) localStorage.removeItem(LS_DEPS + key);
      else localStorage.setItem(LS_DEPS + key, JSON.stringify(effDeps[key]));
    } catch(e) {}
  }

  function chipHTML(key, kind) {
    var meta = itemByKey[key];
    if (!meta) return '';
    var label = meta.en;
    var cls = (kind === 'inf' ? 'chip chip-inf' : 'chip');
    return '<span class="' + cls + '" data-target="' + key + '">' + label + '</span>';
  }
  function renderRowRel(rowEl) {
    var key = rowEl.getAttribute('data-item');
    var depsList = rowEl.querySelector('.deps-list');
    var infList  = rowEl.querySelector('.inf-list');
    var deps = effDeps[key] || [];
    var infs = (infMap[key] || []);
    depsList.innerHTML = deps.length
      ? deps.map(function(k) { return chipHTML(k, 'dep'); }).join('')
      : '<span class="chip chip-empty">none</span>';
    infList.innerHTML = infs.length
      ? infs.map(function(k) { return chipHTML(k, 'inf'); }).join('')
      : '<span class="chip chip-empty">none</span>';
    var defKey = (DEFAULTS[key] || []).slice().sort().join(',');
    var curKey = deps.slice().sort().join(',');
    var btn = rowEl.querySelector('.rel-edit-btn[data-kind="deps"]');
    if (btn) btn.classList.toggle('editing', defKey !== curKey);
  }
  function renderAllRel() {
    document.querySelectorAll('.item-row').forEach(renderRowRel);
  }

  // ===== Click-to-jump ====================================================
  function clearHighlights() {
    document.querySelectorAll('.item-row').forEach(function(r) {
      r.classList.remove('row-selected'); r.classList.remove('row-dep');
    });
  }
  function scrollToItem(key) {
    var row = document.getElementById('row-' + key);
    if (!row) return;
    clearHighlights();
    row.classList.add('row-selected');
    (effDeps[key] || []).forEach(function(d) {
      var dr = document.getElementById('row-' + d);
      if (dr) dr.classList.add('row-dep');
    });
    var firstBar = row.getAttribute('data-first-bar');
    if (firstBar !== null && firstBar !== '') {
      var targetLeft = labelW + parseFloat(firstBar) - scroller.clientWidth * 0.30;
      scroller.scrollTo({ left: Math.max(0, targetLeft), behavior: 'smooth' });
    }
    row.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    updateMinimapViewport();
  }

  // ===== Notes (contenteditable) ==========================================
  function restoreNotes() {
    document.querySelectorAll('.row-note').forEach(function(el) {
      var key = el.getAttribute('data-item');
      var def = el.getAttribute('data-default') || '';
      var saved = null;
      try { saved = localStorage.getItem(LS_NOTE + key); } catch(e) {}
      if (saved !== null && saved !== def) {
        el.textContent = saved;
        el.classList.add('edited');
      }
    });
  }
  document.querySelectorAll('.row-note').forEach(function(el) {
    var key = el.getAttribute('data-item');
    el.addEventListener('blur', function() {
      var txt = el.innerText.trim();
      var dflt = (el.getAttribute('data-default') || '').trim();
      try {
        if (txt === dflt) {
          localStorage.removeItem(LS_NOTE + key);
          el.classList.remove('edited');
        } else {
          localStorage.setItem(LS_NOTE + key, txt);
          el.classList.add('edited');
        }
      } catch(e) {}
    });
    el.addEventListener('click', function(ev) { ev.stopPropagation(); });
  });

  // ===== Bar state ========================================================
  var xlsxBars = {}; var userBars = {};
  function snapshotXlsxBars() {
    document.querySelectorAll('.item-row').forEach(function(row) {
      var key = row.getAttribute('data-item');
      var list = [];
      row.querySelectorAll('.bar[data-source="xlsx"]').forEach(function(b) {
        list.push({
          id:        b.getAttribute('data-bar-id'),
          startCol:  parseInt(b.getAttribute('data-start-col'), 10),
          endCol:    parseInt(b.getAttribute('data-end-col'), 10),
          bucket:    b.getAttribute('data-bucket-key'),
          tip:       b.getAttribute('title') || '',
        });
      });
      xlsxBars[key] = list;
    });
  }
  function loadUserBars() {
    ITEMS.forEach(function(it) {
      var saved = null;
      try { saved = localStorage.getItem(LS_BARS + it.key); } catch(e) {}
      if (saved) {
        try { userBars[it.key] = JSON.parse(saved); }
        catch(e) { userBars[it.key] = {hidden: [], added: []}; }
      } else { userBars[it.key] = {hidden: [], added: []}; }
    });
  }
  function saveUserBars(key) {
    var ub = userBars[key];
    var isDefault = (!ub || (ub.hidden.length === 0 && ub.added.length === 0));
    try {
      if (isDefault) localStorage.removeItem(LS_BARS + key);
      else           localStorage.setItem(LS_BARS + key, JSON.stringify(ub));
    } catch(e) {}
  }
  function effectiveBars(key) {
    var ub = userBars[key] || {hidden: [], added: []};
    var hidden = new Set(ub.hidden);
    var xl = (xlsxBars[key] || []).filter(function(b) { return !hidden.has(b.id); })
                                  .map(function(b) { return Object.assign({source:'xlsx'}, b); });
    var ua = (ub.added || []).map(function(b) { return Object.assign({source:'user'}, b); });
    return xl.concat(ua).sort(function(a, b) { return a.startCol - b.startCol; });
  }
  function buildTip(rowEl, b) {
    var key = rowEl.getAttribute('data-item');
    var meta = itemByKey[key];
    var owner = (rowEl.querySelector('.item-owner') || {}).innerText || '';
    var s = slotByCol[b.startCol], e = slotByCol[b.endCol];
    var dateLabel = (s.col === e.col) ? s.label : (s.label + ' -> ' + e.label);
    var lines = ['[ITEM] ' + meta.en,
                 '[TIME] ' + dateLabel + ' - ' + (bucketByKey[b.bucket] ? bucketByKey[b.bucket].en : b.bucket),
                 '[' + owner + ']'];
    if (b.source === 'user') lines.push('[user-added]');
    return lines.join('\\n');
  }
  function renderRowBars(rowEl) {
    var key   = rowEl.getAttribute('data-item');
    var track = rowEl.querySelector('.item-track');
    track.querySelectorAll('.bar').forEach(function(b) { b.remove(); });
    var bars = effectiveBars(key);
    var firstLeft = null;
    bars.forEach(function(b) {
      var s = slotByCol[b.startCol], e = slotByCol[b.endCol];
      if (!s || !e) return;
      var left = (b.startCol - colStart) * weekW;
      var width = (b.endCol - b.startCol + 1) * weekW;
      if (firstLeft === null || left < firstLeft) firstLeft = left;
      var el = document.createElement('div');
      el.className = 'bar bar-' + b.bucket + (b.source === 'user' ? ' bar-user' : '');
      el.style.left  = left  + 'px';
      el.style.width = width + 'px';
      el.style.background = PALETTE[b.bucket] || '#888';
      el.setAttribute('title', b.tip || buildTip(rowEl, b));
      el.setAttribute('data-item', key);
      el.setAttribute('data-bar-id', b.id);
      el.setAttribute('data-start-col', b.startCol);
      el.setAttribute('data-end-col',   b.endCol);
      el.setAttribute('data-bucket-key', b.bucket);
      el.setAttribute('data-source', b.source);
      track.appendChild(el);
    });
    if (firstLeft !== null) rowEl.setAttribute('data-first-bar', firstLeft);
    else                    rowEl.removeAttribute('data-first-bar');
  }
  function renderAllRowBars() {
    document.querySelectorAll('.item-row').forEach(renderRowBars);
  }

  // ===== Deps editor =====================================================
  var editorEl = null;
  function closeEditor() {
    if (editorEl && editorEl.parentNode) editorEl.parentNode.removeChild(editorEl);
    editorEl = null;
    document.querySelectorAll('.rel-edit-btn[data-kind="deps"]').forEach(function(b) { b.disabled = false; });
  }
  function openEditor(rowEl, btn) {
    closeEditor(); closeBarEditor();
    var key = rowEl.getAttribute('data-item');
    var meta = itemByKey[key];
    editorEl = document.createElement('div');
    editorEl.className = 'dep-editor';
    editorEl.innerHTML =
      '<div class="de-title">Edit dependencies of "' + meta.en + '"</div>' +
      '<div class="de-current"></div>' +
      '<div class="de-add-row">' +
      '  <select class="de-add-select"><option value="">-- pick an upstream item to add --</option></select>' +
      '  <button class="de-btn de-add-btn">+ Add</button>' +
      '</div>' +
      '<div class="de-actions">' +
      '  <button class="de-btn danger de-reset-btn">Restore defaults</button>' +
      '  <button class="de-btn secondary de-close-btn">Done</button>' +
      '</div>';
    document.body.appendChild(editorEl);
    btn.disabled = true;
    var rect = btn.getBoundingClientRect();
    editorEl.style.top  = (window.scrollY + rect.bottom + 6) + 'px';
    var left = window.scrollX + rect.left;
    if (left + 360 > window.innerWidth - 12) left = window.innerWidth - 372;
    editorEl.style.left = Math.max(12, left) + 'px';

    var curBox   = editorEl.querySelector('.de-current');
    var selectEl = editorEl.querySelector('.de-add-select');
    var addBtn   = editorEl.querySelector('.de-add-btn');
    var closeBtn = editorEl.querySelector('.de-close-btn');
    var resetBtn = editorEl.querySelector('.de-reset-btn');
    function refresh() {
      var cur = effDeps[key] || [];
      curBox.innerHTML = '';
      if (cur.length === 0) curBox.innerHTML = '<span class="de-empty">(no dependency)</span>';
      else cur.forEach(function(k) {
        var c = document.createElement('span');
        c.className = 'chip';
        c.innerHTML = (itemByKey[k] ? itemByKey[k].en : k) +
                       '<span class="chip-remove" title="remove">✕</span>';
        c.querySelector('.chip-remove').addEventListener('click', function() {
          effDeps[key] = (effDeps[key] || []).filter(function(x) { return x !== k; });
          saveDeps(key);
          recomputeInfluence(); renderAllRel(); refresh();
        });
        curBox.appendChild(c);
      });
      selectEl.innerHTML = '<option value="">-- pick an upstream item to add --</option>';
      ITEMS.filter(function(it) { return it.key !== key && cur.indexOf(it.key) < 0; })
           .forEach(function(it) {
             var o = document.createElement('option');
             o.value = it.key; o.textContent = it.en;
             selectEl.appendChild(o);
           });
    }
    refresh();
    addBtn.addEventListener('click', function() {
      var v = selectEl.value; if (!v) return;
      if ((effDeps[key] || []).indexOf(v) >= 0) return;
      effDeps[key] = (effDeps[key] || []).concat([v]);
      saveDeps(key); recomputeInfluence(); renderAllRel(); refresh();
    });
    resetBtn.addEventListener('click', function() {
      if (!confirm('Restore the default dependencies for this item?')) return;
      effDeps[key] = (DEFAULTS[key] || []).slice();
      saveDeps(key); recomputeInfluence(); renderAllRel(); refresh();
    });
    closeBtn.addEventListener('click', closeEditor);
  }

  // ===== Bar editor =======================================================
  var barEditorEl = null;
  function closeBarEditor() {
    if (barEditorEl && barEditorEl.parentNode) barEditorEl.parentNode.removeChild(barEditorEl);
    barEditorEl = null;
    document.querySelectorAll('.rel-edit-btn[data-kind="bars"]').forEach(function(b) { b.disabled = false; });
  }
  function openBarEditor(rowEl, btn) {
    closeEditor(); closeBarEditor();
    var key = rowEl.getAttribute('data-item');
    var meta = itemByKey[key];
    barEditorEl = document.createElement('div');
    barEditorEl.className = 'dep-editor bar-editor';
    barEditorEl.innerHTML =
      '<div class="de-title">Edit time periods of "' + meta.en + '"</div>' +
      '<div class="be-list"></div>' +
      '<div class="be-add">' +
      '  <div class="be-row">' +
      '    <select class="be-bucket"></select>' +
      '    <select class="be-start"></select>' +
      '    <span class="be-arrow">→</span>' +
      '    <select class="be-end"></select>' +
      '    <button class="de-btn be-add-btn">+ Add</button>' +
      '  </div>' +
      '</div>' +
      '<div class="de-actions">' +
      '  <button class="de-btn danger be-reset-btn">Restore defaults</button>' +
      '  <button class="de-btn secondary be-close-btn">Done</button>' +
      '</div>';
    document.body.appendChild(barEditorEl);
    btn.disabled = true;
    barEditorEl.style.width = '480px';
    var rect = btn.getBoundingClientRect();
    barEditorEl.style.top  = (window.scrollY + rect.bottom + 6) + 'px';
    var left = window.scrollX + rect.left;
    if (left + 480 > window.innerWidth - 12) left = window.innerWidth - 492;
    barEditorEl.style.left = Math.max(12, left) + 'px';

    var listBox    = barEditorEl.querySelector('.be-list');
    var bucketSel  = barEditorEl.querySelector('.be-bucket');
    var startSel   = barEditorEl.querySelector('.be-start');
    var endSel     = barEditorEl.querySelector('.be-end');
    var addBtn     = barEditorEl.querySelector('.be-add-btn');
    var closeBtn   = barEditorEl.querySelector('.be-close-btn');
    var resetBtn   = barEditorEl.querySelector('.be-reset-btn');

    function fillBucketSelect(sel, value) {
      sel.innerHTML = '';
      BUCKETS.forEach(function(b) {
        var o = document.createElement('option');
        o.value = b.key; o.textContent = b.en;
        if (value === b.key) o.selected = true;
        sel.appendChild(o);
      });
    }
    function fillSlotSelect(sel, value) {
      sel.innerHTML = '';
      TIMESLOTS.forEach(function(s) {
        var o = document.createElement('option');
        o.value = s.col; o.textContent = s.label;
        if (parseInt(value, 10) === s.col) o.selected = true;
        sel.appendChild(o);
      });
    }
    fillBucketSelect(bucketSel, 'revised');
    fillSlotSelect(startSel, '');
    fillSlotSelect(endSel, '');

    function makeId() {
      return 'user-' + Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-4);
    }
    function ensureUB() {
      if (!userBars[key]) userBars[key] = {hidden: [], added: []};
      return userBars[key];
    }
    function refreshList() {
      listBox.innerHTML = '';
      var bars = effectiveBars(key);
      if (bars.length === 0) {
        listBox.innerHTML = '<div class="de-empty">(no periods)</div>'; return;
      }
      bars.forEach(function(b) {
        var row = document.createElement('div');
        row.className = 'be-row be-existing';
        var rowBucket = document.createElement('select'); rowBucket.className = 'be-row-bucket';
        var rowStart  = document.createElement('select'); rowStart.className  = 'be-row-start';
        var rowEnd    = document.createElement('select'); rowEnd.className    = 'be-row-end';
        fillBucketSelect(rowBucket, b.bucket);
        fillSlotSelect(rowStart, b.startCol);
        fillSlotSelect(rowEnd,   b.endCol);
        var del = document.createElement('button');
        del.className = 'de-btn danger be-del-btn'; del.textContent = '✕';
        del.title = (b.source === 'xlsx' ? 'Hide this xlsx-sourced period' : 'Delete this user-added period');
        var srcTag = document.createElement('span'); srcTag.className = 'be-src-tag';
        srcTag.textContent = (b.source === 'xlsx' ? 'xlsx' : 'user');
        srcTag.classList.add(b.source === 'xlsx' ? 'tag-xlsx' : 'tag-user');

        function applyEdit() {
          var newBucket = rowBucket.value;
          var newStart  = parseInt(rowStart.value, 10);
          var newEnd    = parseInt(rowEnd.value, 10);
          if (newEnd < newStart) { var t=newStart; newStart=newEnd; newEnd=t; }
          if (newBucket === b.bucket && newStart === b.startCol && newEnd === b.endCol) return;
          var ub = ensureUB();
          if (b.source === 'xlsx') {
            if (ub.hidden.indexOf(b.id) < 0) ub.hidden.push(b.id);
            ub.added.push({id: makeId(), startCol: newStart, endCol: newEnd, bucket: newBucket});
          } else {
            ub.added = ub.added.map(function(x) {
              return x.id === b.id
                ? {id: x.id, startCol: newStart, endCol: newEnd, bucket: newBucket}
                : x;
            });
          }
          saveUserBars(key); renderRowBars(rowEl); buildMinimap(); refreshList();
        }
        rowBucket.addEventListener('change', applyEdit);
        rowStart.addEventListener('change',  applyEdit);
        rowEnd.addEventListener('change',    applyEdit);
        del.addEventListener('click', function() {
          var ub = ensureUB();
          if (b.source === 'xlsx') {
            if (ub.hidden.indexOf(b.id) < 0) ub.hidden.push(b.id);
          } else {
            ub.added = ub.added.filter(function(x) { return x.id !== b.id; });
          }
          saveUserBars(key); renderRowBars(rowEl); buildMinimap(); refreshList();
        });

        row.appendChild(srcTag);
        row.appendChild(rowBucket);
        row.appendChild(rowStart);
        var arr = document.createElement('span'); arr.className='be-arrow'; arr.textContent='→';
        row.appendChild(arr);
        row.appendChild(rowEnd);
        row.appendChild(del);
        listBox.appendChild(row);
      });
    }
    refreshList();

    addBtn.addEventListener('click', function() {
      if (!startSel.value || !endSel.value) { toast('Please pick start & end'); return; }
      var sCol = parseInt(startSel.value, 10), eCol = parseInt(endSel.value, 10);
      if (eCol < sCol) { var t=sCol; sCol=eCol; eCol=t; }
      var ub = ensureUB();
      ub.added.push({id: makeId(), startCol: sCol, endCol: eCol, bucket: bucketSel.value});
      saveUserBars(key); renderRowBars(rowEl); buildMinimap(); refreshList();
    });
    resetBtn.addEventListener('click', function() {
      if (!confirm('Restore the default time periods for this item? All user adds / edits / deletions will be discarded.')) return;
      userBars[key] = {hidden: [], added: []};
      saveUserBars(key); renderRowBars(rowEl); buildMinimap(); refreshList();
    });
    closeBtn.addEventListener('click', closeBarEditor);
  }

  document.addEventListener('mousedown', function(ev) {
    if (editorEl || barEditorEl) {
      if (ev.target.closest('.dep-editor') || ev.target.closest('.bar-editor')) return;
      if (ev.target.classList && ev.target.classList.contains('rel-edit-btn')) return;
      closeBarEditor(); closeEditor();
    }
  });

  // ===== Click delegation =================================================
  document.body.addEventListener('click', function(ev) {
    var chip = ev.target.closest('.chip[data-target]');
    if (chip) { ev.stopPropagation(); scrollToItem(chip.getAttribute('data-target')); return; }
    var editBtn = ev.target.closest('.rel-edit-btn');
    if (editBtn) {
      ev.stopPropagation();
      var rowEl = editBtn.closest('.item-row');
      var kind  = editBtn.getAttribute('data-kind');
      if (kind === 'bars') openBarEditor(rowEl, editBtn);
      else                 openEditor(rowEl, editBtn);
      return;
    }
    var bar = ev.target.closest('.bar');
    if (bar && bar.getAttribute('data-item')) {
      ev.stopPropagation(); scrollToItem(bar.getAttribute('data-item')); return;
    }
    var lbl = ev.target.closest('.item-label');
    if (lbl) {
      if (ev.target.closest('.row-note')) return;
      if (ev.target.closest('.chip')) return;
      if (ev.target.closest('.rel-edit-btn')) return;
      scrollToItem(lbl.getAttribute('data-item'));
    }
  });

  // ===== Toolbar ==========================================================
  document.getElementById('btn-scroll-today').addEventListener('click', function() {
    var target = todayLeftPx - scroller.clientWidth * 0.30;
    scroller.scrollTo({ left: Math.max(0, target), behavior: 'smooth' });
  });
  document.getElementById('btn-clear-sel').addEventListener('click', clearHighlights);

  document.getElementById('btn-export-notes').addEventListener('click', function() {
    var notes = {}, deps = {}, bars = {};
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (!k) continue;
        if (k.indexOf(LS_NOTE) === 0) notes[k.slice(LS_NOTE.length)] = localStorage.getItem(k);
        if (k.indexOf(LS_DEPS) === 0) { try { deps[k.slice(LS_DEPS.length)] = JSON.parse(localStorage.getItem(k)); } catch(e){} }
        if (k.indexOf(LS_BARS) === 0) { try { bars[k.slice(LS_BARS.length)] = JSON.parse(localStorage.getItem(k)); } catch(e){} }
      }
    } catch(e) {}
    var nCount = Object.keys(notes).length;
    var dCount = Object.keys(deps).length;
    var bCount = Object.keys(bars).length;
    if (nCount === 0 && dCount === 0 && bCount === 0) { toast('No edits yet'); return; }
    var blob = new Blob([JSON.stringify({ notes: notes, deps: deps, bars: bars }, null, 2)],
                       { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'sandwich-edits-en-' + new Date().toISOString().slice(0,10) + '.json';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    toast('Exported: notes ' + nCount + ' / deps ' + dCount + ' / periods ' + bCount);
  });

  document.getElementById('btn-reset-all').addEventListener('click', function() {
    if (!confirm('Clear all edits (notes, dependencies, time periods) saved in this browser?')) return;
    try {
      var rm = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && (k.indexOf(LS_NOTE) === 0 || k.indexOf(LS_DEPS) === 0 || k.indexOf(LS_BARS) === 0)) rm.push(k);
      }
      rm.forEach(function(k) { localStorage.removeItem(k); });
    } catch(e) {}
    document.querySelectorAll('.row-note').forEach(function(el) {
      el.textContent = el.getAttribute('data-default') || '';
      el.classList.remove('edited');
    });
    loadUserBars(); renderAllRowBars();
    loadDeps(); renderAllRel();
    buildMinimap();
    toast('All edits cleared');
  });

  // ===== Minimap ==========================================================
  var mm        = document.getElementById('minimap');
  var mmHeader  = document.getElementById('mmHeader');
  var mmCanvas  = document.getElementById('mmCanvas');
  var mmSvg     = document.getElementById('mmSvg');
  var mmViewport= document.getElementById('mmViewport');
  var mmToggleBtn = document.getElementById('mmToggleBtn');
  var mmHideBtn   = document.getElementById('mmHideBtn');
  document.getElementById('btn-show-mm').addEventListener('click', function() {
    mm.style.display = 'block';
  });
  mmHideBtn.addEventListener('click', function() { mm.style.display = 'none'; });
  mmToggleBtn.addEventListener('click', function() {
    mm.classList.toggle('collapsed');
    mmToggleBtn.textContent = mm.classList.contains('collapsed') ? '+' : '—';
  });

  function buildMinimap() {
    var rows = grid.querySelectorAll('.item-row, .section-row');
    var totalH = grid.offsetHeight;
    var canvasW = mmCanvas.clientWidth || 340;
    var canvasH = mmCanvas.clientHeight || 220;
    mmSvg.setAttribute('viewBox', '0 0 ' + trackW + ' ' + totalH);
    mmSvg.setAttribute('width',  canvasW);
    mmSvg.setAttribute('height', canvasH);
    mmSvg.setAttribute('preserveAspectRatio', 'none');
    var parts = [];
    for (var m = 0; m <= nCols; m += 4) {
      var x = m * weekW;
      parts.push('<line x1="' + x + '" y1="0" x2="' + x + '" y2="' + totalH +
                 '" stroke="#e2e8f0" stroke-width="3" />');
    }
    rows.forEach(function(row) {
      var top = row.offsetTop; var h = row.offsetHeight;
      if (row.classList.contains('section-row')) {
        parts.push('<rect x="0" y="' + top + '" width="' + trackW +
                    '" height="' + h + '" fill="#dbeafe" />');
        return;
      }
      row.querySelectorAll('.bar').forEach(function(b) {
        var left = parseFloat(b.style.left) || 0;
        var w    = parseFloat(b.style.width) || 0;
        var bg   = b.style.background || '#888';
        var bH = h - 6; var bY = top + 3;
        if (b.classList.contains('bar-original')) { bH = 4; bY = top + h - 8; }
        if (b.classList.contains('bar-update'))   { w = Math.max(w, 6); }
        parts.push('<rect x="' + left + '" y="' + bY + '" width="' + w +
                    '" height="' + bH + '" fill="' + bg + '" opacity="0.85" />');
      });
    });
    parts.push('<line x1="' + (todayLeftPx - labelW) + '" y1="0" x2="' +
               (todayLeftPx - labelW) + '" y2="' + totalH +
               '" stroke="#dc2626" stroke-width="6" />');
    mmSvg.innerHTML = parts.join('');
    updateMinimapViewport();
  }

  function updateMinimapViewport() {
    var canvasW = mmCanvas.clientWidth || 340;
    var canvasH = mmCanvas.clientHeight || 220;
    var scaleX = canvasW / trackW;
    var scaleY = canvasH / (grid.offsetHeight || 1);
    var vpLeft = Math.max(0, scroller.scrollLeft - labelW) * scaleX;
    var vpTop  = scroller.scrollTop * scaleY;
    var vpW    = (scroller.clientWidth - labelW) * scaleX;
    var vpH    = scroller.clientHeight * scaleY;
    mmViewport.style.left = vpLeft + 'px';
    mmViewport.style.top  = vpTop  + 'px';
    mmViewport.style.width  = Math.max(20, vpW) + 'px';
    mmViewport.style.height = Math.max(20, vpH) + 'px';
  }

  function scrollFromMinimap(ev) {
    var rect = mmCanvas.getBoundingClientRect();
    var x = ev.clientX - rect.left;
    var y = ev.clientY - rect.top;
    var canvasW = mmCanvas.clientWidth || 340;
    var canvasH = mmCanvas.clientHeight || 220;
    var scaleX = canvasW / trackW;
    var scaleY = canvasH / (grid.offsetHeight || 1);
    var targetLeft = x / scaleX + labelW - scroller.clientWidth / 2;
    var targetTop  = y / scaleY - scroller.clientHeight / 2;
    scroller.scrollTo({ left: Math.max(0, targetLeft),
                        top: Math.max(0, targetTop), behavior: 'auto' });
  }
  var mmDragging = false;
  mmCanvas.addEventListener('mousedown', function(ev) { mmDragging = true; scrollFromMinimap(ev); ev.preventDefault(); });
  document.addEventListener('mousemove', function(ev) { if (mmDragging) scrollFromMinimap(ev); });
  document.addEventListener('mouseup', function() { mmDragging = false; });
  scroller.addEventListener('scroll', updateMinimapViewport);
  window.addEventListener('resize', function() { buildMinimap(); });

  (function() {
    var dragX = 0, dragY = 0, startL = 0, startT = 0, dragging = false;
    mmHeader.addEventListener('mousedown', function(ev) {
      if (ev.target.tagName === 'BUTTON') return;
      dragging = true; mm.classList.add('dragging');
      dragX = ev.clientX; dragY = ev.clientY;
      var r = mm.getBoundingClientRect();
      startL = r.left; startT = r.top;
      mm.style.right = 'auto'; mm.style.bottom = 'auto';
      mm.style.left = startL + 'px'; mm.style.top = startT + 'px';
      ev.preventDefault();
    });
    document.addEventListener('mousemove', function(ev) {
      if (!dragging) return;
      var nx = startL + (ev.clientX - dragX);
      var ny = startT + (ev.clientY - dragY);
      nx = Math.max(0, Math.min(window.innerWidth - mm.offsetWidth, nx));
      ny = Math.max(0, Math.min(window.innerHeight - 40, ny));
      mm.style.left = nx + 'px'; mm.style.top = ny + 'px';
    });
    document.addEventListener('mouseup', function() {
      dragging = false; mm.classList.remove('dragging');
    });
  })();

  // ===== Init =============================================================
  snapshotXlsxBars();
  loadUserBars(); renderAllRowBars();
  loadDeps();     renderAllRel();
  restoreNotes();
  buildMinimap();
  (function() {
    var target = todayLeftPx - scroller.clientWidth * 0.30;
    scroller.scrollLeft = Math.max(0, target);
    updateMinimapViewport();
  })();
})();
""".replace("__TODAY_LEFT__", f"{today_left_px:.1f}") \
   .replace("__TRACK_W__",   str(track_w)) \
   .replace("__COL_START__", str(col_start)) \
   .replace("__N_COLS__",    str(n_cols))

    today_str = TODAY.strftime("%Y-%m-%d")
    doc = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Sandwich Continuous Hydrogenation — Project Schedule</title>
<style>{css}</style>
</head>
<body>
<div class="container">

<div class="page-head">
  <h1>Sandwich Continuous Hydrogenation Skid — Project Schedule</h1>
  <div class="sub">
    Asymchem UK Sandwich site · Pilot plant retrofit ·
    Source: <code>0509-Sandwich Continuous Hydrogen Timeline.xlsx</code> ·
    Report date: {today_str}
  </div>
</div>

<div class="panel" id="ganttPanel">
  <h2>Schedule (deliverables × weeks)</h2>
  <div class="gantt-toolbar">
    <div class="left">
      <div class="legend">{"".join(legend_html)}</div>
    </div>
    <div class="right">
      <button class="toolbar-btn secondary" id="btn-scroll-today">⤴ Jump to today</button>
      <button class="toolbar-btn secondary" id="btn-clear-sel">Clear highlights</button>
      <button class="toolbar-btn secondary" id="btn-show-mm">🗺 Mini-map</button>
      <button class="toolbar-btn" id="btn-export-notes">⇩ Export edits</button>
      <button class="toolbar-btn secondary" id="btn-reset-all">Reset all</button>
    </div>
  </div>
  <div class="toolbar-hint">
    Each row shows two bands. Top band (orange / purple / green) is the current baseline (revised plan / update node / CE involvement).
    Bottom thin light-blue band is the original (previous) baseline.
    Hover a bar for full detail. Click an item label to scroll to its first period and highlight its upstream items.
    Use the ✏️ button to edit dependencies, the 🕓 button to add / modify / delete time periods, and the yellow box to add a free-form note (all saved in your browser; use "Export edits" to share).
  </div>

  <div class="gantt">
    <div class="gantt-scroll" id="ganttScroll">
      <div class="gantt-grid" id="ganttGrid" style="min-width: calc(var(--label-w) + {n_cols*week_w}px);">
        <div class="month-row">
          <div class="header-spacer" style="top:0;"></div>
          {"".join(month_header_html)}
        </div>
        <div class="week-row">
          <div class="header-spacer" style="top:28px;"></div>
          {"".join(week_header_html)}
        </div>
        {"".join(rows_html)}
        <div class="today-marker"><div class="today-label">Today {TODAY.month}/{TODAY.day}</div></div>
      </div>
    </div>
  </div>
</div>

<!-- Draggable mini-map -->
<div class="minimap" id="minimap">
  <div class="mm-header" id="mmHeader">
    <div class="mm-title">🗺 Project overview</div>
    <div class="mm-icons">
      <button id="mmToggleBtn" title="Collapse / expand">—</button>
      <button id="mmHideBtn"  title="Hide (re-open from toolbar)">×</button>
    </div>
  </div>
  <div class="mm-body">
    <div class="mm-canvas draggable" id="mmCanvas" style="height: 220px;">
      <svg id="mmSvg" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg"></svg>
      <div class="mm-viewport" id="mmViewport"></div>
    </div>
    <div class="mm-legend">
      {"".join(minimap_legend_html)}
      <span style="margin-left:auto;color:#94a3b8;">Drag header to move · click canvas to jump</span>
    </div>
  </div>
</div>

<div id="toast" class="toast"></div>

<script id="items-meta"     type="application/json">{items_meta_json}</script>
<script id="default-deps"   type="application/json">{default_deps_json}</script>
<script id="time-slots"     type="application/json">{time_slots_json}</script>
<script id="bucket-palette" type="application/json">{palette_json}</script>

<script>{js}</script>

</div>
</body></html>
"""
    return doc


def main() -> None:
    items, c0, c1 = load_items()
    out = render_en(items, c0, c1)
    with open(HTML_PATH, "w", encoding="utf-8") as fh:
        fh.write(out)
    print(f"wrote {HTML_PATH} ({len(out)} chars, {len(items)} rows)")


if __name__ == "__main__":
    main()
