# HTML layouts

Word, fit, colour and management rules: SKILL.md.

| Layout | Structure | Use when |
| --- | --- | --- |
| `bento-glass` | 4–5 smoke tiles over one continuous photo, black/white page, 70 px bleed, 40 px radius and 24 px round-ended gutters | Photo-backed summary; visible refraction at 50% |
| `bento-plain` · `with-photo-tile` | One whole uncovered rounded photo tile, native text tiles beside or below it | Any glass cap fallback; never opaque tiles over a photo |
| `bento-plain` | Native rounded rectangles, 40 px radius and 24 px gutters, theme ink; #1D1D1F on black or #F5F5F7 on white | No photo; same text hierarchy without glass |
| `glass-callouts` | Capsule faces over product-photo structural edges, one native value ≥40 px and one ≤4-word label ≥20 px/600; captions outside | Product-photo takeaway; never engineering figures |
| `title-evidence` | Chapter title, reporting period, one labelled hero and agenda | Opening a chapter |
| `status-summary` | White section cards; source-stated goal/action status only, key value, short note; top risk; one ask | Slide 2, management status |
| `figure-callouts` | One contained dominant image ≤60% stage width; 2–6 external labels with percentage attachment dots and elbow leaders; optional native label/value table | Part images, drawings, photos |
| `milestone-timeline` | 4–8 dates above nodes on one rule; actions below; current/decision node accent; result or decision block below | Dated development or qualification |
| `peer-columns` | 2–4 open columns, widths 858 / 564 / 417 px; inner width = width−64; aligned header row, identity rules/dividers and evidence per column. Numbered-needs variant: number, title, evidence, Need line | Variants or parallel needs; replaces finish-comparison |
| `goal-progress` | Goal, current value at bar end, target tick/year, source-stated status pill or plain —; section/status fill on neutral track | Current versus target |
| `chart-interpretation` | Native chart with labels and gridlines; adjacent evidence. Optional `column-body-kpi` template for a chart takeaway: ≤2 glass capsules on column bodies, none repeating the headline | Magnitudes or trends; glass only for takeaway pages |
| `multi-year-table` | Year columns, separate units, right-aligned values, section/total rows, qualifier rail | Multiple measures across years |
| `text-kpi-rail` | Two short evidence columns with KPI rail | Approach and measures |
| `site-data` | Parallel site headers, measures and aligned native tables | Facility/vendor comparisons |
| `executive-summary` | Three open evidence groups, primary value with unit, secondary rows/marks | A non-opening evidence summary |
| `decision-tables` | Options table, quantity table, ask and next gate | Trade-offs and resources |
| `evidence-actions` | Action, Site, Target, Status (pill + note), Owner; full-width ask below a short table | Open work and follow-up |

Each layout has a placeholder page in `../assets/layouts/<layout>.html`. The `peer-columns` file includes four-column results and three-column numbered-needs structures. `.card` wraps each editable PowerPoint group; `.card-content` provides the inset. `.label` is uppercase 18 px/600; use `.caption` for mixed-case sublabels. Figure dots sit in the positioned image wrapper, using percentage `left`/`top` coordinates. Leaders use up to two border-only shapes per callout.

## Editable primitives

- `data-role="glass"`: per-face baked 2× JPEG, native rounded picture crop and outer shadow, native text above. Calibration and rim notes run inside the converter; see [glass.md](glass.md). Each face carries a rebuild warning in its name and each slide’s notes.
- `data-role="text"`: native text; `.headline` becomes the real title placeholder. `<br>` and pre-line newlines create paragraphs, inline styles become runs, soft wraps remain editable.
- `data-role="shape"`: native fill, rounded rectangle or border-only line. Text directly inside is a native text pill. Use `.status-pill` inside a table cell for its native pill overlay; the remaining note stays in the cell.
- `.image-needed` with `data-role="shape"`: a native editable shape and caption, transparent fill, 2 px dashed `#A1A1A6` outline, centred 20 px muted text: `Image needed: <what it should show>`. Set `--image-radius` to the slot radius (40 px for bento); otherwise it uses `--radius`. Put the caption directly inside the shape. The user can delete this shape and insert a picture in PowerPoint. List the slot under `Images needed` in the outline. Photo-based glass layouts use plain tiles/capsules until a matching photo is supplied.
- `data-role="table"`: rectangular unmerged native table, explicit column widths, content-sized rows. `tr.section`, `tr.total`, `tr.highlight` carry styling. Header-row flag is enabled; all cells wrap. Images in cells remain separate editable pictures at the cell geometry.
- `data-role="image"`: local-file or embedded picture; default `contain`. Pictures remain pictures, labels stay editable.
- `data-role="chart"`: native chart from `data-chart` JSON. Types: `bar`, `column`, `line`, `stacked_bar`, `stacked_column` (both stacked types are 100%). Fields: `categories`, `series` (`name`, `values`, `color`, optional `point_colors`), `min`, `max`; optional `axis_max` (≥15% headroom; otherwise visible-axis maximum auto-scales), `major_unit` (percentage points for 100% stacked charts; e.g. 25 converts to a .25 native tick), `gap`, `line_width`, `data_labels`, `number_format`, `label_size`, `label_color`, `category_labels`. The template builds the HTML view inside the chart from this JSON. With `category_labels: true`, native category labels and light gridlines follow edited data. Optional `plot_insets: [left,top,right,bottom]` compensates observed renderer geometry.
- `data-name` or the element's class names every shape. Card members are prefixed by the card's name. `.footer-number` becomes a native slide-number field.
- Chart value-axis ticks omit decimal places when every tick is a whole number; data labels retain their own `number_format` decimals in HTML and PowerPoint.
- `data-role="notes"`: hidden notes hold that script; do not transcribe the source.

## Conversion notes

Use absolute paths: `python3 /absolute/skill/scripts/html_to_pptx.py /absolute/deck/deck.html /absolute/deck/deck.pptx --screenshots /absolute/deck/html`.

Offline installed Chrome supplies rendered geometry. The 1920×1080 stage maps to 960×540 pt. Generated slides use Report content layouts, with black and white variants. New slides inherit the 90,107 px title (1740×82, SF Pro Display 68 px), 20 px bullet-free body, classification, footer rule and slide-number field. Block text wraps at its authored column width; pills, badges and bare numeric values remain unwrapped with measured slack. Visible unmarked text is an error. HTML chart internals are view-only and convert as a single native chart.

Outside the dedicated glass role, no arbitrary effects/rotation, SVG/canvas, shadows, gradients, clipping or merged tables. Glass is the explicit exception: its verified runtime optics bake into individual pictures, never a whole slide. Black/white themes are per-page; native text colours and slide backgrounds come from rendered CSS. Chart-engine insets and font metrics vary by application; use the render comparison. `render_pptx.py --app powerpoint|keynote|libreoffice` exports a temporary PDF, rasterises it with PyMuPDF and removes it; the output is PNG only. Application failures print their reason and fall back to LibreOffice, including its standard macOS application path.
