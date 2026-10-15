# Layout index

Read before choosing. Word, fit, colour and management rules live in SKILL.md. If two layouts fit, prefer the one whose needs the source fully meets.

## `title-evidence`
- **Best for:** opening a chapter with one supplied headline metric. **Never for:** a status review or several competing results.
- **Needs from the source:** reporting period, one labelled metric and agenda. **Capacity:** one hero; ≤4 agenda rows; title ≤2 lines.
- **Misuse:** putting four unrelated KPIs in the opener.

## `status-summary`
- **Best for:** management readiness across workstreams. **Never for:** method or context rows with invented status.
- **Needs from the source:** source-stated statuses, key values, risk and ask. **Capacity:** 4 workstreams; one risk; one ask.
- **Misuse:** turning an unstated status into On track.

## `figure-callouts`
- **Best for:** a dominant part image, drawing or evidence photo. **Never for:** decorative photographs or unprinted measurements.
- **Needs from the source:** matching image and 2–6 source labels with attachment locations. **Capacity:** one image ≤60% stage width; 2–6 callouts.
- **Misuse:** estimating a dimension from image pixels.

## `milestone-timeline`
- **Best for:** 4–8 dated milestones. **Never for:** percentage completion or undated workflow steps.
- **Needs from the source:** a printed date and action for each milestone. **Capacity:** 4–8 nodes; ≤4 words per action; one result block.
- **Misuse:** spacing undated steps as if they were dates.

## `peer-columns`
- **Best for:** 2–4 comparable variants or parallel needs. **Never for:** time trends or a numeric matrix needing shared rows.
- **Needs from the source:** matching comparison fields for each peer. **Capacity:** 2–4 columns; widths 858 / 564 / 417 px, inner width −64 px.
- **Misuse:** giving one peer extra unrelated fields.

## `goal-progress`
- **Best for:** current versus target. **Never for:** achievement dates or progress without a denominator.
- **Needs from the source:** current, baseline, target and units; status only if stated. **Capacity:** 4 goal rows; target tick and optional status per row.
- **Misuse:** inventing a percentage from a completion date.

## `chart-interpretation`
- **Best for:** ≥3 periods of one metric or contrasting magnitudes. **Never for:** unprinted values, close values or threshold capsules.
- **Needs from the source:** a printed value per category, units and period. **Capacity:** 3–6 categories; ≤3 interpretation groups; ≤2 column-body capsules.
- **Misuse:** putting a ceiling halfway up a bar as if it were data.

## `multi-year-table`
- **Best for:** multiple measures over comparable periods. **Never for:** one metric better shown as a trend.
- **Needs from the source:** values, units, periods, blanks and qualifications. **Capacity:** ≤5 period columns; ≤10 data rows.
- **Misuse:** mixing yearly totals with quarterly values under year headers.

## `text-kpi-rail`
- **Best for:** a short approach supported by measures. **Never for:** a long narrative or unsupported KPI claims.
- **Needs from the source:** two evidence themes and three labelled measures. **Capacity:** 2 columns; ≤2 groups each; 3 rail KPIs.
- **Misuse:** filling the rail with invented benefits.

## `site-data`
- **Best for:** 2–3 sites or vendors using the same fields. **Never for:** peers measured with different units or periods.
- **Needs from the source:** comparable site measures, units, periods and qualifiers. **Capacity:** 2–3 sites; matching compact tables.
- **Misuse:** comparing an expected benefit with a measured result.

## `executive-summary`
- **Best for:** three evidence-backed management messages. **Never for:** a chapter opening or chronological history.
- **Needs from the source:** one primary measure and supporting evidence per message. **Capacity:** 3 groups; one primary metric each.
- **Misuse:** repeating the same hero in all three groups.

## `decision-tables`
- **Best for:** options and resources leading to one ask. **Never for:** an unprioritised action backlog.
- **Needs from the source:** source options, quantities, qualifications and next gate. **Capacity:** 3 options; 3 quantity rows; one ask.
- **Misuse:** turning preliminary cost into an approved budget.

## `evidence-actions`
- **Best for:** open work with targets and ownership. **Never for:** a general method with no action rows.
- **Needs from the source:** actions, targets, owners, dates and stated statuses; preserve blanks. **Capacity:** ≤4 body rows; full-width content-sized ask below.
- **Misuse:** stretching a short table until its ask touches it.

## `bento-plain`
- **Best for:** a small set of labelled numbers on black or white. **Never for:** dense evidence tables or prose-heavy summaries.
- **Needs from the source:** supplied values and labels; matching image for with-photo-tile. **Capacity:** 4–5 native tiles; 40 px radius; 24 px gutters.
- **Misuse:** painting opaque tiles over a photo instead of using one uncovered photo tile.

## `bento-glass`
- **Best for:** a photo-backed takeaway with visibly refracted edges. **Never for:** precise evidence or a photo unrelated to the message.
- **Needs from the source:** one matching continuous photo with edges behind the tiles. **Capacity:** 4–5 tiles; 70 px bleed; 40 px radius; 24 px gutters.
- **Misuse:** using an atmospheric photo just to create glass.

## `glass-callouts`
- **Best for:** a takeaway over structural edges in a product photo. **Never for:** CAD, dimensioned figures, precise plots or placeholders.
- **Needs from the source:** matching photo and supplied values with short labels. **Capacity:** ≤2 capsules; height ≥88 px; value ≥40 px; label ≤4 words at ≥20 px/600.
- **Misuse:** placing a glass capsule over a dimensioned engineering figure.

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

Outside the dedicated glass role, no arbitrary effects/rotation, SVG/canvas, shadows, gradients, clipping or merged tables. Glass is the explicit exception: its verified runtime optics bake into individual pictures, never a whole slide. Black/white themes are per-page; native text colours and slide backgrounds come from rendered CSS. Chart-engine insets and font metrics vary by application; use the render comparison. `render_pptx.py --app powerpoint|keynote|libreoffice` exports a temporary PDF, rasterises it with PyMuPDF and removes it by default. Use `--pdf /absolute/deck/render.pdf` to retain measurement input for `check_deck.py`; remove it after checking and never deliver it. Use `check_deck.py deck.pptx pptx/ --pdf render.pdf` after rendering; `--before` accepts pre-fix PNGs to flag new empty bands as over-corrections. Application failures print their reason and fall back to LibreOffice, including its standard macOS application path.
