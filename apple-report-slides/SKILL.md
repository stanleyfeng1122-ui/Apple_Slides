---
name: apple-report-slides
description: >-
  Use when creating or rebuilding engineering management slides. Build an editable PowerPoint deck
  for engineering management (status, review, decision or progress updates) from reports, notes,
  data, PDF/PPTX/Keynote exports, slide briefs or an existing too-brief deck, with labelled evidence
  (tables, native charts, KPIs, part images, timelines) and few words, in the Apple Environmental
  Progress Report engineering format. Do not use for marketing or pitch decks, HTML-only or animated
  slides, a single chart, or small text fixes to an existing PPTX.
compatibility: macOS, SF Pro Text and Display installed, local Chrome. LibreOffice, PowerPoint or Keynote for previews.
---

# Report → editable PowerPoint

> **First deck in 5 steps**
> 1. Read the supplied source and classification.
> 2. Outline the story using the layout index.
> 3. Author HTML with the template and snippets.
> 4. Render, check and fix the HTML pages.
> 5. Convert, render in PowerPoint and check the deck.

Requirements: Python ≥3.10 with python-pptx, playwright, Pillow, NumPy and PyMuPDF; local Chrome and a preview application.
- Check availability first; never download or install automatically.
- If missing, name the blocked step and applicable install command (`python3 -m pip install python-pptx playwright Pillow numpy pymupdf`; `brew install --cask google-chrome libreoffice`), preserving authored files.
- The browser scripts default to macOS Chrome; `--chrome` accepts another local executable.
- Everything stays local and offline.

## Authoring procedure

Read [the layout index](references/layouts.md), then use [the template](assets/template.html) and matching [snippet](assets/layouts/). Author HTML directly; do not create a per-deck generator. Rules have one home: this file.

**Inputs and outputs**

- Output: `<working project>/decks/<slug>/`, with fixed names `outline.md`, `assets/`, `deck.html`, `html/`, `pptx/`, `deck.pptx`. Pass absolute paths to scripts; never write inside the skill directory.
- PDF: extract the relevant pages with `pdftotext -layout -f N -l N`, or PyMuPDF.
- PPTX: use python-pptx for text, tables, notes, native chart values and pictures; save pictures into the deck's `assets/`.
- Keynote: ask for a PPTX or PDF export.
- An existing too-brief deck is a source like any other: retain its facts and images and rebuild with these layouts.
- Footer mark: copy the source classification verbatim. If absent, ask once; never default internal material to “Public”.

1. **Read the source.** Identify scope, reporting period, units, goals, qualifications, footnotes and source pages. Use only supplied facts and printed numerical data; never digitise chart images. Separate reported results, forecasts and proposed follow-ups.
   - Preserve the printed spec window when source numbers disagree; never silently recompute it.
   - Flag the conflicting source numbers in the speaker notes.
2. **Write `outline.md`.** One line per slide: `S## · job · layout · dominant object (figure / chart / table / numbers / timeline / actions) · headline (verbatim if supplied) · evidence · source page`.
   - Choose from the index; if two fit, prefer the one whose 'needs' the source fully meets.
   - For a missing layout, use the closest, record the gap in the outline and tell the user; never edit the skill mid-deck.
   - Append ` · exception: <reason>` as the final field only for a recorded exception; prose mentioning exceptions or exemptions is not a marker.
3. **Copy the template** to `deck.html`, replace `{{TITLE}}` and insert snippet pages in `{{SLIDES}}`. Keep the CSS and viewing controls. Copy the template’s local `glass.css` and `glass.js` beside `deck.html`. No remote resources.
4. **Write each page.** Unique filename-safe ID, registered layout, source footer, exact classification, slide number and hidden notes. Mark visible primitives with `data-role`; unmarked containers arrange only. Mark panel backgrounds separately before their contents. No visible CSS pseudo-elements.
5. **Render all HTML pages:** `python3 /absolute/skill/scripts/render_html.py /absolute/deck/deck.html /absolute/deck/html`. Installed Chrome runs offline without a server.
6. **Look and fix, within the correction-round limit.** Inspect every PNG against the source and the rules below.
   - On every slide, including the first and last, check that no empty rectangle larger than one card (the empty-band limit) remains inside the checked content band.
   - Empty space inside a card counts as empty.
   - Never stretch a card or row to fill a band; use the spare-space rule: add supplied evidence, or merge.
   - Restructure or merge; never shrink below the type floors.
   - The two rounds cover HTML and PPTX repairs together.
7. **Convert:** `python3 /absolute/skill/scripts/html_to_pptx.py /absolute/deck/deck.html /absolute/deck/deck.pptx`. Text, supported charts/tables and shapes stay native; images remain pictures. Chart JSON and its HTML view use exactly the same printed values and scale.
8. **Render and compare every PPTX slide:** `python3 /absolute/skill/scripts/render_pptx.py /absolute/deck/deck.pptx /absolute/deck/pptx --app powerpoint --pdf /absolute/deck/render.pdf`. Run `python3 /absolute/skill/scripts/check_deck.py /absolute/deck/deck.pptx /absolute/deck/pptx --pdf /absolute/deck/render.pdf` after the render.
   - Fix reported failures using the remaining correction rounds, then render and check again.
   - Keep pre-fix PNGs and pass `--before /absolute/deck/before` to identify any new empty band as an over-correction.
   - The PDF is temporary measurement input; remove it after checking, never deliver it.
   - Use remaining correction rounds for observed differences. Native app export has the native-export timeout and falls back to LibreOffice with a printed reason. Every review tile names the actual renderer. Deliver PPTX and any requested review image, never a PDF companion; HTML is the design source. Rendering alone does not prove target-app edit/save/reopen acceptance.

## Density and words

1. **Content slides (title and section exempt):** the evidence floor, object count and counted-word ceiling in the limits table, no lower word target and no padding.
   - A labelled value is a number, date, grade/level, status or owner in a labelled slot (KPI with unit/caption, table cell, bar/progress label or chart point).
   - Exclude page references, slide numbers, eyebrow years and values already shown as a hero elsewhere in the deck.
   - Count tokens containing letters or digits, including headline, eyebrow, footer and text-only table cells; exclude slide numbers, chart-internal labels and numeric cells.
   - Use the sentence, block and text-cell limits below.
   - If supplied evidence cannot reach the evidence floor, merge with a neighbour.
   - A source-limited single slide with fewer labelled values than the evidence floor and no neighbour stays one slide; never pad it, and record the exception in `outline.md`.
   - Evidence spans the content width: no empty rectangle larger than the empty-band limit inside the checked content band.
2. **Headline:** use the supplied headline verbatim. Otherwise write an assertion with a verb: the result, risk or ask, within the headline limit. Topic labels belong in the eyebrow. Subtitle only for a needed unit, period or condition, within the subtitle limit; otherwise omit.
   When splitting a source slide, the first slide keeps its verbatim headline; each continuation adds the source’s own section label after it (for example, “— Part distribution”), so the headline never repeats unchanged.
   When the source has a breadcrumb such as `Project | Section` and a full sentence below it, the breadcrumb is the eyebrow and the sentence is the headline, verbatim. Never promote the breadcrumb's section words to the headline.
3. **Number:** unit and caption within the caption limit at the value; write million/billion/thousand where needed. No second explanatory line. A value is a hero (hero size below) within the hero count limits; elsewhere use a key value, row, bar or chart label.
4. **Card/column:** title within the card-title limit, at most one line within the card-line limit, or just evidence. Peer cards share the tallest member's height and grid. “Size to content” applies to the whole group.
5. **Source:** once per slide in the footer; exact page references in notes.
6. **Bottom note:** only a number-changing qualifier, within the bottom-note word, height and end limits. Move other explanations to notes; no abbreviation keys.
7. **Choose the evidence form.**
   - Turn numbered sentences into labelled values or marks.
   - At least the chart-period minimum of one metric → native column/line chart; current versus target → progress row plus target tick; at least the chart-magnitude minimum in one unit → direct-labelled bar; parts of a whole → one whole-as-percent stacked bar.
   - **Close values override the bar rule:** when the range is less than the close-value cutoff of the largest magnitude, use large numerals under identity rules.
   - Bars are for differences that are the message.
   - The KPI-line range of label/value pairs → KPI lines.
   - A table needs the minimum body rows and value columns below.
   - Chart only printed values.
8. **Never cut evidence to save words.** Preserve numbers, tables and charts; a source line fitting the source-line word limit stays visible while the slide is under the counted-word ceiling. Notes take only what does not fit, plus context, never content merely because it is prose. Every removed fact or qualifier stays in the notes as a fact.

## Fit, figures and rhythm

- Make every picture, background photo and photo tile show what its slide is about: the product, part, site, process, data or output it discusses.
- Never use a photo for atmosphere or only because it came from the source.
- Take pictures from the supplied material or from outputs the deck describes.
- Leave an `.image-needed` placeholder when no picture matches.
- List every placeholder under `Images needed` in `outline.md`, with its slide and what it should show.
- Ask the user for the missing pictures in the final reply.
- Use the stage, margins, gutters, evidence end and footer position below. SF Pro Display/Text. Use the type sizes below; table numbers regular and totals bold. Restructure if crowded.
- **Value hierarchy:** `.value` and `.formula` use SF Pro Display, the value size and weight below; `.hero` uses the hero size and weight below. `.label` sits above the value at the label size and weight below, uppercase, muted `#6E6E73`. `.caption` uses the caption size below, regular, muted. A value is never the same size and weight as its label. Use these classes in every applicable snippet.
- Use the heading/body and inter-group gap relation below. Cards have the padding below on every side; their evidence stays inside. Tables on one slide share column edges.
- A caption sits within the caption-gap limit of its evidence. Table rows fit their content: one text line uses the table-row range below; allow extra height only for wrapping or pictures, never the fixed text-row height rejected below.
- Leave the table-clearance minimum below its emitted bottom before the next element, so a cell gaining a line touches nothing. Move whole following cards together; never move footers.
- In `evidence-actions`, a table within the short-action-table row limit uses the action-table line height below, with a full-width, content-sized ask card underneath. Keep the table and ask in flow; never stretch either to fill the band. Use the spare-space rule if evidence is insufficient.
- Lists of label/value pairs, including figure evidence levels and vertical KPI lists, are native tables with wrapping cells and content-sized rows, never separate text boxes stacked at a fixed pitch. This two-column list is exempt from the numeric-table minimum of two value columns.
- Block text wraps within its authored column or card content width. Only content-sized pills, badges and bare numbers stay unwrapped; centre pills with text width + padding + the minimum pill slack below.
- Spare space: first add supplied labelled evidence (trend, baseline→target, breakdown or comparison). At thumbnail size, enlarge values and rules before anything else; never add prose. If supplied evidence remains insufficient, merge.
- Split a source slide only when it cannot fit at the type floors. If each half then leaves an empty band reaching the split/merge height below, merge them back into one slide and size rows to content.
- Let sentences wrap naturally. Hard breaks only between independent labels, never within a sentence.
- A native chart must stay correct when a manager edits its data in PowerPoint.
- Nonnegative value axes start at zero and auto-scale the maximum; use an explicit `axis_max` only with the explicit-axis headroom below.
- Show signed changes on a value axis spanning negative values, zero and positive values; a diverging bar chart is allowed.
- Preserve negative chart values as native data with an axis minimum below zero.
- Six-figure values use thousands (`#,##0,"K"`) on the axis and value labels, identically in the HTML view and the native chart; value labels are bold in both.
- Keep labels apart, reducing them no lower than the chart-label floor below.
- **Decorative art stays secondary:** an illustration or asset carrying no data occupies at most the decorative-art share below of its card or column and is never dominant. Source text, values and status come first.
- A supplied part image, drawing or photo carrying evidence is dominant. Copy image files into the deck's `assets/`, use relative paths and `contain`. Crop with Pillow if necessary, never `object-fit: cover`. Labels, grades and defects remain editable text with leaders, never baked into pixels.
- Figure callouts: use the count and attachment-dot size below; dots positioned as percentages of the image; the leader-segment count and stroke below, border-only; external labels at the word and type limits below; optional value/level row beneath. Split when the callout count limit is exceeded.
- Neither layout nor dominant object repeats on three consecutive slides. At most one coloured card per slide, reserved for the key result or ask. Peer columns use the divider and identity-rule widths below (`#D2D2D7` dividers), with one aligned header row. Grouped goals and status rows sit in white section cards with the section name.

## Management story and status

- At the management-story deck-length minimum, slide 2 is `status-summary`; evidence follows its workstream row order. The last slide repeats the ask, owners and dates. Missing owner/date remains “TBD” or the source blank; never invent one.
- Quote or closely paraphrase only source-stated asks, decisions and recommendations on slides.
- Show **For information — no decision requested** when the source supplies no ask, decision or recommendation.
- Include a deadline only when the source states it.
- Put the agent's own suggestion only in notes, starting with **Suggestion (not in source):**.
- **Source status words stay verbatim** in pills: choose the nearest semantic colour (Shipped, Done and Complete use achieved), without replacing the source word. The colour families are Achieved · On track / In progress · At risk · Blocked · Proposed. No inference from a number alone. A bare date such as “9 / 30” is a date label, never a status.
- **Status only where stated:** pills belong only to goals/actions with a source-stated status. An unstated goal/action status is a plain “—”. “TBD” is only a missing owner or date in an action row. Needs, context and method rows have no status field. Drop the status column entirely if no row has a stated status.
- Never show a standalone “Status: —” line. “—” belongs only inside a status cell or column. If there is no status, show nothing.

## Colour roles

- **Identity:** preserve source category colours on category marks: identity-rule-width column/card top rules, bars/series and swatches.
- For PPTX sources, read theme accents and the most-used non-grey title, rule and pill colours; for PDF/image sources, sample title and rule colours.
- Carry them as `--accent` and identity tokens and name them in `outline.md`.
- For text-only input without colour, use the default accent and say so in the outline.
- Define source identity colours as `--identity-1`, `--identity-2`, etc. only in the deck's `<style>`, never in this package.
- Eyebrow: `Project | Section`, with `--section-text` for the section label.
- **Status:** the closed vocabulary below colours only status pills and progress fills. Other values, correlations and timeline nodes never borrow status colour semantics. Progress may instead use the section/series colour while its pill carries the status.
- **Accent:** one per deck, `--accent` = source accent, else `#0071E3`. Within the accent-use limit per slide, for the connective label, current/decision timeline node, latest period in a single-series chart, or key result. Use dark text where the accent itself has insufficient contrast.
- **Connective label:** one optional `.connective` accent label, within the connective word limit, linking evidence groups when supplied by the source. It counts toward the word ceiling.
- Never use status colours on non-status values, or make an all-grey chart/goal page when section or series colour exists. The converter puts the deck accent and identity colours in PowerPoint theme `accent1`–`accent6`.

| Status | Fill token | Fill | Text |
| --- | --- | --- | --- |
| Achieved | `--status-achieved` | `#00CC2D` | `#006B1B` |
| On track / In progress | `--status-progress` | `#8E8E93` | Ink |
| At risk | `--status-risk` | `#FF9F0A` | `#8A5200` |
| Blocked | `--status-blocked` | `#FF3B30` | `#B3261E` |
| Proposed | `--status-open` | `#D2D2D7` | Ink |

Use status fill only on progress fills and light pills with the dark text token, or Ink text on a light neutral pill.
- Use the text contrast ratios and large-type cutoff below; never use bright status fills as small text.
- Progress tracks are neutral with a target tick at the stroke below and a value label anchored at the bar end; add a status pill only when stated.
- Achievement dates use a dated timeline, not a percentage bar.
- Preserve units, periods, blanks, `/`, approximations, preliminary totals and source qualifications.

## Optional Liquid Glass and bento

Use only when a straight backdrop edge crosses a cap or rim and visibly bends at the review scale below in the PPTX render. ΔL* is printed as information only, never as a threshold or the visual decision. Full rules and verified values: [glass.md](references/glass.md).

| Backdrop | Variant | Native text | Tint |
| --- | --- | --- | --- |
| Product photo / dark or mid content | dark (smoke) | white | smoke tint limits below |
| Light column body | light (clear) | #1D1D1F | clear tint and milk below |
| Saturated accent column body | light dim | white | dim tint below |
| No photo | bento-plain | theme ink | native solid fill; no glass |
| Any CAD, dimensioned or engineering figure | plain callouts | white | native #1D1D1F pill; no glass |

- Set `data-theme="black"` or `"white"` per page; report remains the default. Tone follows the content beneath each box. Copy `glass.js` and `glass.css` beside the deck when resolving the template's local asset paths.
- Build photo-backed glass only over a matching photo. With an image placeholder, use plain `bento-plain` tiles or plain capsules, keeping the placeholder caption visible. Rebuild with glass when the user supplies the matching photo. Chart-body capsules keep their chart-specific rules.
- Use the glass-slide range per section below, never an evidence slide. A chart takeaway may have at most the chart-capsule count below on column bodies, neither repeating the headline number; protect bar tops, axes, labels and limits. No glass on tables, precise plots, defect images or small text.
- A capsule or label on a chart never carries a threshold or target. State a threshold beside the chart as a labelled value.
- At least the glass coverage below of every face sits over content; capsule caps may overhang the page. Bento tiles are fully backed by one continuous photo, the lens bleed, radius and gutter below with round ends; text padding at least the radius. No photo means plain bento (#1D1D1F on black / #F5F5F7 on white), with the same text rules.
- Use `bento-glass`, `bento-plain` or `glass-callouts`; the chart snippet includes an optional column-body capsule. Glass callouts are capsules at least the glass-callout height below; any engineering figure uses plain native pills at the engineering-callout height below. These optional layouts override ordinary card fill/radius defaults only.
- Each glass tile or capsule carries one value and one label at the glass-text limits below. Eyebrows, captions, unit lines and sources belong outside the glass in a plain caption row or beside the capsule.
- The converter calibrates to the glass text contrast ratios below. Before the smoke tint cap forces fallback, it searches the text-zone count and photo-offset step listed in the limits table; it prints each chosen placement. Missing lens bleed is mirror-padded.
- If every placement exceeds the smoke tint cap, use one whole uncovered rounded photo tile with `bento-plain` text tiles beside or below it. Never paint opaque tiles over a photo or leave photo slivers in gutters.
- The converter prints rim ΔL*, then bakes only each face at the bake scale and JPEG quality below. Text, rounded crop and outer shadow stay native. Caption frames keep the capsule's inner width and wrap on; lengthening text requires a rebuild because the baked face cannot grow with it.
- Every glass slide's notes end with: “Glass is a baked picture: rebuild the deck after moving it, editing what is under it, or lengthening its text.”

## Numeric limits

All dimensions are stage pixels unless a ratio, count or unit is given.

<table>
<tr><th>Rule</th><th>Limit</th></tr>
<tr><td>Stage and footer</td><td>1920×1080; side margins 90 px; gutters 24 px; evidence ends y=980; footer rule y=1004</td></tr>
<tr><td>Evidence and words</td><td>≥10 labelled values in 2–4 objects; ≤90 counted words; no word floor</td></tr>
<tr><td>Prose</td><td>sentences total ≤25 words; block/subtitle/card line/source line/bottom note ≤10; text cell/caption ≤6; headline/connective ≤8; card title ≤4</td></tr>

<tr><td>Notes and repairs</td><td>notes ≤150 words; ≤2 correction rounds across HTML and PPTX; native export 120 s</td></tr>
<tr><td>Empty space</td><td>no rectangle exceeding 560×300 px between y=200 and y=980; split halves with empty bands ≥200 px tall merge</td></tr>
<tr><td>Body, tables, headlines</td><td>body 18–22 px; table numbers 18–22 px; headline 64–68 px</td></tr>
<tr><td>Values and heroes</td><td>value/formula 32–44 px, weight 600–700; hero 72–96 px, weight 600–700; each hero once per deck, ≤1 per slide</td></tr>

<tr><td>Labels and captions</td><td>label 16–18 px, weight 600; caption 18–20 px; chart labels ≥14 px</td></tr>
<tr><td>Group spacing</td><td>heading→body ≤half inter-group gap, defaults 8 / 48 px; card padding 32 px; caption within 24 px</td></tr>
<tr><td>Tables</td><td>≥2 body rows and ≥2 value columns; text row 56–64 px; never fixed 96 px text rows; clearance ≥body line pitch +8 px and ≥28 px</td></tr>
<tr><td>Action tables and pills</td><td>≤4 body rows: 1.5× line height; pill slack ≥max(6 px, 8% of text width)</td></tr>

<tr><td>Bottom note</td><td>one 24 px-high line ending by y=980</td></tr>
<tr><td>Evidence forms</td><td>≥3 periods → column/line; ≥3 magnitudes → bar; whole → 100% stacked bar; 1–2 label/value pairs → KPI lines</td></tr>
<tr><td>Chart differences and scale</td><td>close-value range &lt;15% of largest magnitude → numerals; explicit axis maximum ≥15% headroom</td></tr>
<tr><td>Figures</td><td>2–6 callouts; 12 px dots; ≤2 leader segments at 2 px; labels ≤6 words at 20–24 px; split above six</td></tr>

<tr><td>Decorative art</td><td>≤25% of its card/column</td></tr>
<tr><td>Chrome</td><td>divider 1 px; identity top rule 6 px; target tick 2 px; accent at most two uses per slide</td></tr>
<tr><td>Story</td><td>≥3 slides: slide 2 status summary; no layout or dominant object on three consecutive slides</td></tr>
<tr><td>Text contrast</td><td>≥4.5:1; ≥3:1 at ≥40 px</td></tr>

<tr><td>Glass frequency and review</td><td>1–2 slides per section; ≤2 column-body capsules; visible bend at 50%</td></tr>
<tr><td>Glass geometry</td><td>coverage ≥70%; lens bleed 70 px; radii 40 px; gutters 24 px; padding ≥radius; glass capsule ≥88 px tall; engineering pill 48 px</td></tr>
<tr><td>Glass text</td><td>one value ≥40 px; one label ≤4 words at ≥20 px/600</td></tr>
<tr><td>Glass optics</td><td>smoke .14 floor / ≤.25 cap; clear .06 with milk 0 on white; dim .06–.07</td></tr>

<tr><td>Glass calibration</td><td>5.0:1 below 40 px; 3.5:1 at ≥40 px; five text zones; offsets in 40 px steps; bake 2× JPEG q93</td></tr>
</table>

## Presenter notes

Within the notes word limit per slide: the point, context each number needs, qualifications and source pages.
- Do not repeat visible text or transcribe the source.
- Never include earlier wording, drafts or revision history.
- Keep confidential briefs, sources and slide content outside this package.
- No skill file may cite repo paths, run IDs or personal paths.
- Use the placeholder `{{PRESENTER SCRIPT ≤150 WORDS: POINT · CONTEXT · QUALIFIERS · SOURCE PAGES}}`.
