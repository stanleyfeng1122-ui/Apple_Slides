---
name: apple-report-slides
description: Use when creating or rebuilding engineering management slides. Build an editable PowerPoint deck for engineering management (status, review, decision or progress updates) from reports, notes, data, PDF/PPTX/Keynote exports, slide briefs or an existing too-brief deck, with labelled evidence (tables, native charts, KPIs, part images, timelines) and few words, in the Apple Environmental Progress Report engineering format. Do not use for marketing or pitch decks, HTML-only or animated slides, a single chart, or small text fixes to an existing PPTX.
compatibility: macOS, SF Pro Text and Display installed, local Chrome. LibreOffice, PowerPoint or Keynote for previews.
---

# Report → editable PowerPoint

Requirements: Python ≥3.10 with python-pptx, playwright, Pillow, NumPy and PyMuPDF; local Chrome and a preview application. Check availability first; never download or install automatically. If missing, name the blocked step and applicable install command (`python3 -m pip install python-pptx playwright Pillow numpy pymupdf`; `brew install --cask google-chrome libreoffice`), preserving authored files. The browser scripts default to macOS Chrome; `--chrome` accepts another local executable. Everything stays local and offline.

## Authoring procedure

Read [the layout registry](references/layouts.md), then use [the template](assets/template.html) and matching [snippet](assets/layouts/). Author HTML directly; do not create a per-deck generator. Rules have one home: this file.

**Inputs and outputs**

- Output: `<working project>/decks/<slug>/`, with fixed names `outline.md`, `assets/`, `deck.html`, `html/`, `pptx/`, `deck.pptx`. Pass absolute paths to scripts; never write inside the skill directory.
- PDF: extract the relevant pages with `pdftotext -layout -f N -l N`, or PyMuPDF.
- PPTX: use python-pptx for text, tables, notes, native chart values and pictures; save pictures into the deck's `assets/`.
- Keynote: ask for a PPTX or PDF export.
- An existing too-brief deck is a source like any other: retain its facts and images and rebuild with these layouts.
- Footer mark: copy the source classification verbatim. If absent, ask once; never default internal material to “Public”.

1. **Read the source.** Identify scope, reporting period, units, goals, qualifications, footnotes and source pages. Use only supplied facts and printed numerical data; never digitise chart images. Separate reported results, forecasts and proposed follow-ups.
2. **Write `outline.md`.** One line per slide: `job · layout · dominant object (figure / chart / table / numbers / timeline / actions) · headline (verbatim if supplied) · evidence · source page`. Choose a registered layout. If one is missing, use the closest, record the gap in the outline and tell the user; never edit the skill mid-deck.
3. **Copy the template** to `deck.html`, replace `{{TITLE}}` and insert snippet pages in `{{SLIDES}}`. Keep the CSS and viewing controls. Copy the template’s local `glass.css` and `glass.js` beside `deck.html`. No remote resources.
4. **Write each page.** Unique filename-safe ID, registered layout, source footer, exact classification, slide number and hidden notes. Mark visible primitives with `data-role`; unmarked containers arrange only. Mark panel backgrounds separately before their contents. No visible CSS pseudo-elements.
5. **Render all HTML pages:** `python3 /absolute/skill/scripts/render_html.py /absolute/deck/deck.html /absolute/deck/html`. Installed Chrome runs offline without a server.
6. **Look and fix, at most two rounds.** Inspect every PNG against the source and the rules below. On every slide, including the first and last, check that no empty rectangle larger than one card (about 560×300 px) remains between y=200 and y=980. Empty space inside a card counts as empty. Never stretch a card or row to fill a band; use the spare-space rule: add supplied evidence, or merge. Restructure or merge; never shrink below the type floors. The two rounds cover HTML and PPTX repairs together.
7. **Convert:** `python3 /absolute/skill/scripts/html_to_pptx.py /absolute/deck/deck.html /absolute/deck/deck.pptx`. Text, supported charts/tables and shapes stay native; images remain pictures. Chart JSON and its HTML view use exactly the same printed values and scale.
8. **Render and compare every PPTX slide:** `python3 /absolute/skill/scripts/render_pptx.py /absolute/deck/deck.pptx /absolute/deck/pptx --app powerpoint`. Use remaining correction rounds for observed differences. Native app export has a 120 s limit and falls back to LibreOffice with a printed reason. Every review tile names the actual renderer. Deliver PPTX and any requested review image, never a PDF companion; HTML is the design source. Rendering alone does not prove target-app edit/save/reopen acceptance.

## Density and words

1. **Content slides (title and section exempt):** at least **10 labelled values in 2–4 evidence objects**, at most **90 counted words**, no lower word target and no padding. A labelled value is a number, date, grade/level, status or owner in a labelled slot (KPI with unit/caption, table cell, bar/progress label or chart point). Exclude page references, slide numbers, eyebrow years and values already shown as a hero elsewhere in the deck. Count tokens containing letters or digits, including headline, eyebrow, footer and text-only table cells; exclude slide numbers, chart-internal labels and numeric cells. Sentences total ≤25 words, text blocks ≤10, text cells ≤6. If supplied evidence cannot reach 10, merge with a neighbour. A source-limited single slide with fewer than 10 labelled values and no neighbour stays one slide; never pad it, and record the exception in `outline.md`. Evidence spans the content width: no empty rectangle larger than about 560×300 px between y=200 and y=980.
2. **Headline:** use the supplied headline verbatim. Otherwise write an assertion with a verb: the result, risk or ask, ≤8 words. Topic labels belong in the eyebrow. Subtitle only for a needed unit, period or condition, ≤10 words; otherwise omit.
   When splitting a source slide, the first slide keeps its verbatim headline; each continuation adds the source’s own section label after it (for example, “— Part distribution”), so the headline never repeats unchanged.
   When the source has a breadcrumb such as `Project | Section` and a full sentence below it, the breadcrumb is the eyebrow and the sentence is the headline, verbatim. Never promote the breadcrumb's section words to the headline.
3. **Number:** unit and caption ≤6 words at the value; write million/billion/thousand where needed. No second explanatory line. A value is a hero (72–96 px) once per deck, at most one per slide; elsewhere use a key value, row, bar or chart label.
4. **Card/column:** title ≤4 words, at most one ≤10-word line, or just evidence. Peer cards share the tallest member's height and grid. “Size to content” applies to the whole group.
5. **Source:** once per slide in the footer; exact page references in notes.
6. **Bottom note:** only a number-changing qualifier, ≤10 words in one 24 px-high line ending by y=980. Move other explanations to notes; no abbreviation keys.
7. Turn numbered sentences into labelled values or marks. ≥3 periods of one metric → native column/line chart; current versus target → progress row plus target tick; ≥3 magnitudes in one unit → direct-labelled bar; parts of a whole → one 100% stacked bar. **Close values override the bar rule:** when the range is less than 15% of the largest magnitude, use large numerals under identity rules. Bars are for differences that are the message. 1–2 label/value pairs → KPI lines. A table needs ≥2 body rows and ≥2 value columns. Chart only printed values.
8. **Never cut evidence to save words.** Preserve numbers, tables and charts; a fitting source line of ≤10 words stays visible while the slide is under the 90-word ceiling. Notes take only what does not fit, plus context, never content merely because it is prose. Every removed fact or qualifier stays in the notes as a fact.

## Fit, figures and rhythm

- 1920×1080 stage, 90 px side margins, 24 px gutters; evidence ends by y=980, footer rule y=1004. SF Pro Display/Text. Body 18–22 px; table numbers 18–22 px regular, totals bold; headlines 64–68 px. Restructure if crowded.
- **Value hierarchy:** `.value` and `.formula` use SF Pro Display, 32–44 px, weight 600–700; `.hero` uses 72–96 px, weight 600–700. `.label` sits above the value at 16–18 px, weight 600, uppercase, muted `#6E6E73`. `.caption` is 18–20 px, regular, muted. A value is never the same size and weight as its label. Use these classes in every applicable snippet.
- Heading→body gap ≤half inter-group gap (defaults 8 / 48 px). Cards have 32 px padding on every side; their evidence stays inside. Tables on one slide share column edges.
- A caption sits within 24 px of its evidence. Table rows fit their content: one text line is 56–64 px; allow extra height only for wrapping or pictures, never fixed 96 px text rows.
- Leave at least the table's body line pitch + 8 px (and at least 28 px) below its emitted bottom before the next element, so a cell gaining a line touches nothing. Move whole following cards together; never move footers.
- In `evidence-actions`, a table of ≤4 body rows uses 1.5× line height, with a full-width, content-sized ask card underneath. Keep the table and ask in flow; never stretch either to fill the band. Use the spare-space rule if evidence is insufficient.
- Lists of label/value pairs, including figure evidence levels and vertical KPI lists, are native tables with wrapping cells and content-sized rows, never separate text boxes stacked at a fixed pitch. This two-column list is exempt from the numeric-table minimum of two value columns.
- Block text wraps within its authored column or card content width. Only content-sized pills, badges and bare numbers stay unwrapped; centre pills with text width + padding + at least max(6 px, 8% of text width) of slack.
- Spare space: first add supplied labelled evidence (trend, baseline→target, breakdown or comparison). At thumbnail size, enlarge values and rules before anything else; never add prose. If supplied evidence remains insufficient, merge.
- Split a source slide only when it cannot fit at the type floors. If each half then leaves an empty band ≥200 px tall, merge them back into one slide and size rows to content.
- Let sentences wrap naturally. Hard breaks only between independent labels, never within a sentence.
- A native chart must stay correct when a manager edits its data in PowerPoint. Visible axes start at zero and auto-scale the maximum; use an explicit `axis_max` only with at least 15% headroom. Six-figure values use thousands (`#,##0,"K"`) on the axis and value labels, identically in the HTML view and the native chart; value labels are bold in both. Keep labels apart, reducing them no lower than 14 px.
- **Decorative art stays secondary:** an illustration or asset carrying no data occupies at most 25% of its card or column and is never dominant. Source text, values and status come first.
- A supplied part image, drawing or photo carrying evidence is dominant. Copy image files into the deck's `assets/`, use relative paths and `contain`. Crop with Pillow if necessary, never `object-fit: cover`. Labels, grades and defects remain editable text with leaders, never baked into pixels.
- Figure callouts: 2–6; 12 px attachment dots positioned as percentages of the image; ≤2 border-only 2 px leader segments; external ≤6-word labels at 20–24 px; optional value/level row beneath. Split if more than six.
- Neither layout nor dominant object repeats on three consecutive slides. At most one coloured card per slide, reserved for the key result or ask. Peer columns use 1 px `#D2D2D7` vertical dividers or 6 px identity top rules, with one aligned header row. Grouped goals and status rows sit in white section cards with the section name.

## Management story and status

- For ≥3 slides, slide 2 is `status-summary`; evidence follows its workstream row order. The last slide repeats the ask, owners and dates. Missing owner/date remains “TBD” or the source blank; never invent one.
- Ask = verb + decision + by-when + what it unblocks. If no decision is supplied: **For information — no decision requested**. “Assign owners” or “confirm status” is never the ask. Label proposed asks and historical status.
- **Source status words stay verbatim** in pills: choose the nearest semantic colour (Shipped, Done and Complete use achieved), without replacing the source word. The colour families are Achieved · On track / In progress · At risk · Blocked · Proposed. No inference from a number alone. A bare date such as “9 / 30” is a date label, never a status.
- **Status only where stated:** pills belong only to goals/actions with a source-stated status. An unstated goal/action status is a plain “—”. “TBD” is only a missing owner or date in an action row. Needs, context and method rows have no status field. Drop the status column entirely if no row has a stated status.
- Never show a standalone “Status: —” line. “—” belongs only inside a status cell or column. If there is no status, show nothing.

## Colour roles

- **Identity:** preserve source category colours on category marks: 6 px column/card top rules, bars/series and swatches. For PPTX sources, read theme accents and the most-used non-grey title, rule and pill colours; for PDF/image sources, sample title and rule colours. Carry them as `--accent` and identity tokens and name them in `outline.md`. For text-only input without colour, use the default accent and say so in the outline. Define source identity colours as `--identity-1`, `--identity-2`, etc. only in the deck's `<style>`, never in this package. Eyebrow: `Project | Section`, with `--section-text` for the section label.
- **Status:** the closed vocabulary below colours only status pills and progress fills. Other values, correlations and timeline nodes never borrow status colour semantics. Progress may instead use the section/series colour while its pill carries the status.
- **Accent:** one per deck, `--accent` = source accent, else `#0071E3`. At most two uses per slide, for the connective label, current/decision timeline node, latest period in a single-series chart, or key result. Use dark text where the accent itself has insufficient contrast.
- **Connective label:** one optional `.connective` accent label, ≤8 words, linking evidence groups when supplied by the source. It counts toward the word ceiling.
- Never use status colours on non-status values, or make an all-grey chart/goal page when section or series colour exists. The converter puts the deck accent and identity colours in PowerPoint theme `accent1`–`accent6`.

| Status | Fill token | Fill | Text |
| --- | --- | --- | --- |
| Achieved | `--status-achieved` | `#00CC2D` | `#006B1B` |
| On track / In progress | `--status-progress` | `#8E8E93` | Ink |
| At risk | `--status-risk` | `#FF9F0A` | `#8A5200` |
| Blocked | `--status-blocked` | `#FF3B30` | `#B3261E` |
| Proposed | `--status-open` | `#D2D2D7` | Ink |

Use status fill only on progress fills and light pills with the dark text token, or Ink text on a light neutral pill. Text contrast ≥4.5:1 (≥3:1 at ≥40 px); never use bright status fills as small text. Progress tracks are neutral with a 2 px target tick and a value label anchored at the bar end; add a status pill only when stated. Achievement dates use a dated timeline, not a percentage bar. Preserve units, periods, blanks, `/`, approximations, preliminary totals and source qualifications.

## Optional Liquid Glass and bento

Use only when a straight backdrop edge crosses a cap or rim and visibly bends at 50% in the PPTX render. ΔL* is printed as information only, never as a threshold or the visual decision. Full rules and verified values: [glass.md](references/glass.md).

| Backdrop | Variant | Native text | Tint |
| --- | --- | --- | --- |
| Product photo / dark or mid content | dark (smoke) | white | .14 floor, calibrated to ≤.25 |
| Light column body | light (clear) | #1D1D1F | .06, milk 0 on white |
| Saturated accent column body | light dim | white | .06–.07 |
| No photo | bento-plain | theme ink | native solid fill; no glass |
| Any CAD, dimensioned or engineering figure | plain callouts | white | native #1D1D1F pill; no glass |

- Set `data-theme="black"` or `"white"` per page; report remains the default. Tone follows the content beneath each box. Copy `glass.js` and `glass.css` beside the deck when resolving the template's local asset paths.
- Use 1–2 glass slides per section, never an evidence slide. A chart takeaway may have at most two capsules on column bodies, neither repeating the headline number; protect bar tops, axes, labels and limits. No glass on tables, precise plots, defect images or small text.
- At least 70% of every glass face sits over content; capsule caps may overhang the page. Bento tiles are fully backed by one continuous photo, 70 px lens bleed, 40 px radii and 24 px gutters with round ends; text padding ≥radius. No photo means plain bento (#1D1D1F on black / #F5F5F7 on white), with the same text rules.
- Use `bento-glass`, `bento-plain` or `glass-callouts`; the chart snippet includes an optional column-body capsule. Glass callouts are capsules ≥88 px tall; any engineering figure uses plain 48 px native pills. These optional layouts override ordinary card fill/radius defaults only.
- Each glass tile or capsule carries one value ≥40 px and one label of ≤4 words at ≥20 px/600. Eyebrows, captions, unit lines and sources belong outside the glass in a plain caption row or beside the capsule.
- The converter calibrates to 5.0:1 below 40 px and 3.5:1 at ≥40 px. Before the .25 tint cap forces fallback, it searches five text zones and photo offsets in 40 px steps; it prints each chosen placement. Missing lens bleed is mirror-padded.
- If every placement exceeds .25, use one whole uncovered rounded photo tile with `bento-plain` text tiles beside or below it. Never paint opaque tiles over a photo or leave photo slivers in gutters.
- The converter prints rim ΔL*, then bakes only each face at 2× JPEG q93. Text, rounded crop and outer shadow stay native. Caption frames keep the capsule's inner width and wrap on; lengthening text requires a rebuild because the baked face cannot grow with it.
- Every glass slide's notes end with: “Glass is a baked picture: rebuild the deck after moving it, editing what is under it, or lengthening its text.”

## Presenter notes

≤150 words per slide: the point, context each number needs, qualifications and source pages. Do not repeat visible text or transcribe the source. Never include earlier wording, drafts or revision history. Keep confidential briefs, sources and slide content outside this package. No skill file may cite repo paths, run IDs or personal paths. Use the placeholder `{{PRESENTER SCRIPT ≤150 WORDS: POINT · CONTEXT · QUALIFIERS · SOURCE PAGES}}`.
