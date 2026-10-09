# Liquid Glass in the slide skill: rules

Optional material, ported from the verified round-5 prototype. The full build is integrated into `scripts/html_to_pptx.py`; `scripts/glass.py` handles calibration, rim measurement and baking. The optical runtime is `assets/glass.js` with `assets/glass.css`.

**What counts as Liquid Glass.** The pane bends the content behind it at the rim, shows that content through a nearly clear face, and has a thin rim with one bright arc. The test is to look at the PowerPoint render at 50% and point to one backdrop feature that bends or steps at a rim. If you can't point to one, the pane is acrylic or a frosted card: drop the glass.

## 1. When to use it, and when never to

**Photo fit comes first.** Photo-backed glass is built only over a picture that shows the slide's subject, from supplied material or the output discussed. Never choose an atmospheric photo to make glass possible. If none matches, keep a visible `.image-needed` shape and use plain `bento-plain` tiles or plain capsules beside it. Keep the caption unobstructed, list the image in `outline.md`, and ask the user for it. Rebuild with glass after the matching photo arrives. Chart-body glass remains subject to §4.

**Use glass only where it has something to bend.** A high-contrast structure has to cross the rim. Examples: a photo, a product or part render, or the body of a chart bar.
- **Information only:** the converter prints ΔL* as p95−p5 of CIE L* in each sampled rim band, without a threshold warning. A straight structural edge must cross a cap or rim; random texture can pass this number and still fail the 50% visual test.
- Judged glass capsules also worked at ΔL* 20. Contrast spread does not establish refraction; inspect the structural edge at 50%.
- Thin lines, such as a CAD outline, cover under 5% of the band. They don't count toward the gate, and at 50% they don't carry the lens either.

**Never use glass over any of the following:**
- **Flat #000 or #FFF.** Over black it reads as acrylic. Over white it is pixel-identical to an opaque white card (round-4 C4).
- **Random texture with no edge,** such as metal powder. Put plain native text there instead (C2 "100%").
- **Evidence the audience must read precisely:**
  - tables;
  - body text of 24 px or less;
  - SPC and metrology plots;
  - scatter plots;
  - line charts read against gridlines;
  - defect images;
  - dimensioned drawings, or any CAD where position is the data.
- **Data ends:** bar tops, line points, value labels, axis ticks, spec limits and defect markers. Keep each one at least pull + 8 px (about 28 px) outside every rim. Glass may sit on a bar body, because bar width is not data.
- **Every CAD, dimensioned or engineering figure, whatever its ΔL*.** Use plain native callouts (§7); the figure remains unaltered evidence.

**Where it belongs:** glass frames the takeaway and never sits on the proof. Use it on hero, KPI, summary and photo slides. Evidence slides stay plain.

## 2. Black page vs white page: tone follows the content

Tone follows the content under the box, not the page.

| Content under the glass | Variant | Text | Tint |
|---|---|---|---|
| Photo, or dark/mid content | **dark (smoke)** | white; labels #f2f2f7; soft halo | calibrated; .14 floor, **.25 cap** (a darker face is a slab) |
| Bright photo zone | **light (clear)** | #1d1d1f | existing light optics; calibrated to **≤.25** |
| Light figure, bar body | **light (clear)** | #1d1d1f | .06 over bars; milk .04 over light figures on black; **milk 0 on a white page** |
| Saturated accent fill (e.g. accent bar) | **light dim** | white | black tint, about .06–.07 |

**Black page**
- A shadow is invisible on black, so the lift comes from the rim, the inner bevel and the backdrop seen through the face.
- On photo tiles, smoke of .14–.20 is enough.

**White page**
- At least 70% of every glass face sits over content; capsule caps may overhang the page.
- For a bento on white, back every tile 100% with the photo and use smoke tiles with white text (C4). The white page shows around the bento and in no tile.
- The depth cue is a soft, wide shadow: 0 16 px 48 px, 13–18% black.
- **Light glass over white content** (C1 caps): the dark outer hairline and the lower-right arc carry the edge.
  - Keep the verified C1 values: lower-right dark arc .45, outer hairline .14.

## 3. Shapes and sizes

**Callouts: capsules only**
- Radius = h/2 and h ≥ 88 px. Content is value-only or at most 2 lines.
- Place them so a straight structural edge crosses a cap.
- On product photos, at least 70% of the face must sit over the object.

**Bento**
- At most 5 tiles, with one shared 40 px radius and 24 px gutters.
- One continuous photo backs 100% of every tile.
- The lens samples a bleed of 70 px past the bento, so outer rims never pull in the page colour. A rim samples at most pull + 2·fringe + blur (about 30 px) outside its edge; 70 px leaves margin.
- The gutters show the unrefracted photo. The step between gutter and tile is what proves the lens.
- Clip the page photo to the tile outline:
  - no photo "ears" in the outer notches;
  - no flat cuts where a gutter meets the outer edge. Each gutter stub ends in a 12 px round end whose tip sits on the outer corners' tangent line.
- A 12 px "tray" around the whole bento was tried and rejected: it draws an unrefracted photo lip over every outer rim (the 12 px lip covers the outer refraction band).

**Text inside tiles**
- Keep text padding at least equal to the radius: 44 px on the big tile, 40 px on small tiles, from every rim.
- Anchor the text where the photo is darkest. In C4 that is the bottom of each tile, over the solar canopies.
- Before falling back, try top-left, bottom-left, bottom-right, top-right and centre-left text zones, plus photo offsets in 40 px steps within available bleed. Choose the combination needing the lowest tint and print the choice. The tint cap stays .25.
- On photos, also test the existing light material when the selected zone is bright; tone follows the content beneath the box. Preserve its verified rim and lens settings. Print the chosen tone with the placement.

**The photo must be sharp**
- Its native pixels must be at least the displayed size: at most 1.0× upscale at 1× slide scale. The 2× bake may upscale.
- Take photos from the PDF's embedded image stream (pymupdf `Pixmap(doc, xref)`), not from a page render:
  - a render at any dpi only resamples the same pixels;
  - it also adds the page's type over the photo.
- The p. 23 photo is only 919 × 508 px, so no dpi could make it sharp. That is why C4 changed photo and story.

## 4. Charts

- At most two glass KPI capsules per chart, and never one that repeats the headline number. Put them only on a chart takeaway slide; evidence slides stay plain.
- Put capsules on bar bodies, at least 60 px below the bar top and its value label.
- Bars stay native shapes, as columns from zero.
- Light glass over bars:
  - tint ≤ .06 and saturate ≤ 1.15;
  - the bar colour shift under glass stays ≤ 5% (C1: (106,167,236) → (107,173,249)).
- On the accent bar, use the dim variant with white text.
- Never cover a data end, marker, axis, trend line or error bar.

## 5. Callouts on product photos

- A capsule sits at least 70% over content, positioned over the object, with a clean structural edge crossing a cap.
- Over fine high-frequency texture (watch mesh, foliage, gravel) the channel fringe is 0.6; the 1.0 default sparkles orange/blue along the rim. `scripts/glass.py` measures the rim band's fine detail and sets 0.6 automatically above 12, printing the value per box. An authored `data-fringe` always wins.
- Captions, unit lines and sources go beside the capsule on the plain page, never inside its glass face.

## 6. Text on glass

- Each tile or capsule carries one value ≥40 px and one label of ≤4 words at ≥20 px, weight 600.
- Eyebrows, captions, unit lines and sources sit outside glass, in a plain caption row under the bento or beside the capsule.
- All foreground text stays native; it is never refracted or baked into the face.
- `scripts/glass.py` measures the worst decile of the refracted face beneath each line: target 3.5:1 at ≥40 px, 5.0:1 for the smaller label.
- Keep the tint at or below .25. Placement search precedes fallback; it never relaxes the contrast target or the cap.

## 7. Plain callouts: every engineering figure

- Use a native rounded rectangle 48 px tall with radius 24.
- Fill #1d1d1f, with white 22 px semibold text, centred horizontally and vertically.
- Place it beside its marker with a 12 px gap. Keep one style for every callout on a figure; never mix glass and plain callouts on one figure.

## 8. How many

- **Outside a bento:** at most 3 glass objects per slide. Keep the verified maximum of three outside a bento; charts have the stricter limit of two.
- **In a bento:** at most 5 tiles.
- **Per section:** at most 1–2 glass slides. Never on an evidence slide. This proof set collects cases for comparison; it is not a report sequence.

## 9. PowerPoint limits (tell the author)

**How each glass box is built**
- It is one baked 2× JPEG face, rendered over the real content beneath it.
- The face is cropped to a native rounded rectangle and has a native soft shadow.
- All text on it is native and sits above the picture.
- The content beneath stays native: text, bars, pictures.

**The refraction is frozen**
- After moving or resizing a glass box, editing anything beneath it, or lengthening its text, re-author the HTML to fit, re-run `scripts/html_to_pptx.py`, then `scripts/render_pptx.py --app powerpoint`. Calibration and rim notes are part of the conversion, not separate gates.
- Caption frames use the capsule's inner width with wrapping on. The text is editable, but its baked face and text halo cannot grow with it. A longer caption can wrap into the value or beyond the face; reflow or enlarge the authored capsule and rebuild, never stretch the text box outside it.
- Never hand-move or resize a glass picture in PowerPoint. The picture's name in the Selection Pane says "glass picture: rebuild if moved or content beneath changes".
- A slide that must be edited live in PowerPoint should not use glass.
- Every glass slide’s speaker notes end with: “Glass is a baked picture: rebuild the deck after moving it, editing what is under it, or lengthening its text.” The converter appends this exact warning.

**Other limits**
- **File size:** about 1.5 MB per glass slide (6.5 MB for these 5 slides, 4 of them with glass).
- **Verification boundary:** the original prototype was checked through PowerPoint for Mac’s own export. Every newly built deck still needs its own named render comparison. A slideshow screen capture and Keynote remain unverified.
- **Confidential figures:** stay local in source and deck, never in the skill package. Snippets use generic placeholders; no Apple photos are packaged.

## 10. Plain bento, including a whole-photo fallback

No matching photo means no glass. Keep an `.image-needed` slot with its centred caption visible; arrange plain native tiles beside or below it. Use native rounded rectangles, 40 px radius and 24 px gutters, fill #1D1D1F on black and #F5F5F7 on white. Keep the same text sizes, inset ≥40 px, word ceiling, source fidelity and evidence rules. Use theme ink and muted tokens; status and source identity retain their separate roles. For `glass-callouts`, keep the capsule geometry with a native solid fill and editable text, beside the placeholder. Rebuild with glass when the user supplies the matching photo.

If every searched combination exceeds .25, show the photo whole in one uncovered rounded tile and put native `bento-plain` text tiles beside or below it on the page. Use the registered `with-photo-tile` variant. Never place opaque tiles over the photo or leave photographic slivers in gutters.

## 11. Authoring and rebuilding

Everything that can sit beneath glass belongs to the page's `.backdrop`. Glass boxes are direct children of `.page`, in front of that backdrop, with native `data-role="text"` children. Give each box a `data-name`. Use `data-tone="dark|light"`; the default follows the black/white page, but content determines the right choice.

For a bento, first crop a sharp local photo to the bento plus 70 px bleed on all four sides. Set the photo's `data-bento-bleed` to that file, and `data-lens-box="left,top,width,height"` to its slide placement. When a source lacks 70 px bleed, mirror-pad the lens source and print the padding. The renderer prepares the page mask from that page's glass tile geometry, using the round gutter ends; the same source backs every tile. The clipped page photo and the larger lens source share pixels. No photographic crop is generated or downloaded by the skill.

Verified lens defaults: bezel clamp(.3 × min(width,height),16,56); pull clamp(.35 × bezel + 6,10,26); magnification .06; radial .8; fringe 1.0; smoke blur 1.5 and saturation 1.5. Large bento tile radial .5; others .8. Optional attributes are `data-bezel`, `data-pull`, `data-mag`, `data-radial`, `data-fringe`, `data-blur`, `data-tint-base` and calibrated `data-tint`. Fine mesh can use fringe .6, blur .5 and magnification .10 as the verified product-photo capsule did.

Keep C1 light optics: pull 20, tint .06, saturation 1.15, blur 1, milk 0 on white; the accent capsule uses `class="glass dim"`, white text and .07 smoke. Keep dark/light rim layers in `glass.css` unchanged. No untested rim softening.

Calibration samples the worst decile beneath each visible text line with text hidden, then remeasures the refracted face until it reaches 5.0:1 under 40 px or 3.5:1 at ≥40 px. Tint is capped at .25 in every variant. If the initial position fails, search all five text zones and available photo offsets in 40 px steps, keeping at least 30 px optical sampling margin. Pick the lowest-tint combination; a tie keeps the least photo movement. Fall back to the whole-photo/plain-text variant only when every option exceeds .25. The final bake captures each face at 2× with only its own text transparent (halo retained), JPEG q93 with 4:4:4 chroma, 1 px outer-hairline padding, native roundRect crop and native outer shadow.

The template loads local `glass.css` and `glass.js`. Rendering and conversion share the bento preparation and calibration, so HTML and PPTX show the same material. `--app powerpoint` uses a temporary directory in PowerPoint's container for export and removes the temporary PDF after rendering PNGs. LibreOffice fallback is identified in output; it is not native PowerPoint proof.
