# Apple_Slides

An agent skill for Claude Code and Codex. It turns a report into an **editable PowerPoint deck** in the engineering style of Apple's public Environmental Progress Reports: assertion headlines, labelled evidence, native charts and tables, and few words.

Most AI slide skills produce decks that look AI-made: decorative, brief and hard to edit. This skill takes its rules from the reports' own pages and checks every deck in real PowerPoint.

## What you get

- **Native PowerPoint:** text, tables, charts and shapes stay editable, and every slide has a speaker-notes script.
- **16 layouts:** status summary, milestone timeline, goal progress, figure callouts, peer columns, chart interpretation, multi-year table, decision tables, evidence actions, bento and others.
- **Density rules per slide:** at least 10 labelled values in 2–4 evidence objects, at most 90 words, headlines of 8 words or fewer, body type of at least 18 px.
- **Pictures that fit the slide:** every photo shows what its slide is about. If you supply none, the slide gets an editable "Image needed" placeholder, and the agent asks you for the photo.
- **Optional Liquid Glass:** baked glass faces with native text on top, for photo-backed summary slides. It falls back to a clean layout when a photo is too busy.

## Requirements

- macOS with SF Pro Text and SF Pro Display installed (not included here).
- Google Chrome, plus PowerPoint for Mac. LibreOffice works as a preview fallback.
- Python 3.10 or later with `python-pptx`, `playwright`, `Pillow`, `numpy` and `pymupdf`.

## Install

Clone the repo, then link the skill folder into your agent's skills directory:

```bash
git clone https://github.com/stanleyfeng1122-ui/Apple_Slides.git
ln -s "$PWD/Apple_Slides/apple-report-slides" ~/.claude/skills/apple-report-slides
ln -s "$PWD/Apple_Slides/apple-report-slides" ~/.codex/skills/apple-report-slides
```

## Use

Ask in plain words, for example: *"Make a management deck from this report for my engineering directors."* Give the source (PDF, PPTX, Keynote export or notes) and its classification mark. The agent writes `outline.md` (one line per slide) for you to check, then builds `deck.html` and converts it to `deck.pptx`. Everything runs locally; nothing is uploaded.

Don't use it for marketing or pitch decks, single charts, or small text fixes to an existing PPTX.

## License

MIT, see [LICENSE](LICENSE). Third-party code keeps its own notice in `apple-report-slides/assets/THIRD_PARTY.md`.

## Credits and notice

- The HTML-first method follows [frontend-slides](https://github.com/zarazhangrui/frontend-slides) and [guizang-ppt-skill](https://github.com/op7418/guizang-ppt-skill). No code or templates were copied from them.
- The glass optics are adapted from [rdev/liquid-glass-react](https://github.com/rdev/liquid-glass-react) (MIT). See `apple-report-slides/assets/THIRD_PARTY.md`.
- This is an independent project. It is not affiliated with or endorsed by Apple Inc. Apple and SF Pro are trademarks of Apple Inc.
