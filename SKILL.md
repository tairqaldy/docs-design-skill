---
name: docs-design
description: House style for every finished document — assignments, lab reports, practice submissions, essays, research notes, papers, handouts, work reports. Produces a typeset-looking PDF (plus editable DOCX) in a classical academic style (LaTeX article / journal paper / Tufte essay) with serif type, real small caps, old-style figures, booktabs tables, Computer Modern maths, side notes and monochrome diagrams, and text that reads like a person wrote it. Use this whenever the user asks to "make a report", "write up the assignment", "create a document/PDF/handout", "prepare the submission", "format this nicely", or hands over task files to be completed and submitted — even when design is never mentioned. Prefer it over plain python-docx/reportlab output or generic docx/pdf skills whenever the deliverable is a finished document someone will read; use those only for editing an existing file in place, tracked changes, or a mandated template.
---

# docs-design

The goal is documents that look typeset by someone who cares (LaTeX, a good journal, a Tufte
handout) and text that reads like a person wrote it. Generic AI output fails on both counts:
Calibri, blue headings, shaded grid tables, bullet storms, and padded prose. This skill is the
fix. It has three parts, and a deliverable needs all three:

1. **Content** that sounds human: read `references/writing.md` before writing a single sentence.
2. **Layout** from `scripts/tairdoc.py`, which encodes the design rules so they don't have
   to be re-derived. The reasons behind them are in `references/design-guide.md`.
3. **Inspection**: render the PDF, look at every page, fix, re-render.

`SKILL_DIR` below means the folder containing this file. Resolve it to an absolute path once and
use that.

## 0 · Setup (first use on a machine)

```bash
pip install -r SKILL_DIR/requirements.txt     # python-docx, matplotlib, Pillow, PyMuPDF
```

LibreOffice must be installed for PDF output, and it also brings the main font, Linux Libertine G.
Check with `soffice --version`. If it is missing, ask the user to install it from
libreoffice.org, or deliver the DOCX only and say why.

## 1 · Pick a preset

| Preset | Looks like | Use for |
|---|---|---|
| `article` (default) | LaTeX `article`: centred title, numbered bold headings, Libertine + CM maths | assignments, lab/practice reports, problem sets, technical write-ups |
| `journal` | old journal paper: title between hairlines, small-caps author, abstract, keywords, letter-spaced section heads, running footer | research papers, literature reviews, anything with an abstract |
| `essay` | Tufte handout: italic left-aligned title, Palatino, wide margin with side notes and margin captions | essays, reflective writing, notes, explainers, reading responses |

If the user gives a sample or names a look, map it to the closest preset. If unsure, use `article`.

## 2 · Write the content first

Draft everything as plain text before touching layout. Follow `references/writing.md`: answer
first, concrete numbers, prose over bullets, and none of the stock phrases it lists. For an
assignment, keep the task's own structure (Part A, Q1…) so a grader can find each answer, and
show working where the task asks for it.

## 3 · Build it

```python
import sys; sys.path.insert(0, r"SKILL_DIR/scripts")
from tairdoc import Doc, Diagram, mpl_style, render_terminal, preview

d = Doc("Practice 6: Bayesian Inference and HMMs", author="Jane Doe",
        email="jane@uni.edu", affiliation="AI Fundamentals · Group 3",
        date="25 October 2026", theme="article")       # journal: running=, keywords=[...]
d.abstract("…")                     # optional; essay renders it as a 'Summary' block
d.contents()                        # optional book-style contents page, real page numbers
d.section("Bayes’ theorem")         # numbered=False for unnumbered ("Part A: …" style)
d.subsection("The 100-case table")
d.para("Text with **bold**, *italic*, `code`, $x_t$ and a note.^[Side note in essay, footnote-style elsewhere.]",
       lead="Question 1.")          # lead = bold run-in label
n = d.equation(r"P(O\mid B)=\frac{P(B\mid O)\,P(O)}{P(B)}")   # CM maths, numbered (n) at the right
t = d.table([["Time", "Reading", "P(O)"], ["t1", "BLOCKED", "0.6316"]], caption="Filtering results.")
f = d.figure("plot.png", "What the figure shows, in one sentence.")
d.para(f"Table {t} and Figure {f} agree with equation ({n}).")     # numbers are returned: refer, never hard-code
d.callout("The one thing the reader must not miss.", kind="info")   # info|note|tip|answer|warning
d.code(open("solution.py").read(), caption="solution.py")
d.bullets(["Only for real lists."], numbered=False)                 # or [("Lead.", "text"), …]
d.quote("A short quotation.", "Author")
d.pagebreak()
docx, pdf = d.save("Doe_Jane_Practice6.docx")                       # writes .docx and .pdf
preview(pdf)                                                        # page PNGs in ./_preview
```

Inline markup in any text argument: `**bold**`, `*italic*`, `` `mono` ``, `$math-ish$` (italic
lining figures), `^[note]`. Straight quotes and `--`/`---` become typographic quotes and dashes
automatically, and ` | ` gets non-breaking spaces so `P(A | B)` never splits. Write
multiplication as `×`, never `*`, because `*` is the italic marker.

### Figures

Every figure must look like it belongs to the page: monochrome, the document's serif, thin lines.

- **Plots:** call `mpl_style(theme)` before plotting with matplotlib. Label lines directly instead
  of using a legend where possible, put no title inside the plot (the caption is the title), and
  save at 300 dpi.
- **Diagrams** (networks, pipelines, state machines, HMMs): use `Diagram`. Coordinates are in
  centimetres on the final figure, so make it the width it will print at (≤ text width, or wider
  for essay figures that span the margin).
  ```python
  g = Diagram(15, 3, theme="article")
  g.node("w", 2, 1.5, "Weather", "Dry / Rainy"); g.node("p", 7.5, 1.5, "Path", "Obstacle / Clear")
  g.node("x", 12, 1.5, "$X_1$", shape="circle", w=1.2, h=1.2)
  g.edge("w", "p", "influences")              # rad= for curves, dashed=True, side=-1 flips label
  g.save("bn.png")
  ```
  Leave at least ~2 cm between boxes when an edge has a label.
- **Program output / "screenshot" requirements:** use `render_terminal(output_text, "out.png",
  command="python solution.py")`. Run the real program to get that text; never type output by hand.
- **Real screenshots** go in as ordinary figures with a caption.

## 4 · Look at it before handing it over

Render with `preview(pdf)` and read every PNG. Check:

- Nothing overflows, no figure or table is split from its caption, and no heading sits alone at
  the bottom of a page.
- No last page holding only a line or two, and no orphaned single word ending a paragraph. Tighten
  the wording rather than shrinking the type.
- Tables are sized to their content, and numbers line up.
- Every equation, table, and figure is referred to in the text.
- The title block holds the right name, group, and date. Ask for any of these you don't know; never
  invent them, and leave a visible blank (`Group ____`) rather than a guess.

Fix and rebuild until a page could pass for LaTeX output. Then deliver the **PDF** (primary), the
**DOCX** (editable), and any source code the task asked for, named the way the course or client
wants (often `Surname_Name_Task.pdf`).

## Notes

- Fonts: Linux Libertine G (ships with LibreOffice) for `article` and `journal`; Palatino for
  `essay` (Palatino Linotype on Windows, Palatino on macOS, TeX Gyre Pagella or URW P052 on
  Linux). The library picks whichever is installed.
- `examples/showcase.py` builds a specimen of each preset into `examples/out/`. Run it after
  changing `tairdoc.py` and look at the output: it is the regression test.
- Need something the library lacks (a two-column journal body, a new block type)? Add it to
  `tairdoc.py` in the same style rather than hand-formatting one document.
