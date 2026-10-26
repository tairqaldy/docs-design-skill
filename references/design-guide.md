# Design guide

The rules `tairdoc.py` implements, and why. They come from the standard references:
Bringhurst's *The Elements of Typographic Style*, Butterick's *Practical Typography*,
Tufte's handouts and *tufte-latex*, LaTeX's `article`, `booktabs` and `memoir` defaults,
and the journal and book pages that inspired the presets. Read this when you change the
library or hand-format something it doesn't cover.

## The look in one paragraph

A page should read as typeset rather than "designed". Use one serif family with real small caps
and old-style figures. Build the hierarchy from size, weight, italics and small caps, never from
colour. Draw only hairline rules and no boxes. Leave generous margins, and keep the line length
comfortable. Figures should be drawn in the same ink and the same type as the text. The
reference pages behind the presets share this restraint: LaTeX `article`, *The Computer Journal*, a
Garamond book contents page, and a Tufte essay.

## Type

| Element | Rule | Why |
|---|---|---|
| Body | Libertine (article/journal) or Palatino (essay), 10.5–11 pt, justified, auto-hyphenated | Justified text without hyphenation gets rivers; LaTeX always hyphenates. |
| Leading | 1.15–1.3 × size | Tight enough to stay a block, loose enough for long reading; essays get more. |
| Measure | 45–75 characters per line (article ≈ 74, essay ≈ 58) | Longer lines lose the reader on the return sweep. This is why margins are wide. |
| Paragraphs | First-line indent of 1–1.5 em, no space between, no indent after a heading | The book/LaTeX convention. Space plus indent says the same thing twice. |
| Figures (numerals) | Old-style (`onum`) in prose, lining tabular (`lnum&tnum`) in tables and equations | Old-style figures sit in running text like lowercase; tabular figures line up in columns. |
| Small caps | True small caps (`:smcp`) for author names, captions (journal), running heads, the contents title | Synthetic small caps look thin and spotty. |
| Emphasis | Italics; bold only for headings and run-in labels | A page sprinkled with bold looks like a slide deck. |
| Monospace | Consolas, ~1.2 pt smaller than body, only for code, file names and emails | |
| Maths | Computer Modern; display fractions full size; equations numbered flush right | LaTeX's own look, and it signals "this was done properly". |
| Quotes and dashes | ‘ ’ “ ” – — (auto-converted) | Straight quotes are the commonest sign of untypeset text. |

## Hierarchy

At most three heading levels. Article headings are bold and numbered with an em space
(`1 Title`, `1.1 Title`). Journal section heads are small, bold, letter-spaced capitals with
italic subsections. Essay headings are italic, unbolded and unnumbered. Every heading has more
space above than below, so it belongs to the text that follows it. Headings never end with
a full stop.

## Tables (booktabs)

- Use three rules only: a heavy rule on top, a light rule under the header, and a heavy rule at the bottom.
- Use no vertical lines, no shading, and no zebra stripes.
- Size the table to its content (natural width). Full text width only when content demands it.
- Centre numeric columns, left-align text, and keep one precision per column.
- The caption goes **above** the table, figure captions go **below** the figure, and both are referenced in the text.
- Never draw a table from a single row of key–value pairs that reads better as a sentence.

## Figures

- Use black, greys and white only; distinguish series by line style (solid, dashed, dotted) or direct labels.
- Use the document's serif for every label, and CM for maths in labels.
- Put no title inside the plot (the caption is the title) and no chart junk: no gridlines unless values must be read off, no 3D, no gradients, no drop shadows.
- Draw diagrams like TikZ: thin black boxes with small rounded corners, labels inside, italic edge labels, and light arrowheads.
- Size the figure at the width it will be printed so the text in it matches the body size (roughly 8–10 pt). A figure scaled down 50% has unreadable 5 pt labels.
- In essays, captions go in the margin and wide figures span text plus margin.

## Asides

- **Callout:** one per page at most, for something the reader must not miss (a warning, the key result). Use an icon and a hairline rule, never a coloured box.
- **Side notes** (essay) or **notes** (article/journal): for references, asides and caveats that would interrupt the sentence.
- **Block quotes:** indented on both sides, slightly smaller type, and no quotation marks.

## Title blocks

- **Article:** centred. Title 18 pt regular (not bold), then name, a mono email, and `Affiliation — Date`.
- **Journal:** a hairline rule above and below. Title 21 pt, author in small caps, then affiliation, email and date in small italics, an indented abstract, and a keywords line.
- **Essay:** left-aligned italic title, author and date, with a small-caps affiliation line and an indented "Summary" block.

## Never

- Use sans-serif body text, Calibri, Arial or Times New Roman.
- Colour headings, shade table headers, use grid tables, add emoji, or put icons next to headings.
- Centre body text, set it in all caps, or underline anything except links.
- Leave space between paragraphs **and** indent them.
- Put a "Table of Contents" in a document shorter than about six pages.
- Drop in a screenshot of code; set code as code, and only output goes in a rendered output figure.
