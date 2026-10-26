"""Builds one specimen document per preset into ./out (run: python showcase.py).
It doubles as a regression test for tairdoc: every component appears at least once."""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
from tairdoc import Diagram, Doc, mpl_style, preview, render_terminal  # noqa: E402

OUT = HERE / "out"
OUT.mkdir(exist_ok=True)


def measure_plot(path, theme):
    mpl_style(theme)
    import matplotlib.pyplot as plt
    chars = list(range(30, 121, 5))
    speed = [100 - 0.012 * (c - 68) ** 2 for c in chars]
    fig, ax = plt.subplots(figsize=(4.6, 2.3))
    ax.plot(chars, speed)
    ax.axvspan(45, 75, color="#e9e9e9", lw=0)
    ax.text(60, 62, "comfortable\nrange", ha="center", fontsize=8.5, style="italic")
    ax.set_xlabel("characters per line")
    ax.set_ylabel("relative reading speed")
    ax.set_ylim(55, 103)
    fig.savefig(path)
    plt.close(fig)


def pipeline(path, theme):
    g = Diagram(15.6, 2.4, theme=theme)
    g.node("src", 1.6, 1.2, "Draft", "plain text")
    g.node("doc", 5.9, 1.2, "Doc()", "tairdoc")
    g.node("docx", 9.6, 1.2, ".docx")
    g.node("pdf", 14.1, 1.2, ".pdf", emphasis=True)
    g.edge("src", "doc", "markup")
    g.edge("doc", "docx", "save")
    g.edge("docx", "pdf", "LibreOffice")
    g.save(path)


def body(d, theme):
    d.section("Why the line length matters")
    d.para("Most of what makes a page tiring has nothing to do with the font. It is the length of the "
           "line. Past about seventy-five characters the eye loses its place on the way back to the left "
           "margin, and readers start skipping lines without noticing.^[Bringhurst puts the comfortable "
           "range at 45–75 characters for a single column; Butterick says much the same.]")
    d.para("Below forty-five the opposite happens: the reader spends more time jumping between lines than "
           "reading them, and justified text develops gaps. The specimen in Figure 1 is invented, but the "
           "shape is the one every style guide describes.")
    d.figure(OUT / f"measure_{theme}.png", "Reading speed against line length (illustrative curve).")
    d.subsection("What the presets do about it")
    d.para("Each preset fixes the text width so that 11-point type lands inside the range. The essay "
           "preset narrows it further and hands the spare width to side notes, which is where Tufte "
           "keeps references, small figures and asides that would otherwise interrupt a sentence.")
    d.table([["Preset", "Body type", "Text width", "Characters / line"],
             ["article", "Libertine 11 pt", "14.8 cm", "≈ 74"],
             ["journal", "Libertine 10.5 pt", "16.2 cm", "≈ 86 (or 2 × 40)"],
             ["essay", "Palatino 11 pt", "10.8 cm", "≈ 58"]],
            caption="Measure by preset, measured on this specimen.")
    d.section("A worked formula")
    d.para("A filter that combines a prior belief with one noisy reading uses nothing more than Bayes' rule:")
    n = d.equation(r"P(H\mid E)=\frac{P(E\mid H)\,P(H)}{P(E\mid H)\,P(H)+P(E\mid \neg H)\,P(\neg H)}")
    d.para(f"Equation ({n}) is set in Computer Modern, the same face LaTeX uses, and numbered at the right "
           "margin. With a prior of 0.30 and a sensor that is right 80% of the time, one positive reading "
           "moves the posterior to $0.24/0.38 = 0.632$.")
    d.callout("The posterior is lower than the sensor's 80% accuracy because the prior is low. A rare event "
              "stays fairly unlikely even after one good piece of evidence.", kind="info")
    d.subsection("Code and output")
    d.code("def bayes(prior, hit, false_alarm):\n    num = prior * hit\n    return num / (num + (1 - prior) * false_alarm)\n\nprint(round(bayes(0.30, 0.80, 0.20), 4))",
           caption="bayes.py")
    render_terminal("0.6316", OUT / "term.png", command="python bayes.py", width_cm=9)
    d.figure(OUT / "term.png", "Program output.", width_cm=9)
    d.section("The pipeline")
    pipeline(OUT / f"pipe_{theme}.png", theme)
    d.figure(OUT / f"pipe_{theme}.png", "From draft to PDF.")
    d.bullets(["Write the text first, in plain sentences.",
               "Build it with one call per element; the preset handles spacing.",
               "Open the rendered pages and fix what looks wrong before sending anything."], numbered=True)
    d.quote("Typography exists to honour content.", "Robert Bringhurst")


for theme in ("article", "journal", "essay"):
    measure_plot(OUT / f"measure_{theme}.png", theme)
    d = Doc("On the Length of a Line", author="Tair Kaldybayev",
            affiliation="docs-design specimen", date="25 October 2026", theme=theme,
            running="docs-design specimen", keywords=["typography", "measure", "document design"])
    if theme != "article":
        d.abstract("A short specimen of the three presets. It shows headings, side notes, a booktabs table, "
                   "display maths, a callout, code with its output, diagrams and lists, so the presets can "
                   "be compared on the same text.")
    if theme == "article":
        d.contents()
    body(d, theme)
    docx, pdf = d.save(OUT / f"specimen_{theme}.docx")
    print(theme, pdf, [p.name for p in preview(pdf, dpi=70)])
