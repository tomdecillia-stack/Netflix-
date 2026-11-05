"""Build slide4.pptx for the Netflix Together strategy project (Person 4).

Run:  python3 build_slide4.py   ->  writes slide4.pptx next to this script.
Unofficial student concept; not affiliated with Netflix.
"""
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

OUT = Path(__file__).with_name("slide4.pptx")

BG = RGBColor(0x14, 0x14, 0x14)
CARD = RGBColor(0x22, 0x22, 0x22)
ROW_ALT = RGBColor(0x1C, 0x1C, 0x1C)
HIGHLIGHT = RGBColor(0x3A, 0x0B, 0x0E)
RED = RGBColor(0xE5, 0x09, 0x14)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0xB3, 0xB3, 0xB3)
DIM = RGBColor(0x80, 0x80, 0x80)
FONT = "Arial"
SYMBOL_FONT = "Segoe UI Symbol"

YES, PART, NO = "✓", "~", "✕"

COLS = ["Effort to choose", "Fits both viewers", "Private votes", "Same-TV use"]
ROWS = [
    ("Netflix today", ["High", NO, NO, YES]),
    ("AI Discover (US pilot)", ["Medium", NO, NO, YES]),
    ("Watch-party tools", ["High", PART, NO, NO]),
    ("Netflix Together (target)", ["Low", YES, YES, YES]),
]

NOTES = """[~2:00] Slide 4 - strategic fit.

(0:00-0:20) Our claim: Netflix Together fits Netflix's strategy where rivals' group features don't. Ratings in this table are our team's judgement, based on the sources in the footer.

(0:20-0:45) Netflix today: one profile, a lot of scrolling, nothing built for two people choosing together. AI Discover launched on 3 November as a US, English-language, TV-only test. You pick a mood such as "scare me", which shortens the choice, but it is personalised to one profile. It does not help two people agree.

(0:45-1:05) Watch-party tools solve a different problem: syncing separate screens. Disney+ removed GroupWatch in September 2023, and Hulu's Watch Party works in a web browser only. We found no evidence of a major streamer offering private voting for two or more people on one TV.

(1:05-1:30) Netflix Together combines AI Discover's mood idea with a 3-film shortlist, private votes and a mutual match, all on the same TV. The benefits are hypotheses we would test in a UK pilot: faster choice, higher satisfaction, and retention. Retention matters more after UK prices rose on 3 September to 7.99, 13.99 and 20.99 pounds.

(1:30-1:50) Costs: we name categories, not figures. Design, engineering, recommender changes for two people, testing, and maintenance. Netflix can afford a focused test: Q3 2026 revenue was 12.87 billion dollars with a 34.6% operating margin. Ad revenue guidance was cut to about 2.7 billion, so engagement matters, but these numbers show capacity to invest, not the return.

(1:50-2:00) Positioning: Less scrolling. More watching. That is the one-line reason Netflix should test this in the UK."""


def set_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_text(slide, x, y, w, h, runs, size=14, color=WHITE, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, name=None):
    """runs: str, or list of paragraphs; each paragraph is a list of (text, opts) tuples."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = [[(runs, {})]] if isinstance(runs, str) else runs
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for text, o in para:
            r = p.add_run()
            r.text = text
            f = r.font
            f.name = o.get("font", FONT)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
        if "space_after" in (para[0][1] if para else {}):
            p.space_after = Pt(para[0][1]["space_after"])
    return tb


def cell_border_bottom(cell, color="333333", width_emu=9525):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(tag))
        if tag == "a:lnB":
            ln.set("w", str(width_emu))
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
    # schema order: borders must precede fill in tcPr
    fill = tcPr.find(qn("a:solidFill"))
    if fill is not None:
        tcPr.remove(fill)
        tcPr.append(fill)


def style_cell(cell, text, fill, color=WHITE, size=14, bold=False,
               align=PP_ALIGN.CENTER, font=FONT):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = cell.margin_right = Inches(0.1)
    cell.margin_top = cell.margin_bottom = Inches(0.04)
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG

    # Kicker + title
    add_text(slide, 0.6, 0.4, 8, 0.3, "04  ·  STRATEGIC FIT", size=12,
             color=RED, bold=True, name="Kicker")
    add_text(slide, 0.6, 0.75, 12.1, 1.1,
             "Netflix Together fits Netflix’s strategy where rivals’ group features don’t",
             size=30, bold=True, name="Title")

    # Comparison table (left)
    tx, ty, label_w, col_w, row_h = 0.6, 2.15, 2.4, 1.3, 0.72
    shape = slide.shapes.add_table(len(ROWS) + 1, len(COLS) + 1, Inches(tx), Inches(ty),
                                   Inches(label_w + col_w * len(COLS)),
                                   Inches(row_h * (len(ROWS) + 1)))
    shape.name = "Comparison table"
    tbl_xml = shape._element.graphic.graphicData.tbl
    # "No Style, No Grid" so our fills/borders are the only styling
    tbl_xml.tblPr.find(qn("a:tableStyleId")).text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    tbl = shape.table
    tbl.first_row = True
    tbl.horz_banding = False
    tbl.columns[0].width = Inches(label_w)
    for c in range(1, len(COLS) + 1):
        tbl.columns[c].width = Inches(col_w)
    for r in range(len(ROWS) + 1):
        tbl.rows[r].height = Inches(row_h)

    style_cell(tbl.cell(0, 0), "", BG)
    for c, head in enumerate(COLS, start=1):
        style_cell(tbl.cell(0, c), head, BG, color=GREY, size=12, bold=True)
    for r, (label, vals) in enumerate(ROWS, start=1):
        ours = label.startswith("Netflix Together")
        fill = HIGHLIGHT if ours else (CARD if r % 2 else ROW_ALT)
        style_cell(tbl.cell(r, 0), label, fill, color=WHITE, size=14, bold=ours,
                   align=PP_ALIGN.LEFT)
        for c, v in enumerate(vals, start=1):
            symbol = v in (YES, PART, NO)
            if ours:
                col = RED if v in (YES, "Low") else WHITE
            else:
                col = WHITE if v == YES else GREY
            style_cell(tbl.cell(r, c), v, fill, color=col,
                       size=20 if symbol else 14, bold=ours or v == YES,
                       font=SYMBOL_FONT if v in (YES, NO) else FONT)
    for r in range(len(ROWS) + 1):
        for c in range(len(COLS) + 1):
            cell_border_bottom(tbl.cell(r, c))

    legend_y = ty + row_h * (len(ROWS) + 1) + 0.18
    add_text(slide, tx, legend_y, 7.6, 0.3,
             [[(YES, {"font": SYMBOL_FONT}), (" yes   ", {}), ("~", {}), (" partly   ", {}),
               (NO, {"font": SYMBOL_FONT}), (" no   ·   ratings = team judgement", {})]],
             size=11, color=GREY, name="Legend")

    # Right-hand cards
    cx, cw, card_h, gap = 8.95, 3.78, 1.16, 0.2
    cards = [
        ("BENEFITS (HYPOTHESES)", "Faster choice, higher satisfaction, retention", 15, False),
        ("COSTS (CATEGORIES)", "Design, engineering, recommender changes, testing, maintenance", 15, False),
        ("POSITIONING", "“Less scrolling. More watching.”", 18, True),
    ]
    for i, (head, body, body_size, bold) in enumerate(cards):
        y = ty + i * (card_h + gap)
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(y),
                                      Inches(cw), Inches(card_h))
        card.name = f"Card {i + 1}"
        card.adjustments[0] = 0.08
        set_fill(card, HIGHLIGHT if i == 2 else CARD)
        add_text(slide, cx + 0.22, y + 0.16, cw - 0.44, 0.25, head, size=11, color=RED,
                 bold=True, name=f"Card {i + 1} heading")
        add_text(slide, cx + 0.22, y + 0.43, cw - 0.44, card_h - 0.5, body,
                 size=body_size, bold=bold, name=f"Card {i + 1} body")

    # Footer
    add_text(slide, 0.6, 6.78, 12.1, 0.4,
             "Sources: Netflix Help, TechCrunch (Nov 2026); TechRadar (Sep 2023); "
             "iClarified (May 2026). Unofficial student concept.",
             size=10, color=DIM, name="Footer")

    slide.notes_slide.notes_text_frame.text = NOTES
    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
