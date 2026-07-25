#!/usr/bin/env python3
"""Generate a polished PowerPoint presentation for Lindsey."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# Visual direction: deep forest + soft parchment + copper accent
FOREST = RGBColor(0x1A, 0x2F, 0x23)
SAGE = RGBColor(0x3D, 0x5C, 0x4A)
COPPER = RGBColor(0xC4, 0x7E, 0x5A)
PARCHMENT = RGBColor(0xF7, 0xF2, 0xEA)
INK = RGBColor(0x1E, 0x1A, 0x16)
MUTED = RGBColor(0x5C, 0x54, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

OUTPUT = Path(__file__).resolve().parent / "Lindsey_Presentation.pptx"


def set_run(run, *, size, bold=False, color=INK, font="Georgia"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, left, top, width, height, text, *, size=18, bold=False, color=INK, font="Georgia", align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf.vertical_anchor = anchor
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def add_paragraphs(box, lines, *, size=18, color=INK, font="Calibri", bold=False, space_after=10, bullet=False):
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(space_after)
        if bullet:
            p.level = 0
        run = p.add_run()
        run.text = (("•  " if bullet else "") + line)
        set_run(run, size=size, bold=bold, color=color, font=font)


def fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    fill_solid(shape, color)
    return shape


def add_accent_bar(slide, left, top, width=Inches(1.2), height=Inches(0.08)):
    return add_rect(slide, left, top, width, height, COPPER)


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # blank


def style_title_slide(slide):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, FOREST)
    # Soft band at bottom
    add_rect(slide, 0, Inches(5.9), SLIDE_W, Inches(1.6), SAGE)
    add_accent_bar(slide, Inches(0.9), Inches(2.35), width=Inches(1.6))
    add_textbox(
        slide, Inches(0.9), Inches(2.5), Inches(11), Inches(1.2),
        "Lindsey", size=60, bold=True, color=WHITE, font="Georgia",
    )
    add_textbox(
        slide, Inches(0.9), Inches(3.7), Inches(10), Inches(0.6),
        "A clear story, thoughtfully told.", size=24, color=RGBColor(0xD8, 0xE0, 0xD4), font="Calibri",
    )
    add_textbox(
        slide, Inches(0.9), Inches(6.25), Inches(10), Inches(0.4),
        "Overview presentation", size=16, color=PARCHMENT, font="Calibri",
    )


def style_section_header(slide, eyebrow, title, subtitle):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_rect(slide, 0, 0, Inches(0.22), SLIDE_H, FOREST)
    add_textbox(
        slide, Inches(0.9), Inches(2.2), Inches(11), Inches(0.4),
        eyebrow.upper(), size=14, bold=True, color=COPPER, font="Calibri",
    )
    add_accent_bar(slide, Inches(0.9), Inches(2.7))
    add_textbox(
        slide, Inches(0.9), Inches(2.95), Inches(11), Inches(1),
        title, size=40, bold=True, color=FOREST, font="Georgia",
    )
    add_textbox(
        slide, Inches(0.9), Inches(4.1), Inches(10), Inches(0.8),
        subtitle, size=20, color=MUTED, font="Calibri",
    )


def style_content_slide(slide, title, bullets, footer="Lindsey"):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.35), FOREST)
    add_textbox(
        slide, Inches(0.9), Inches(0.4), Inches(11), Inches(0.7),
        title, size=28, bold=True, color=WHITE, font="Georgia",
    )
    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.9), Inches(11.2), Inches(4.5))
    add_paragraphs(box, bullets, size=22, color=INK, font="Calibri", space_after=16, bullet=True)
    add_textbox(
        slide, Inches(0.9), Inches(6.9), Inches(6), Inches(0.35),
        footer, size=12, color=MUTED, font="Calibri",
    )


def style_two_column(slide, title, left_title, left_items, right_title, right_items):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.35), FOREST)
    add_textbox(
        slide, Inches(0.9), Inches(0.4), Inches(11), Inches(0.7),
        title, size=28, bold=True, color=WHITE, font="Georgia",
    )

    # Left panel
    add_rect(slide, Inches(0.75), Inches(1.8), Inches(5.6), Inches(4.6), WHITE)
    add_accent_bar(slide, Inches(1.05), Inches(2.1), width=Inches(0.9))
    add_textbox(
        slide, Inches(1.05), Inches(2.35), Inches(5), Inches(0.5),
        left_title, size=22, bold=True, color=FOREST, font="Georgia",
    )
    left_box = slide.shapes.add_textbox(Inches(1.05), Inches(3.0), Inches(5), Inches(3))
    add_paragraphs(left_box, left_items, size=18, color=INK, font="Calibri", space_after=12, bullet=True)

    # Right panel
    add_rect(slide, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.6), WHITE)
    add_accent_bar(slide, Inches(7.2), Inches(2.1), width=Inches(0.9))
    add_textbox(
        slide, Inches(7.2), Inches(2.35), Inches(5), Inches(0.5),
        right_title, size=22, bold=True, color=FOREST, font="Georgia",
    )
    right_box = slide.shapes.add_textbox(Inches(7.2), Inches(3.0), Inches(5), Inches(3))
    add_paragraphs(right_box, right_items, size=18, color=INK, font="Calibri", space_after=12, bullet=True)


def style_closing(slide):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, FOREST)
    add_accent_bar(slide, Inches(0.9), Inches(2.6), width=Inches(1.6))
    add_textbox(
        slide, Inches(0.9), Inches(2.85), Inches(11), Inches(1),
        "Thank you", size=48, bold=True, color=WHITE, font="Georgia",
    )
    add_textbox(
        slide, Inches(0.9), Inches(4.0), Inches(10), Inches(0.6),
        "Let’s keep the conversation going.", size=22, color=RGBColor(0xD8, 0xE0, 0xD4), font="Calibri",
    )
    add_textbox(
        slide, Inches(0.9), Inches(6.3), Inches(10), Inches(0.4),
        "Lindsey  ·  Overview presentation", size=14, color=PARCHMENT, font="Calibri",
    )


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # 1. Title
    style_title_slide(blank_slide(prs))

    # 2. Agenda
    style_content_slide(
        blank_slide(prs),
        "Agenda",
        [
            "Who this is for",
            "What matters most",
            "How the work comes together",
            "Where things go next",
        ],
    )

    # 3. About
    style_section_header(
        blank_slide(prs),
        "Introduction",
        "Meet Lindsey",
        "A simple frame for who she is, what she values, and how she shows up.",
    )

    # 4. Values
    style_two_column(
        blank_slide(prs),
        "What guides the work",
        "Values",
        [
            "Clarity over clutter",
            "Care in the details",
            "Honesty in the message",
            "Warmth without noise",
        ],
        "Approach",
        [
            "Start with the audience",
            "Say the important thing first",
            "Keep the design calm and intentional",
            "Leave room to breathe",
        ],
    )

    # 5. Focus areas
    style_content_slide(
        blank_slide(prs),
        "Focus areas",
        [
            "Storytelling that feels personal and grounded",
            "Presentations that are easy to follow and hard to forget",
            "Visual systems that stay consistent across formats",
            "Collaboration that moves from idea to finished piece",
        ],
    )

    # 6. Looking ahead
    style_content_slide(
        blank_slide(prs),
        "Looking ahead",
        [
            "Sharpen the narrative for new audiences",
            "Build a reusable slide system for future decks",
            "Pair strong writing with restrained design",
            "Invite feedback and iterate with purpose",
        ],
    )

    # 7. Closing
    style_closing(blank_slide(prs))

    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    build()
