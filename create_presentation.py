#!/usr/bin/env python3
"""Lindsey Dempsey — EMEA SDR Leader Challenge PowerPoint.

Full recreate from past prep history:
  - Cursor competitor differentiation
  - Enterprise study guide + customer proof
  - Whiteboard bottleneck / software-factory framing
  - Official positioning + verbatim speaker notes
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

# Cursor brand tokens (cursor.com/brand + marketing system)
PARCHMENT = RGBColor(0xF7, 0xF7, 0xF4)  # canvas
BONE = RGBColor(0xF2, 0xF1, 0xED)
LINEN = RGBColor(0xE6, 0xE5, 0xE0)
INK = RGBColor(0x26, 0x25, 0x1E)  # warm near-black
MUTED = RGBColor(0x80, 0x7D, 0x72)
EMBER = RGBColor(0xF5, 0x4E, 0x00)  # Cursor orange — accent only
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xEF, 0xEE, 0xE8)

# Aliases so existing slide builders keep working
CHARCOAL = INK
SLATE = RGBColor(0x3A, 0x38, 0x32)
PAPER = PARCHMENT
CITRUS = EMBER

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "Lindsey_Presentation.pptx"
LOGO_DARK = ROOT / "assets" / "cursor-lockup-dark.png"
LOGO_LIGHT = ROOT / "assets" / "cursor-lockup-light.png"
CUBE_DARK = ROOT / "assets" / "cursor-cube-dark.png"

FONT = "Arial"  # closest widely available stand-in for CursorGothic


def set_run(run, *, size, bold=False, color=INK, font=FONT):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_textbox(slide, left, top, width, height, text, *, size=18, bold=False, color=INK, font=FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
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


def add_bullets(slide, left, top, width, height, lines, *, size=17, color=INK, space_after=10):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = f"•  {line}"
        set_run(run, size=size, color=color)
    return box


def fill_solid(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    fill_solid(shape, color)
    return shape


def add_accent(slide, left, top, width=Inches(1.05), height=Inches(0.06)):
    """Thin ember accent bar — Cursor brand voltage, used sparingly."""
    return add_rect(slide, left, top, width, height, EMBER)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text.strip()


def add_logo(slide, *, light=False, left=Inches(0.75), top=Inches(0.28), width=Inches(1.55)):
    path = LOGO_LIGHT if light else LOGO_DARK
    if path.exists():
        slide.shapes.add_picture(str(path), left, top, width=width)


def header(slide, title):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.15), INK)
    add_logo(slide, light=True, left=Inches(0.7), top=Inches(0.32), width=Inches(1.35))
    add_textbox(slide, Inches(2.3), Inches(0.32), Inches(10.2), Inches(0.55), title, size=22, bold=False, color=WHITE, font=FONT)
    add_rect(slide, 0, Inches(1.15), SLIDE_W, Inches(0.06), EMBER)


def title_slide(prs, note):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_logo(s, light=False, left=Inches(0.85), top=Inches(0.55), width=Inches(2.1))
    add_accent(s, Inches(0.85), Inches(2.05), width=Inches(1.4))
    add_textbox(s, Inches(0.85), Inches(2.3), Inches(11.5), Inches(0.9), "Up-Level the Machine", size=40, bold=False, color=INK, font=FONT)
    add_textbox(s, Inches(0.85), Inches(3.3), Inches(11.5), Inches(0.45), "Diagnose  →  Protect  →  Experiment  →  Execute", size=20, color=EMBER, font=FONT)
    add_textbox(s, Inches(0.85), Inches(3.95), Inches(11.5), Inches(0.45), "Player-coach plan for a high-performing EMEA SDR org", size=17, color=MUTED, font=FONT)
    add_rect(s, 0, Inches(5.85), SLIDE_W, Inches(1.65), INK)
    add_textbox(s, Inches(0.85), Inches(6.15), Inches(11.5), Inches(0.4), "Lindsey Dempsey  ·  EMEA SDR Leader Candidate", size=18, bold=False, color=WHITE, font=FONT)
    add_textbox(s, Inches(0.85), Inches(6.6), Inches(11.5), Inches(0.35), "Cursor wins on workflow · context · execution", size=14, color=SOFT, font=FONT)
    notes(s, note)
    return s


def section_slide(prs, part, title, note):
    s = blank(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, PARCHMENT)
    add_logo(s, light=False, left=Inches(0.85), top=Inches(0.55), width=Inches(1.8))
    add_accent(s, Inches(0.85), Inches(2.55), width=Inches(1.35))
    add_textbox(s, Inches(0.85), Inches(2.8), Inches(11), Inches(0.35), part.upper(), size=13, bold=True, color=EMBER, font=FONT)
    add_textbox(s, Inches(0.85), Inches(3.3), Inches(11.6), Inches(1.2), title, size=30, bold=False, color=INK, font=FONT)
    notes(s, note)
    return s


def content_slide(prs, title, bullets, note, *, size=17):
    s = blank(prs)
    header(s, title)
    add_bullets(s, Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.1), bullets, size=size, space_after=11)
    notes(s, note)
    return s


def two_col(prs, title, left_title, left_items, right_title, right_items, note, *, size=14):
    s = blank(prs)
    header(s, title)
    add_rect(s, Inches(0.6), Inches(1.4), Inches(5.9), Inches(5.25), WHITE)
    add_accent(s, Inches(0.9), Inches(1.65), width=Inches(0.8))
    add_textbox(s, Inches(0.9), Inches(1.9), Inches(5.3), Inches(0.35), left_title, size=16, bold=True, color=CHARCOAL, font=FONT)
    add_bullets(s, Inches(0.9), Inches(2.4), Inches(5.3), Inches(3.9), left_items, size=size, space_after=7)

    add_rect(s, Inches(6.75), Inches(1.4), Inches(5.9), Inches(5.25), WHITE)
    add_accent(s, Inches(7.05), Inches(1.65), width=Inches(0.8))
    add_textbox(s, Inches(7.05), Inches(1.9), Inches(5.3), Inches(0.35), right_title, size=16, bold=True, color=CHARCOAL, font=FONT)
    add_bullets(s, Inches(7.05), Inches(2.4), Inches(5.3), Inches(3.9), right_items, size=size, space_after=7)
    notes(s, note)
    return s


def market_slide(prs, note):
    """History: two product philosophies + two-layer race + competitive shapes."""
    s = blank(prs)
    header(s, "Competitor differentiation — market map")

    # Philosophies strip
    add_rect(s, Inches(0.6), Inches(1.35), Inches(5.9), Inches(1.55), WHITE)
    add_textbox(s, Inches(0.8), Inches(1.45), Inches(5.5), Inches(0.3), "AI-native IDE", size=14, bold=True, color=CHARCOAL, font=FONT)
    add_textbox(s, Inches(0.8), Inches(1.8), Inches(5.5), Inches(0.9), "Cursor · Windsurf\nRebuild the editor around agents — not bolt AI onto an existing one.", size=13, color=MUTED)

    add_rect(s, Inches(6.75), Inches(1.35), Inches(5.9), Inches(1.55), WHITE)
    add_textbox(s, Inches(6.95), Inches(1.45), Inches(5.5), Inches(0.3), "IDE plugin / extension", size=14, bold=True, color=CHARCOAL, font=FONT)
    add_textbox(s, Inches(6.95), Inches(1.8), Inches(5.5), Inches(0.9), "GitHub Copilot · JetBrains AI · Amazon Q\nKeep the current toolchain; add AI as a layer.", size=13, color=MUTED)

    # Two layers
    add_rect(s, Inches(0.6), Inches(3.1), Inches(5.9), Inches(2.35), WHITE)
    add_rect(s, Inches(0.6), Inches(3.1), Inches(5.9), Inches(0.42), CHARCOAL)
    add_textbox(s, Inches(0.8), Inches(3.15), Inches(5.5), Inches(0.35), "Layer 1 — Harness / product", size=13, bold=True, color=CITRUS)
    add_bullets(s, Inches(0.8), Inches(3.7), Inches(5.5), Inches(1.55), [
        "Plans, tools, repo navigation, workflow fit",
        "What buyers feel in first conversations",
        "Cursor’s established commercial strength",
    ], size=13, space_after=5)

    add_rect(s, Inches(6.75), Inches(3.1), Inches(5.9), Inches(2.35), WHITE)
    add_rect(s, Inches(6.75), Inches(3.1), Inches(5.9), Inches(0.42), SLATE)
    add_textbox(s, Inches(6.95), Inches(3.15), Inches(5.5), Inches(0.35), "Layer 2 — Model", size=13, bold=True, color=WHITE)
    add_bullets(s, Inches(6.95), Inches(3.7), Inches(5.5), Inches(1.55), [
        "Reasoning + long-horizon coding quality",
        "Composer = Cursor coding-specialist path",
        "Labs can be suppliers and competitors",
    ], size=13, space_after=5)

    add_rect(s, Inches(0.6), Inches(5.65), Inches(12.05), Inches(1.15), CHARCOAL)
    add_textbox(
        s,
        Inches(0.8),
        Inches(5.8),
        Inches(11.6),
        Inches(0.85),
        "Shapes:  Cursor = platform  ·  Copilot = plugin / GitHub ecosystem  ·  Claude Code = terminal agent  ·  Windsurf = Cascade IDE challenger  ·  Codex = OpenAI async agents  ·  Often the real rival = ungoverned shadow AI",
        size=13,
        color=SOFT,
    )
    notes(s, note)
    return s


def positioning_slide(prs, note):
    s = blank(prs)
    header(s, "Cursor positioning — workflow · context · execution")

    add_rect(s, Inches(0.6), Inches(1.35), Inches(12.05), Inches(0.85), CHARCOAL)
    add_textbox(
        s,
        Inches(0.8),
        Inches(1.45),
        Inches(11.6),
        Inches(0.65),
        "Leading agentic coding platform — a coding agent across IDE, CLI, Automations, and Cloud Agents with a shared AI layer on the codebase. Helps teams get more from frontier models via better workflow, context, and integrated execution.",
        size=13,
        color=SOFT,
    )

    # Surfaces
    for i, (label, sub) in enumerate([("IDE", "Desktop agentic"), ("CLI", "Terminal / scripts"), ("Automations", "Event-driven"), ("Cloud Agents", "Async / mobile")]):
        x = Inches(0.6) + i * Inches(3.1)
        add_rect(s, x, Inches(2.4), Inches(2.95), Inches(1.1), WHITE)
        add_accent(s, x + Inches(0.18), Inches(2.55), width=Inches(0.6))
        add_textbox(s, x + Inches(0.18), Inches(2.75), Inches(2.55), Inches(0.28), label, size=14, bold=True, color=CHARCOAL, font=FONT)
        add_textbox(s, x + Inches(0.18), Inches(3.1), Inches(2.55), Inches(0.25), sub, size=12, color=MUTED)

    # Enterprise buckets
    buckets = [
        ("Productivity", "Multi-file agents, codebase context, task completion — not just autocomplete"),
        ("Governance", "SSO · SCIM · RBAC · audit logs · admin controls over models / repos / MCP"),
        ("Privacy & security", "Privacy Mode / ZDR · SOC 2 Type II · CMEK · reduce shadow AI risk"),
    ]
    for i, (t, b) in enumerate(buckets):
        x = Inches(0.6) + i * Inches(4.15)
        add_rect(s, x, Inches(3.7), Inches(4.0), Inches(1.55), WHITE)
        add_textbox(s, x + Inches(0.2), Inches(3.85), Inches(3.55), Inches(0.3), t, size=14, bold=True, color=CHARCOAL, font=FONT)
        add_textbox(s, x + Inches(0.2), Inches(4.25), Inches(3.55), Inches(0.8), b, size=12, color=MUTED)

    add_textbox(
        s,
        Inches(0.8),
        Inches(5.5),
        Inches(11.6),
        Inches(1.2),
        "Harness = the engineering layer between developer and model (context packing, tool use, turn management, routing).\nProof: Coinbase (agent-first) · Stripe (standardized rollout) · PayPal (enterprise velocity) — outcomes, not vanity AI metrics.",
        size=13,
        color=MUTED,
    )
    notes(s, note)
    return s


def differentiators_slide(prs, note):
    s = blank(prs)
    header(s, "Four differentiators — lead with one based on buyer pain")
    cards = [
        ("01", "Model neutrality", "Best model changes. No single-vendor lock-in. Labs are suppliers and competitors — optionality is strategic."),
        ("02", "Large codebase / harness", "Not just wrapping a model. Better finding + using context in messy enterprise repos. Buyers feel this day one."),
        ("03", "Faster time to value", "Strong out of the box. Standardize without every developer inventing a custom setup. Adoption + consistency."),
        ("04", "Platform across the SDLC", "Bugbot, Agent Review, Automations. Clear the review bottleneck AI creates. Better execution, not just answers."),
    ]
    positions = [(Inches(0.6), Inches(1.4)), (Inches(6.75), Inches(1.4)), (Inches(0.6), Inches(4.0)), (Inches(6.75), Inches(4.0))]
    for (left, top), (num, title, body) in zip(positions, cards):
        add_rect(s, left, top, Inches(5.9), Inches(2.35), WHITE)
        add_accent(s, left + Inches(0.25), top + Inches(0.25), width=Inches(0.7))
        add_textbox(s, left + Inches(0.25), top + Inches(0.45), Inches(0.85), Inches(0.3), num, size=15, bold=True, color=CHARCOAL, font=FONT)
        add_textbox(s, left + Inches(1.1), top + Inches(0.45), Inches(4.4), Inches(0.3), title, size=15, bold=True, color=CHARCOAL, font=FONT)
        add_textbox(s, left + Inches(0.25), top + Inches(1.0), Inches(5.35), Inches(1.05), body, size=13, color=MUTED)
    notes(s, note)
    return s


def gains_slide(prs, note):
    """History: quality compounding + bottleneck / software factory."""
    s = blank(prs)
    header(s, "Where gains live — quality + the review/production bottleneck")

    add_bullets(s, Inches(0.8), Inches(1.45), Inches(6.3), Inches(3.8), [
        "Conversation quality — relevance speed + correct wedge",
        "Qualification bar — fewer junk meetings, higher AE acceptance",
        "Competitive fluency — Copilot / Claude Code / Windsurf frames",
        "Earlier multi-threading — Platform + Security on enterprise accounts",
    ], size=16, space_after=12)

    add_rect(s, Inches(7.3), Inches(1.45), Inches(5.3), Inches(4.5), WHITE)
    add_accent(s, Inches(7.55), Inches(1.7), width=Inches(0.8))
    add_textbox(s, Inches(7.55), Inches(1.95), Inches(4.8), Inches(0.35), "Software factory insight", size=15, bold=True, color=CHARCOAL, font=FONT)
    add_textbox(
        s,
        Inches(7.55),
        Inches(2.45),
        Inches(4.8),
        Inches(3.2),
        "AI tools speed writing → bottleneck shifts to review & production.\n\nCursor helps by automating across the lifecycle — Bugbot / Agent Review, background agents, Composer efficiency — so teams ship outcomes, not just more PRs.\n\nAssisted → Agentic → Cloud agents → Governed factory.",
        size=13,
        color=MUTED,
    )
    notes(s, note)
    return s


def competitive_battlecard(prs, note):
    s = blank(prs)
    header(s, "Why we win — competition differentiation battlecard")

    add_rect(s, Inches(0.5), Inches(1.3), Inches(12.25), Inches(0.7), CHARCOAL)
    add_textbox(
        s,
        Inches(0.7),
        Inches(1.4),
        Inches(11.8),
        Inches(0.5),
        "Not exclusive smarter intelligence forever — apply intelligence better: workflow + codebase context + SDLC integration + model neutrality.",
        size=12,
        color=SOFT,
    )

    cols = [
        (Inches(0.5), "Cursor", CITRUS, CHARCOAL, [
            "AI-native IDE + enterprise platform",
            "Full-repo context + agent-first",
            "IDE · CLI · Automations · Cloud Agents",
            "Multi-model + Composer path",
            "Bugbot clears PR bottleneck",
            "Governance: SSO / SCIM / audit",
        ]),
        (Inches(3.65), "vs Copilot", WHITE, SLATE, [
            "Biggest by market share",
            "Plugin in existing IDEs",
            "Wins: GitHub ecosystem + friction",
            "Often local / snippet help",
            "Our move: task completion",
            "Our move: codebase workflow",
        ]),
        (Inches(6.8), "vs Claude Code", WHITE, SLATE, [
            "Strong — don’t trash-talk",
            "Terminal-first agent shape",
            "Not a model-war pitch",
            "Our edge: neutrality",
            "Our edge: SDLC integration",
            "Our edge: team time-to-value",
        ]),
        (Inches(9.95), "Also / real rival", WHITE, SLATE, [
            "Windsurf: Cascade UX challenger",
            "Codex: OpenAI async agents",
            "Shadow AI = real enterprise rival",
            "Win = governed standardization",
            "Proof: Coinbase · Stripe · PayPal",
            "Outcomes per $ > autocomplete",
        ]),
    ]
    for left, title, tcolor, bar, items in cols:
        add_rect(s, left, Inches(2.2), Inches(3.0), Inches(4.4), WHITE)
        add_rect(s, left, Inches(2.2), Inches(3.0), Inches(0.48), bar)
        add_textbox(s, left + Inches(0.12), Inches(2.28), Inches(2.75), Inches(0.35), title, size=12, bold=True, color=tcolor if bar != WHITE else CHARCOAL, font=FONT)
        add_bullets(s, left + Inches(0.12), Inches(2.9), Inches(2.75), Inches(3.4), items, size=11, space_after=5)
    notes(s, note)
    return s


def numbered_slide(prs, title, items, note):
    s = blank(prs)
    header(s, title)
    top = Inches(1.4)
    for num, heading, body in items:
        add_textbox(s, Inches(0.8), top, Inches(0.65), Inches(0.32), num, size=16, bold=True, color=CHARCOAL, font=FONT)
        add_textbox(s, Inches(1.5), top, Inches(10.9), Inches(0.3), heading, size=15, bold=True, color=CHARCOAL)
        add_textbox(s, Inches(1.5), top + Inches(0.32), Inches(10.9), Inches(0.5), body, size=13, color=MUTED)
        top += Inches(1.15)
    notes(s, note)
    return s


def calendar_slide(prs, note):
    """15 / 30-day operating calendar from Lindsey's Part 1 clarifications."""
    s = blank(prs)
    header(s, "Operating calendar — first 15 and 30 days")

    # Day 15 column
    add_rect(s, Inches(0.55), Inches(1.4), Inches(6.0), Inches(5.3), WHITE)
    add_rect(s, Inches(0.55), Inches(1.4), Inches(6.0), Inches(0.7), CHARCOAL)
    add_textbox(s, Inches(0.75), Inches(1.5), Inches(5.5), Inches(0.25), "DAYS 1–15", size=12, bold=True, color=CITRUS)
    add_textbox(s, Inches(0.75), Inches(1.78), Inches(5.5), Inches(0.25), "Inspection · Data · Training design", size=16, bold=True, color=WHITE, font=FONT)
    add_bullets(s, Inches(0.8), Inches(2.35), Inches(5.4), Inches(4.0), [
        "Inspect the machine — no big changes yet",
        "Pull hard data: qualified vs unqualified meetings",
        "AE interviews: titles, pain, industry, competitor, next steps",
        "Funnel: connect → conversation → meeting",
        "Channel mix: email vs LinkedIn vs phone",
        "Spot patterns → design coaching & training",
        "Build objection, discovery, and differentiator modules",
    ], size=14, space_after=8)

    # Day 30 column
    add_rect(s, Inches(6.75), Inches(1.4), Inches(6.0), Inches(5.3), WHITE)
    add_rect(s, Inches(6.75), Inches(1.4), Inches(6.0), Inches(0.7), SLATE)
    add_textbox(s, Inches(6.95), Inches(1.5), Inches(5.5), Inches(0.25), "DAYS 16–30", size=12, bold=True, color=CITRUS)
    add_textbox(s, Inches(6.95), Inches(1.78), Inches(5.5), Inches(0.25), "Train · Activity · Daily cadence", size=16, bold=True, color=WHITE, font=FONT)
    add_bullets(s, Inches(7.0), Inches(2.35), Inches(5.4), Inches(4.0), [
        "Run the training (openers, objections, discovery)",
        "Increase activity metrics once diagnosis is clear",
        "Daily data pull — what worked yesterday?",
        "Daily team call — first 15 minutes of the day",
        "Review: pitch landing · personas · differentiator that won",
        "One customer story reviewed every day",
        "Models that leapfrogged overnight → update talk tracks",
        "Game-plan the day for success before dials start",
    ], size=13, space_after=6)
    notes(s, note)
    return s


def diagnosis_slide(prs, note):
    """Lindsey's diagnosis framework."""
    s = blank(prs)
    header(s, "Diagnosis first — learn before you change anything")

    cols = [
        (Inches(0.5), "AE acceptance truth", [
            "% meetings → qualified vs unqualified",
            "Unqualified = wasted AE time",
            "Talk to AEs: where did it fall flat?",
            "Pattern hunt: titles · pain · industry",
            "Competitor? Why no next steps?",
            "Gaps → coaching plan",
        ]),
        (Inches(4.6), "Funnel leaks", [
            "Connect rate",
            "Connect → conversation",
            "Conversation → meeting",
            "Low C→Conv: openers / objections → objection training",
            "Low Conv→Mtg: discovery / pain → differentiator or competitor training",
            "Already using a competitor? Train the frame",
        ]),
        (Inches(8.7), "Channel mix", [
            "% meetings via email",
            "% meetings via LinkedIn",
            "% meetings via phone",
            "Which is strongest?",
            "Which is weakest?",
            "Double down on what’s working — lift the weak bucket",
        ]),
    ]
    for left, title, items in cols:
        add_rect(s, left, Inches(1.4), Inches(3.95), Inches(5.25), WHITE)
        add_accent(s, left + Inches(0.2), Inches(1.65), width=Inches(0.75))
        add_textbox(s, left + Inches(0.2), Inches(1.9), Inches(3.5), Inches(0.4), title, size=15, bold=True, color=CHARCOAL, font=FONT)
        add_bullets(s, left + Inches(0.2), Inches(2.45), Inches(3.5), Inches(3.9), items, size=13, space_after=7)
    notes(s, note)
    return s


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    N1 = """Hi everyone — I’m Lindsey Dempsey. Thank you for having me.

Today I’m going to walk through how I would up-level the EMEA SDR machine — not rebuild it. My framing is simple: diagnose, protect, experiment, and execute, as a hands-on player-coach.

I’ll keep this practical and leave room for discussion and a live pitch."""

    N2 = """My thesis is this: EMEA is already high-performing. You have AI-augmented workflows, a clear ICP, and you’re generating eight-figure qualified pipeline every month. So the question isn’t “how do I stand something up from zero?” It’s “given a function that already works, how do I make it materially better?”

My job would be to find the real constraints, protect what’s working, and run a small number of high-leverage experiments.

Every SDR on the team should be fluent in how Cursor wins — on workflow, context, and execution.

And I would lead as a player-coach. Trains on time. Inspection over activity theater. Success for me is better conversation quality and higher AE acceptance — not reinventing the org chart."""

    N3 = """Here’s how I’ll use the time.

Part one is the first fifteen, thirty, and sixty days — diagnosis, what I’d protect, where the gains are, and two experiments.

Part two is fundamentals — how I hire, coach, and run the team day to day as a player-coach.

Part three is lighter on slides and partly live — about ten minutes. I’ll show how I’d earn a first meeting, how that changes by persona, how I cut through noise, and why I believe Cursor wins. I’m drawing from the whiteboard motion, not a rehearsed script.

Clarity over polish. If you want me to go live earlier on the pitch, I’m happy to."""

    N5 = """Diagnosis comes first. I will not change the machine until I understand where it leaks.

I’d start with meeting quality: what percentage of meetings move to qualified versus unqualified? Unqualified meetings waste AE time. I’d sit with the AE team and ask where those meetings fell flat — is there a pattern? Is it the titles? No pain? The industry? A competitor already in? Why couldn’t we get next steps? Once I see the pattern in the gaps, I can diagnose what’s missing and build coaching around it.

I’d also inspect the funnel hard: connect rates, connect-to-conversation, conversation-to-meeting. If connect-to-conversation is low — how are we opening the call, and what objections are we getting? If I know the objections, I can lead objection trainings. If conversation-to-meeting is low — are we doing proper discovery to create enough pain that they want the meeting? Are we hearing “already using a competitor”? If so, that’s competitor training or Cursor differentiator training.

And I’d look at channel mix: what percentage of meetings are set via email, LinkedIn, and phone? Which is weakest, which is strongest? How do we increase one bucket while keeping the others strong — and double down on what’s already working?"""

    N5B = """Here’s how that diagnosis turns into a calendar.

Days one through fifteen are inspection, pulling data, and creating the training. No thrash — learn the patterns first, then design the coaching modules for openers, objections, discovery, and differentiators.

Days sixteen through thirty are when we take the training live and increase activity. Every day we pull the data to see what worked. The first fifteen minutes of the day is a team call: what worked yesterday, what pitch is landing, which personas sparked interest, what differentiator secured the meeting. We review one customer story a day. We review which models leapfrogged overnight so talk tracks stay current. Then we game-plan the day for success before dials start."""

    N6 = """What I would deliberately not change in the first sixty days is what’s already working well.

I would protect the team’s ability to know and deliver the Cursor pitch. I would protect reaching out to the right people — the ICP and persona discipline that’s already creating meetings. I’m not here to rip out a motion that works. I’m here to diagnose the leaks and raise conversion quality around it."""

    N7 = """Given where EMEA already is, the next level of improvement doesn’t come from inventing a new org. It comes from tightening the places strong teams usually leak.

After diagnosis, the gains are in call openers and objection handling, discovery that creates real pain, competitive and differentiator fluency when Copilot or Claude Code is already in, and channel mix — doubling down on what’s converting while lifting the weaker channel.

I wouldn’t bet on that blindly. I’d test it with hard data: connect, conversation, meetings — and then how many of those meetings convert to qualified. What worked, what didn’t, keep or kill."""

    N8 = """Quick market context, because it shapes how we coach messaging.

Agentic coding is now a two-layer race.

Layer one is the harness and product layer — how the system plans, uses tools, navigates repositories, and fits developer workflow. That’s where Cursor has already built commercial strength, and it’s what buyers feel in a first conversation.

Layer two is the model layer — reasoning quality and long-horizon coding performance. Composer is Cursor’s coding-specialist model path, and that layer continues to strengthen.

Model neutrality still matters, because the best model changes over time, and the frontier labs can be both suppliers and competitors.

For EMEA SDRs, the implication is simple: sell the system buyers can feel today — workflow, context, and execution — and stay fluent on where the model layer is going. I won’t turn customer calls into infrastructure or acquisition conversations."""

    N9 = """This is the spine I’d train every rep on.

Cursor is the leading agentic coding platform that a large share of the Fortune one thousand use to reimagine how engineering teams build software with AI.

It gives teams a shared AI layer that understands their codebase, dependencies, and workflows — connecting developers, tools, and AI agents into one system.

It’s helpful to think of Cursor as a coding agent that can be accessed, coordinated, and controlled through product surfaces like the IDE, CLI, Automations, and Cloud Agents.

It bridges planning, execution, and validation so work moves through the SDLC with less friction.

The summary line I want reps to be able to say is: Cursor helps teams get more out of frontier models by pairing them with better workflow, better context, and a more integrated way to build software."""

    N10 = """Cursor’s core differentiation is the combination of four things, and I’d teach SDRs to lead with only one based on the buyer’s pain.

One — model neutrality. Customers want flexibility. The best model changes. They don’t want to be locked to one vendor.

Two — large codebase performance. Enterprise work spans messy, real-world repos. Cursor isn’t just wrapping a model — it improves how context is found and used. That matters more as complexity increases.

Three — faster time to value. Cursor is useful quickly without every developer inventing a custom setup. Enterprises care about adoption and consistency across teams.

Four — platform across the SDLC. Not just code generation — planning, writing, reviewing, debugging, iterating — through things like Bugbot, Agent Review, and Automations. The value is better execution, not just better answers.

If a rep tries to say all four on every call, they’ll lose the room. Diagnose first, then pick the wedge."""

    N11 = """The one or two changes I’d test first are practical and measured.

First: once we’ve diagnosed the problem with call openers, objection handling, and discovery — increase activity. Don’t crank dials into a broken talk track. Fix the skill, then raise the volume, and watch connect, conversation, and meetings move.

Second: improve messaging through personalization and staying up to date with AI models and labs constantly releasing new improvements. That means daily awareness of what leapfrogged overnight and talk tracks that stay current.

I’d measure both with hard data — connect, conversation, meetings — and then ask the quality question: how many of these meetings are now converting to qualified? What worked? What did not work?"""

    N12 = """The feedback loops with Growth and Sales already exist — so the question isn’t inventing them, it’s making them quicker and implementing change faster.

I’d run weekly SE sessions so SDRs stay sharp on product and competitive framing. I’d make SDR and AE connects mandatory — not optional — so meeting quality feedback is real-time, not buried. And I’d send weekly recaps to AEs, SVPs, RDs, and SEs by territory: what converted, what failed, and one change we’re making.

Faster loops, clearer owners, trains on time."""

    N14 = """The role has changed because AI can take the rote work — research, first drafts, list building — so SDRs can be as efficient as possible and focus on what humans are actually good at.

Humans still want to be sold to by humans. Even with AI, messaging and cold calls need humanization. The way we stand out isn’t louder sequences — it’s talking to prospects from a place of help: start a real discussion so they can reap the benefits of Cursor.

Exceptional is someone who uses AI for the rote work, then stands out through genuine discovery, care, understanding the current state and the pains underneath it, and knowing which customer story matches the use case. They need organization and time management, research and business intuition, written and verbal communication, active listening and adaptability — carry the conversation and react with value — objection handling, and storytelling that paints technical pain and business value."""

    N15 = """When I hire, I’m looking for someone who will be excited every single day — and someone I can see moving into the next role.

My bar is usually one to two years in sales, or something technical that required intelligence. Some of my best hires have been teachers or people with real work experience who want into sales — they understand concepts quickly.

Every candidate needs four traits plus one non-negotiable: intelligence so they can gain knowledge, drive so they can gain the skills, coachable so they adapt, integrity so they’re honest — and a genuine passion for AI and Cursor. It’s not enough to want a fast-growing company. I want people who nerd out on which model is best and what’s changing in the space.

Process-wise I source through LinkedIn Recruiter, and I prefer native language for the role — I’ve seen non-native speakers struggle when they have to sell technically in a second language."""

    N15B = """Coaching and development all follow the same framework: Tell — what, who, why. Show — what good looks like. Observe — let them struggle. Provide feedback — but first ask what they did well and what they’d do differently. Repeat until it’s mastered.

For new reps I spend a lot of time with them — constant training, homework, fix gaps quickly, and always the why. I give a ninety-day ramp broken down by week so they know exactly what to do when the week starts, and a buddy on the team for questions and advice.

For top people I learn their motivation early, catch them doing something good, let them mentor, and keep coaching the gaps in their workflow. For underperformance I put a coaching plan in place as soon as I see someone off track — structure before it becomes a bigger issue.

Culture is scoreboards, competitions, collaboration, and ending every week with a shoutout for someone else on the team so collaboration doesn’t go unnoticed."""

    N17 = """This section is live — I’m happiest going here without slides doing the talking. Same craft as the whiteboard session with Derek and Dave at J&J.

If I’m earning a first meeting with one of you, I’m listening and diagnosing, not pitching a deck. Permission first — not slide-ware. I’d open with: what prompted you to take the call — or, the reason I’m calling is teams are writing code faster with AI and then hitting a bottleneck getting work through review and into production.

I’d ask where engineering time actually goes — keep-the-lights-on, refactor, innovate, moonshot — because that’s where capacity for the projects that matter is hiding. The software factory frame: ideas in, valuable outputs out. If Copilot, Claude Code, or Codex is already in, I don’t trash-talk. I meet them on the maturity curve: assisted, agentic, cloud agents, governed factory — and ask where the team is stuck.

Brush-off: if I get “just send an email,” I don’t collapse. Happy to — so I send the useful one: are you more focused on standardizing AI coding, or on the review bottleneck after Copilot?

Ask: worth twenty minutes with you — and if useful, your platform lead — to map where AI coding helps versus where workflow still breaks. In a working session I’d close the way the whiteboard did: what’s the most important thing you saw today — for you personally, and for the company.

I’m ready to do this live on one of you right now."""

    N18 = """Inboxes and LinkedIn are saturated, so I don’t believe volume tricks break through with this audience. Breaking through is persona-true relevance — one sharp question tied to a real trigger.

For a hands-on developer: lead with real-repo context and daily workflow. Avoid corporate ROI theater. Peer proof and “try this on your actual pain” beat feature lists.

For a VP of Engineering: lead with standardization, shipping speed, and capacity — the whiteboard motion of freeing time from keep-the-lights-on into moonshots. Avoid feature laundry lists.

For Security or Platform: lead with governed adoption — Privacy Mode, admin controls, SSO/SCIM/audit — so AI coding doesn’t become shadow IT. Avoid hype. Earn the right to go deep later.

What still works: trigger-based relevance — Copilot sprawl, AI policy moments, migrations, DevEx ownership — plus a binary question that earns the next conversation. What fails: “quick chat” spam, fake personalization, and slide-ware before you’ve earned diagnosis."""

    N19 = """Getting attention right now is about earning thirty seconds with a technical buyer, not winning the inbox lottery.

What works: a reason for the call tied to their world; diagnosing where time and bottlenecks actually sit; meeting them on the AI maturity curve; using customer stories that match the use case — Coinbase for agent-first, Stripe for rollout, PayPal for enterprise velocity, plus whiteboard stories like Nike RFID, Amplitude background agents, T-Mobile at scale; and asking for a working session, not a vague chat. Token efficiency and outcomes per dollar matter when finance is in the room — that’s attention currency too.

What used to work that I’m confident will not work with Cursor’s audience: generic sequences, “just checking in,” ROI theater for developers, and feature dumps. Technical buyers smell fake personalization instantly, and they punish anyone who tries to out-engineer them on a cold call. Curiosity and diagnosis beat polish."""

    N20 = """My own read — not the company line — is this:

Cursor wins because the market moved from “help me write code” to “help my org ship software with agents without losing control.” AI coding tools create speed and then move the bottleneck to review and production. Cursor wins when it becomes the governed software factory — workflow, codebase context, model neutrality, and execution across the SDLC — so teams free capacity for the work that actually changes the business.

To a skeptical engineering leader in two sentences: Cursor doesn’t win by claiming the permanently smarter model. We win by applying intelligence better across real engineering workflows — and giving the enterprise the control plane to standardize that without shadow AI or wasted tokens."""

    N21 = """I’ll close where I started.

In the first sixty days I won’t invent a new EMEA. I’ll diagnose the real constraints with data and AE truth — especially qualified versus unqualified meetings. I’ll protect the pitch and the right-person discipline that’s already working. I’ll train and raise activity only after we fix openers, objections, and discovery. I’ll run faster loops with Growth, AEs, and SEs. And I’ll lead as a player-coach — Tell, Show, Observe, Feedback — so the floor rises with the ceiling.

That’s how you up-level a function that’s already working — and keep the trains on time.

I’m Lindsey Dempsey. I’d love your questions — and if useful, I’m ready to pitch one of you live right now."""

    # 1
    title_slide(prs, N1)

    # 2
    content_slide(prs, "Thesis: don’t rebuild a working machine — raise its ceiling", [
        "EMEA already high-performing: AI-augmented workflows, clear ICP, eight-figure monthly pipeline",
        "Job = find real constraints, protect what works, run a few high-leverage experiments",
        "Every SDR fluent in how Cursor wins — workflow, context, execution",
        "Lead as a player-coach: trains on time, inspection over activity theater",
        "Success = better conversation quality + higher AE acceptance — not a reorg",
    ], N2)

    # 3
    content_slide(prs, "Agenda", [
        "Part 1 (20 min) — Diagnosis, 15/30 calendar, protect, gains, experiments, Growth/Sales loops",
        "Part 2 (12 min) — AI-era SDR, hiring bar, Tell→Show→Observe coaching",
        "Part 3 (~10 min, partly live) — Pitch, personas, attention, why we win",
        "Clarity over polish — happy to go live earlier on the pitch",
    ], N3)

    # 4
    section_slide(prs, "Part 1", "How I’d up-level EMEA in the first sixty days", "Part one — how I’d up-level EMEA in the first sixty days.")

    # 5 — Lindsey diagnosis
    diagnosis_slide(prs, N5)

    # 6 — Calendar 15 / 30
    calendar_slide(prs, N5B)

    # 7 — Protect
    content_slide(prs, "What I’d protect in the first 60 days", [
        "What’s already working well — don’t rip out a motion that produces pipeline",
        "Knowing the Cursor pitch — keep that fluency intact",
        "Reaching out to the right people — ICP and persona discipline",
        "I’m here to diagnose leaks and raise conversion quality — not reinvent the org chart",
    ], N6)

    # 8 — Gains
    content_slide(prs, "Where the next level of gains come from", [
        "Call openers + objection handling (after we know the real objections)",
        "Discovery that creates enough pain to earn the meeting",
        "Competitor / differentiator fluency when Copilot or Claude Code is already in",
        "Channel mix — double down on the strongest; lift the weakest",
        "Test before betting: connect · conversation · meetings · % that convert to qualified",
    ], N7)

    # 9 — competitor map from history
    market_slide(prs, N8)

    # 10 — positioning
    positioning_slide(prs, N9)

    # 11 — differentiators
    differentiators_slide(prs, N10)

    # 12 — experiments (Lindsey)
    two_col(prs, "Two experiments — test first, measure hard", "1 — Skill then activity", [
        "Diagnose openers, objection handling, and discovery first",
        "Then increase activity metrics — don’t crank dials into a broken talk track",
        "Measure: connect · conversation · meetings",
        "Quality check: % converting to qualified meetings",
    ], "2 — Messaging currency", [
        "Personalization that earns the next 30 seconds",
        "Stay current as AI models and labs release improvements",
        "Daily: which models leapfrogged → update talk tracks",
        "What worked / what didn’t — keep or kill on data",
    ], N11)

    # 13 — Growth & Sales loops (Lindsey)
    content_slide(prs, "Working with Growth & Sales — faster feedback loops", [
        "Loops already exist — make them quicker and implement change faster",
        "Weekly SE sessions — product + competitive sharpness",
        "Mandatory SDR ↔ AE connects — real-time meeting quality feedback",
        "Weekly recaps to AEs, SVPs, RDs, and SEs by territory",
        "What converted · what failed · one change we’re making",
        "Clear owners. Faster loops. Trains on time.",
    ], N12)

    # Part 2
    section_slide(prs, "Part 2", "Fundamentals — how I lead day to day as a player-coach", "Part two — fundamentals. How I lead day to day.")

    content_slide(prs, "The ideal AI-era SDR", [
        "AI takes the rote work — humans still want to be sold to by humans",
        "Humanize messaging and cold calls — stand out by helping, not spamming",
        "Exceptional: AI for prep + genuine discovery, care, current-state pain, use-case stories",
        "Skills: org/time management · research/intuition · written & verbal",
        "Active listening + adaptability — carry the conversation and react with value",
        "Objection handling + storytelling — technical pain and business value",
    ], N14, size=16)

    two_col(prs, "Hiring great talent", "Bar", [
        "Excited every day — can I see them in the next role?",
        "1–2 years sales or technical / high-intelligence background",
        "Teachers & career-switchers often win — they learn concepts fast",
        "LinkedIn Recruiter · prefer native language for the role",
    ], "Must-haves", [
        "Intelligence — can gain knowledge",
        "Drive — will gain the skills",
        "Coachable — adapt to change",
        "Integrity — honest",
        "Obsessed with AI + Cursor — not just “hot company”",
    ], N15, size=14)

    content_slide(prs, "Coaching system — Tell → Show → Observe → Feedback", [
        "Tell (what / who / why) → Show good → Observe (let them struggle) → Feedback (ask them first) → Repeat",
        "Ramp: 90-day plan by week + homework + buddy system + heavy on the why",
        "Top performers: learn motivation early · catch them doing good · make them mentors",
        "Underperformance: coaching plan as soon as someone is off track — structure early",
        "Culture: scoreboards · competitions · collaboration · Friday shoutouts for someone else",
    ], N15B, size=15)

    # Part 3 — lighter, whiteboard-informed, live-ready
    section_slide(
        prs,
        "Part 3",
        "Craft & curiosity — lighter slides, mostly live",
        "Part three — craft and curiosity. This is where I’m happiest going live.",
    )

    numbered_slide(prs, "Live pitch — earn the first meeting (whiteboard craft)", [
        ("01", "Open", "Permission, not slide-ware. “What prompted the call?” — or: writing faster → stuck before production."),
        ("02", "Diagnose", "Time split: KTLO / refactor / innovate / moonshot. Maturity: assisted → agentic → cloud → governed factory."),
        ("03", "Position", "Software factory: ideas in → valuable outputs out. Competitor in? Don’t trash-talk — meet them on the curve."),
        ("04", "Ask", "20 min working session. Brush-off → binary useful email. Close: “Most important thing you saw today?”"),
    ], N17)

    content_slide(prs, "Cutting through noise — by persona", [
        "Developer: real-repo context + daily workflow · peer proof · no ROI theater",
        "VP Engineering: standardize AI coding · free KTLO capacity for moonshots · shipping outcomes",
        "Security / Platform: governed adoption · Privacy Mode · SSO/SCIM/audit · kill shadow AI",
        "Triggers that earn 30 seconds: Copilot sprawl · AI policy · migrations · DevEx ownership",
        "Fail: “quick chat,” fake personalization, feature dumps, slide-ware before diagnosis",
    ], N18, size=15)

    two_col(prs, "Getting attention — what works / what dies", "Works now", [
        "Reason for the call tied to their world",
        "Diagnose time split + bottleneck shift",
        "Meet them on the maturity curve",
        "Use-case-true stories (Coinbase / Stripe / PayPal · Nike / Amplitude)",
        "Ask for a working session, not a “quick chat”",
        "Outcomes per dollar when finance pressure shows up",
    ], "Dies with this audience", [
        "Generic sequences / “just checking in”",
        "ROI theater aimed at developers",
        "Feature laundry lists",
        "Fake personalization",
        "Slide-ware before you’ve earned diagnosis",
        "Trying to out-engineer the room on a cold call",
    ], N19, size=13)

    content_slide(prs, "Why we win — my read for a skeptical eng leader", [
        "Market moved from “help me write code” → “help my org ship with agents — without losing control”",
        "AI creates speed, then moves the bottleneck to review and production",
        "Cursor wins as the governed software factory: workflow · context · neutrality · SDLC execution",
        "One-liner: we don’t win by claiming the permanently smarter model — we win by applying intelligence better, with the control plane to standardize it",
    ], N20, size=16)

    content_slide(prs, "Close — up-level, don’t rebuild", [
        "Diagnose qualified vs unqualified + funnel leaks with AE truth",
        "Protect the Cursor pitch and right-person discipline already working",
        "Train openers / objections / discovery — then raise activity",
        "Faster loops with Growth, AEs, SEs — trains on time",
        "Player-coach: Tell → Show → Observe → Feedback",
        "Questions — and happy to pitch one of you live right now",
    ], N21)

    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
