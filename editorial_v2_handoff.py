"""
Editorial v2 Template — Claude Code Handoff Document
=====================================================
Generated: March 12, 2026
Context: Multi-session design collaboration between Chad and Claude (claude.ai)
Purpose: Integrate the editorial v2 template into the Slide Creator Flask app

TABLE OF CONTENTS
1. What This Is
2. Architecture & Interface
3. Builder Inventory (21 builders, 22 entries)
4. Design Language Reference
5. Data Contract Compatibility
6. Integration Notes
7. Known Issues & Intentional Quirks
8. Future Work

=====================================================
1. WHAT THIS IS
=====================================================

template_editorial_v2.py is a complete PPTX template for the Slide Creator app.
It is NOT a reskin of the existing templates (slick, colorful, bold, noir, editorial v1).
It is a ground-up redesign where each slide builder's layout serves its rhetorical
purpose rather than following a uniform grid.

Key differences from the other 5 templates:
- 21 purpose-driven builders (vs 38 uniform builders in the others)
- Multi-slide builders that generate 2-5 slides from one config entry
- Progressive reveal patterns (in_brief, wsn_reveal, comparison_reveal, summary)
- Asymmetric layouts (comparison_reveal, persona)
- Newspaper/magazine design language (agenda masthead, text_narrative columns)
- "Every slide should look like an infographic" as the design principle

The file is 1,093 lines of Python using python-pptx. No external dependencies
beyond python-pptx itself.

=====================================================
2. ARCHITECTURE & INTERFACE
=====================================================

The template follows the exact same interface as the other 5 templates:

    from template_editorial_v2 import build_deck

    slide_configs = [
        ("title", {"title": "My Deck", "subtitle": "A subtitle", ...}),
        ("stat_callout", {"stat": "42%", "headline": "...", ...}),
        ("closer", {"title": "Thank You", ...}),
    ]

    build_deck(slide_configs, "output.pptx")

Each config entry is a tuple of (slide_type_string, data_dict).
The BUILDERS dict maps slide_type strings to builder functions.
Unknown slide types are silently skipped (same behavior as other templates).

Slide dimensions: 10" x 5.625" (widescreen 16:9)

=====================================================
3. BUILDER INVENTORY
=====================================================

STRUCTURAL SLIDES (dark green background):

  title
    Keys: title, subtitle, author, date
    Output: 1 slide
    Notes: Gold top rule, large serif title, gold separator, footer

  section_divider
    Keys: title, items (list of agenda items), current (0-indexed, optional)
    Output: 1 slide
    Notes: When "current" param is passed, shows breadcrumbs of completed
           sections at top, hero section in center, upcoming sections below.
           When called without "current", behaves like a standard divider.
           IMPORTANT: This builder's signature is build_section_divider(prs, c, current=0).
           The "current" param is NOT in the data dict — it's a function parameter.
           The Flask app will need to pass it explicitly.

  closer
    Keys: title, subtitle, contact
    Output: 1 slide
    Notes: Bookends with title slide. Italic closing thought.

AGENDA (white background, newspaper layout):

  agenda
    Keys: title, items (list of dicts with "title" and optional "detail")
    Output: 1 slide
    Notes: Newspaper masthead design. "Agenda" as banner headline.
           All items equal weight in 2-3 column grid with colored left bars.
           The masthead uses 9pt centered text — this is INTENTIONAL, not a bug.
           Do not "fix" it to match the 14pt running heads on other slides.

SINGLE-POINT SLIDES:

  stat_callout
    Keys: title, stat, headline, detail, source
    Output: 1 slide
    Notes: Giant number (120pt) on left 55%, context stacked on right.
           Also handles stat_hero and multi_stat use cases from other templates.

  quote_full
    Keys: quote, attribution, context
    Output: 1 slide (dark green background)
    Notes: Large serif italic quote with gold decorative quotation mark.
           Auto-sizes font based on quote length (28/24/20pt).
           Also handles the regular "quote" slide type from other templates.

COMPARISON SLIDES:

  comparison
    Keys: title, leftLabel, rightLabel, leftItems, rightItems
    Output: 1 slide
    Notes: Row-based paired cards. Warm tint (left) + cool tint (right).
           Green/cobalt left-edge accent bars. Clean, scannable default.

  comparison_reveal
    Keys: title, leftLabel, rightLabel, leftItems, rightItems
    Output: 2 slides
    Notes: SAME data contract as comparison. Slide 1: left dominant (large cards),
           right compressed (small boxes with faint left-edge accent). Slide 2: flipped.
           Rows are aligned across both sides for cross-reading.

PROGRESSIVE REVEAL SLIDES:

  in_brief
    Keys: title, bullets (list of strings)
    Output: N slides (one per bullet)
    Notes: Each bullet gets hero treatment (28pt bold). Previous bullets stack
           below as "PREVIOUSLY" running summary. Progress bar + "3 of 5" counter.

  wsn_reveal
    Keys: title, what (dict), soWhat (dict), nowWhat (dict)
           Each zone dict has: headline, detail
    Output: 3 slides
    Notes: Progressive What -> So What -> Now What. Active zone is hero,
           previous zones compress into sidebar with colored accent bars.
           Colors: What=DK_GREEN, So What=COBALT, Now What=PURPLE.

  summary
    Keys: title, heading, sections (list of dicts with heading + points), points (list)
    Output: 2 slides
    Notes: Slide 1: sections as dominant column cards, takeaways teased in compressed
           band at bottom. Slide 2: sections compress to narrow band, takeaways
           expand into 2x2 warm cards with colored left bars.

ANALYSIS SLIDES:

  findings_recs
    Keys: title, items (list of dicts with "finding" and "recommendation")
    Output: 1 slide
    Notes: Finding->Recommendation paired rows. Warm cards (left) with gold arrows
           pointing to cool cards (right). Green/cobalt left-edge accents.
           Also handles findings_recs_dense and hypotheses from other templates.

  process_flow
    Keys: title, steps (list of dicts with "label"/"title" and "detail")
    Output: 1 slide
    Notes: Horizontal layout for <=4 steps. Colored top bars, numbered circles,
           gold arrows between. Title is a 22pt headline (not a running head).

  process_flow_vertical
    Keys: title, steps (list of dicts with "label"/"title" and "detail")
    Output: 1 slide
    Notes: Film strip layout. Full-width horizontal bands with giant watermark
           numbers, colored left bars, alternating warm/light backgrounds.
           Default for 5+ steps.

  process_flow_accordion
    Keys: title, steps (list of dicts with "label"/"title" and "detail")
    Output: 1 slide
    Notes: VARIANT — alternating left/right cards with center spine.
           Warm cards on left, cool cards on right, numbered dots on spine.
           All cards have colored left-edge accent bars.

  timeline
    Keys: title, milestones (list of dicts with date, title/label, detail, status)
    Output: 1 slide
    Notes: Route map layout. Horizontal line with alternating above/below stops.
           Status colors: complete=DK_GREEN, current=GOLD, upcoming=QUIET.

CHARACTER / TEXT SLIDES:

  persona
    Keys (single): title, name, archetype, traits, detail, strategy
    Keys (duo): title, personas (list of 2 dicts with name, archetype, traits, detail)
    Output: 1 slide
    Notes: Auto-detects single vs duo based on presence of "personas" key.
           Single: dark green nameplate (left 1/3) with image placeholder + name +
           archetype + traits. Right side: detail in light card + strategy in warm card.
           Duo: side-by-side nameplates (green vs cobalt), traits below, detail at bottom.
           Image placeholders render as dark tinted rectangles with "IMG" label.
           Title uses 18pt uppercase (larger than running head, smaller than headline).

  persona_duo
    Maps to: build_persona (same function handles both)

  text_narrative
    Keys: title, lede, body
    Output: 1 slide
    Notes: Magazine article layout. Bold lede (22pt) across full width, gold rule,
           then body text auto-split into two columns (warm card left, cool card right).
           Column split finds nearest sentence break at midpoint.

WRAP-UP SLIDES:

  open_questions
    Keys: title, questions (list of strings or dicts with "text")
    Output: 1 slide
    Notes: 2x2 provocation cards with giant watermark "?" in each card.
           Colored top bars, numbered. Title is 22pt headline.

  data_table
    Keys: title, headers (list), rows (list of lists), note, highlightCol (int)
    Output: 1 slide
    Notes: Dark green header row, alternating row backgrounds. Highlighted column
           gets warm tint + bold green text. Title is 22pt headline.

  appendix
    Keys: title, label, sections (list of dicts with "label" and "content")
    Output: 1 slide
    Notes: INTENTIONALLY the quietest slide. All gray text, no colored accents.
           Endnote energy. Label/content two-column layout.

=====================================================
4. DESIGN LANGUAGE REFERENCE
=====================================================

PALETTE:
  DK_GREEN    #044014   Title/section/closer bg, primary accent
  GOLD        #D4A843   Emphasis, separators, numbers, stat callout
  CHARCOAL    #2D2D2D   Primary body text
  BLACK       #1A1A1A   Headlines
  MID         #666666   Secondary body text, detail
  QUIET       #999999   Running heads, captions, muted elements
  FAINT       #CCCCCC   Watermarks, compressed items, faint accents
  RULE_CLR    #DDDDDD   Hairline rules, column dividers
  WHITE       #FFFFFF   Main content background
  WARM_CARD   #F0EDE6   Left/primary cards, warm containment
  COOL_CARD   #EAEEF0   Right/secondary cards, cool containment
  LIGHT_BOX   #F7F6F4   Subtle containment, compressed items
  COBALT      #04547C   Right-side accent, secondary color
  PURPLE      #5B2C8F   Third accent (Now What zone, item markers)
  CREAM       #E8E0CC   Text on dark backgrounds
  LIGHT_GREEN #A8C4A0   Footer text on dark backgrounds

ACCENT CYCLE (for numbered items, card markers):
  DK_GREEN -> COBALT -> PURPLE -> GOLD -> TEAL -> RED -> OCHRE

FONTS:
  TITLE_FONT  "Fidelity Slab"   Headlines, titles, numbers
  BODY_FONT   "Fidelity Sans"   Body text, labels, running heads
  Fallback handling: inherits from existing templates' _font_fallback system

TYPOGRAPHY SCALE:
  120pt  Stat callout hero number
  48pt   Section divider number, film strip watermark numbers
  28pt   In_brief hero bullet
  22pt   Slide headlines (process_flow, open_questions, data_table, summary)
  20pt   Persona name, comparison_reveal dominant label
  18pt   Persona running head (compromise size)
  16pt   Process flow step labels
  14pt   Running head (most slides), section card headings
  13pt   Body text in cards, comparison items, process detail
  12pt   Standard body text, findings/recs items
  11pt   Table cells, compressed labels
  10pt   Traits, small detail text
  9pt    Agenda masthead (INTENTIONAL — newspaper design)
  8-9pt  Sources, captions, "PREVIOUSLY" labels, progress counters

VISUAL PATTERNS:
  - DK_GREEN top rule (4pt) on ALL white-background slides
  - GOLD top rule (4pt) on ALL dark-background slides
  - Left-edge accent bars on cards (0.05-0.06" wide)
  - Colored top bars on cards (0.05" tall)
  - Warm/cool card pairing for left/right comparisons
  - Gold rules as section separators (2-3pt)
  - RULE_CLR hairlines for subtle structure (1pt)
  - Running head: UPPERCASE, 14pt, QUIET gray, top-left

=====================================================
5. DATA CONTRACT COMPATIBILITY
=====================================================

The editorial v2 template accepts the same data dicts as the other templates.
However, it only implements 21 of the 38 builder types. When editorial v2
encounters an unknown slide type, it silently skips it (same behavior).

MAPPING FROM OTHER TEMPLATES' 38 TYPES:

  Direct support (same key):
    title, closer, agenda, stat_callout, comparison, process_flow,
    findings_recs, timeline, data_table, appendix, open_questions

  Renamed/merged:
    quote -> quote_full (editorial only has the full-page version)
    quote_full -> quote_full
    in_brief -> in_brief (but generates N slides, not 1)
    in_brief_featured -> in_brief (featured bullet becomes first in reveal)
    wsn_dense -> wsn_reveal (editorial only does the reveal)
    wsn_reveal -> wsn_reveal
    persona -> persona
    persona_duo -> persona (auto-detected)
    process_flow_vertical -> process_flow_vertical
    text_narrative -> text_narrative
    text_split -> text_narrative (headline=lede, detail+points=body)
    text_columns -> text_narrative

  Editorial-only (no equivalent in other templates):
    comparison_reveal
    process_flow_accordion
    summary (2-slide reveal version)
    section_divider (with currentSection support)

  Not implemented (silently skipped):
    section_divider (standard — use with current=0 for equivalent)
    stat_hero, multi_stat (use stat_callout)
    findings_recs_dense (use findings_recs)
    hypotheses (use findings_recs with different labels)
    progressive_reveal (use in_brief)
    before_after, risk_tradeoff (use comparison)
    matrix (not implemented)
    methods (use text_narrative)
    text_cards, text_nested, text_annotated, text_graph (not implemented)
    summary (standard single-slide — editorial uses 2-slide reveal)

=====================================================
6. INTEGRATION NOTES
=====================================================

FILE PLACEMENT:
  Copy template_editorial_v2.py to C:\Users\chadm\Desktop\slide-creator\
  alongside the existing template_slick.py, template_colorful.py, etc.

FLASK APP CHANGES:
  1. Add "editorial_v2" to the template selector dropdown
  2. Import: from template_editorial_v2 import build_deck
  3. The build_deck interface is identical — no adapter needed
  4. For the section_divider with currentSection support:
     The Flask app currently passes (slide_type, data_dict) tuples.
     The section_divider's "current" param is a function argument, not a data key.
     Options:
       a) Add "currentSection" to the data dict and have the builder read it
       b) Keep the function param and have the Flask app call it differently
     Recommendation: option (a) — add this to build_section_divider:
       current = c.get("currentSection", 0)

MULTI-SLIDE BUILDERS:
  The Flask app's slide list will show 1 entry but the output will have
  multiple slides. This is consistent with wsn_reveal in the existing
  templates (which already generates 3 slides from 1 config).

=====================================================
7. KNOWN ISSUES & INTENTIONAL QUIRKS
=====================================================

INTENTIONAL:
  - Agenda masthead is 9pt centered (newspaper design, not a running head)
  - Appendix has no colored accents (intentionally muted)
  - Persona uses 18pt title (compromise between design purity and readability)
  - process_flow horizontal has NO running head (headline does the job)
  - Comparison_reveal uses FAINT left-edge accent on compressed items
  - WSN reveal slide 1 ("What") is full-width with no sidebar

KNOWN ISSUES TO FIX:
  - Timeline: with 6 milestones, rightmost label clips off slide edge.
    Fix: reduce card_w or add margin when n >= 6.
  - Section_divider: "+N more" text can overlap with the 4th upcoming item.
    Fix: move "+N more" to a new line below the items.
  - In_brief: slide 1 (first bullet) has empty bottom half.
    Acceptable — it's a progressive reveal, context builds on subsequent slides.
  - Text_narrative: column split is character-based, not semantic.
    Works well in practice but could split mid-sentence on edge cases.

=====================================================
8. FUTURE WORK
=====================================================

  - Image support: Replace "IMG" placeholders with actual image rendering
    (logo_path, image_path, or icon support via python-pptx image insertion)
  - Stacked cards agenda variant: The diagonal cascade layout was prototyped
    and liked but not included as a builder. Could add as agenda_cards.
  - Route map process flow: Prototyped but cut in favor of film strip.
    The layout lives on in the timeline builder.
  - Title underline removal: slick, editorial v1, and bold templates still
    have the colored title underline. Remove:
      slick: _slide_title line: _add_rect(slide, LM, Inches(0.85), Inches(2.5), Inches(0.04), GREEN)
      editorial v1: _content_title line: _line(slide, LM, Inches(1.05), Inches(1.8), GOLD_MUTED, 1.5)
      bold: _content_title line: _rect(slide, LM, Inches(0.3), Inches(2.0), Inches(0.08), edge_color)
  - Consider making the 14pt running head size a constant so it can be
    adjusted in one place if the "haters" win again.
"""
