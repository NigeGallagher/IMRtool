"""
IMR contributor submission processor.

This deliberately does NOT hand-build IDML XML. The earlier version of this
tool spent a long debugging session fighting InDesign over hand-rolled IDML
(ID clashes, missing required Spread/Page attributes, font inheritance via
Properties/AppliedFont, etc) and still came out the other side with fonts
not resolving correctly.

Instead, this produces a clean .docx with the IMR paragraph styles already
named and formatted (Article Title, Byline, Standfirst, Body Text First,
Body Text, Subhead, Pull Quote, Caption, Endnote Text, Footer). When you
File > Place that .docx into an InDesign document that already has
paragraph styles with those same names, InDesign maps the incoming text to
your existing styles automatically (tick "Preserve Styling" / use the style
mapping dialog on place). No raw IDML to fight with, and you can always
open the .docx yourself to sanity-check the content before it ever touches
InDesign.
"""

import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn

# ─── IMR typographic system ───
# (font, size, leading, bold, italic, all_caps)
STYLES = {
    'Article Title':   ('Barlow Condensed', 36, 34, True,  False, True),
    'Byline':          ('Barlow Condensed', 12, 14, True,  False, True),
    'Standfirst':      ('Source Serif 4',   13, 18, True,  False, False),
    'Body Text First': ('Source Serif 4',   10, 15, False, False, False),
    'Body Text':       ('Source Serif 4',   10, 15, False, False, False),
    'Subhead':         ('Barlow Condensed', 13, 16, True,  False, True),
    'Pull Quote':      ('Barlow Condensed', 18, 20, True,  False, True),
    'Caption':         ('Source Serif 4',    9, 12, False, True,  False),
    'Endnote Text':    ('Source Serif 4',    9, 12, False, False, False),
    'Footer':          ('Barlow Condensed',  9, 11, True,  False, True),
}

MARKER_STYLE = {
    '[PULLQUOTE]': 'Pull Quote',
    '[SUBHEAD]':   'Subhead',
    '[CAPTION]':   'Caption',
    '[ENDNOTE]':   'Endnote Text',
}


def _set_all_caps(style, value):
    """python-docx doesn't expose all_caps on every version cleanly via the
    high level API in all paragraph-style contexts, so set it directly on
    the underlying rPr to be safe."""
    rpr = style.element.get_or_add_rPr()
    caps = rpr.find(qn('w:caps'))
    if value:
        if caps is None:
            caps = rpr.makeelement(qn('w:caps'), {})
            rpr.append(caps)
    else:
        if caps is not None:
            rpr.remove(caps)


def _get_or_add_style(doc, name):
    """'Body Text' etc already exist as Word built-ins, so reuse them
    rather than crashing on add_style."""
    for s in doc.styles:
        if s.name == name and s.type == WD_STYLE_TYPE.PARAGRAPH:
            return s
    return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def _build_styles(doc):
    for name, (font, size, leading, bold, italic, caps) in STYLES.items():
        style = _get_or_add_style(doc, name)
        style.font.name = font
        # Force the eastasian font tag too, otherwise Word/InDesign can
        # silently fall back to a default font for some character ranges.
        style.element.rPr.rFonts.set(qn('w:eastAsia'), font)
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.italic = italic
        _set_all_caps(style, caps)
        pf = style.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(leading)
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
    return doc


def extract_body_paragraphs(filepath):
    """Read the contributor's uploaded .docx and classify each paragraph
    into an IMR style based on [MARKER] tags or Word heading levels."""
    src = Document(filepath)
    result = []
    first_body_seen = False

    for para in src.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = None
        for marker, mapped in MARKER_STYLE.items():
            if text.upper().startswith(marker):
                text = text[len(marker):].strip()
                style_name = mapped
                break

        if style_name is None:
            word_style = (para.style.name or '').lower()
            if 'heading' in word_style:
                style_name = 'Subhead'
            else:
                style_name = 'Body Text First' if not first_body_seen else 'Body Text'
                first_body_seen = True

        result.append((style_name, text))

    return result


def process_docx(filepath, title, author, standfirst, article_type,
                  output_folder, timestamp):
    """Build the IMR-styled output .docx and return its path."""
    out = Document()

    # Strip the default boilerplate styles isn't necessary - we just add ours
    _build_styles(out)

    out.add_paragraph(title, style='Article Title')
    out.add_paragraph(f'By {author}' if author else '', style='Byline')
    if standfirst:
        out.add_paragraph(standfirst, style='Standfirst')

    for style_name, text in extract_body_paragraphs(filepath):
        out.add_paragraph(text, style=style_name)

    safe_title = ''.join(c if c.isalnum() or c in ' -_' else '' for c in title)[:50].strip() or 'untitled'
    filename = f"{timestamp}_{safe_title.replace(' ', '_')}_IMR.docx"
    out_path = os.path.join(output_folder, filename)
    out.save(out_path)
    return out_path
