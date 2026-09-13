from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from html import escape
import os

DOCX_FILE = "resume.docx"
OUTPUT_DIR = "site"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "index.html")

os.makedirs(OUTPUT_DIR, exist_ok=True)

doc = Document(DOCX_FILE)


def pt_to_px(points):
    if points is None:
        return None
    return round(points * 96 / 72, 2)


def twips_to_px(twips):
    if twips is None:
        return None
    return round(twips / 20 * 96 / 72, 2)


def paragraph_alignment(paragraph):
    alignment = paragraph.alignment

    if alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    elif alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    elif alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    return "left"


def get_paragraph_spacing(paragraph):
    fmt = paragraph.paragraph_format

    styles = []

    # Alignment
    styles.append(f"text-align: {paragraph_alignment(paragraph)}")

    # Space before
    if fmt.space_before is not None:
        px = pt_to_px(fmt.space_before.pt)
        styles.append(f"margin-top: {px}px")
    else:
        styles.append("margin-top: 0")

    # Space after
    if fmt.space_after is not None:
        px = pt_to_px(fmt.space_after.pt)
        styles.append(f"margin-bottom: {px}px")
    else:
        styles.append("margin-bottom: 0")

    # Line spacing
    if fmt.line_spacing is not None:
        if hasattr(fmt.line_spacing, "pt"):
            px = pt_to_px(fmt.line_spacing.pt)
            styles.append(f"line-height: {px}px")
        elif isinstance(fmt.line_spacing, (int, float)):
            styles.append(f"line-height: {fmt.line_spacing}")

    # Left indent
    if fmt.left_indent is not None:
        px = pt_to_px(fmt.left_indent.pt)
        styles.append(f"margin-left: {px}px")

    # Right indent
    if fmt.right_indent is not None:
        px = pt_to_px(fmt.right_indent.pt)
        styles.append(f"margin-right: {px}px")

    # First line / hanging indent
    if fmt.first_line_indent is not None:
        px = pt_to_px(fmt.first_line_indent.pt)
        styles.append(f"text-indent: {px}px")

    return "; ".join(styles)


def get_run_style(run):
    styles = []

    # Font size
    if run.font.size is not None:
        px = pt_to_px(run.font.size.pt)
        styles.append(f"font-size: {px}px")

    # Font name
    if run.font.name:
        styles.append(f"font-family: '{escape(run.font.name)}'")

    return "; ".join(styles)


def render_run(run):
    text = escape(run.text)

    if not text:
        return ""

    style = get_run_style(run)

    if style:
        text = f'<span style="{style}">{text}</span>'

    if run.bold:
        text = f"<strong>{text}</strong>"

    if run.italic:
        text = f"<em>{text}</em>"

    if run.underline:
        text = f"<u>{text}</u>"

    return text


def render_hyperlink(hyperlink, paragraph):
    text = ""

    for child in hyperlink.iter():
        if child.tag == qn("w:t") and child.text:
            text += child.text

    text = escape(text)

    rel_id = hyperlink.get(qn("r:id"))

    if not rel_id:
        return text

    try:
        url = paragraph.part.rels[rel_id].target_ref
        return f'<a href="{escape(url)}" target="_blank">{text}</a>'
    except KeyError:
        return text


def render_paragraph_contents(paragraph):
    output = []

    for child in paragraph._p:
        if child.tag == qn("w:r"):
            for run in paragraph.runs:
                if run._r is child:
                    output.append(render_run(run))
                    break

        elif child.tag == qn("w:hyperlink"):
            output.append(render_hyperlink(child, paragraph))

    return "".join(output)


def is_list_paragraph(paragraph):
    style_name = paragraph.style.name if paragraph.style else ""

    if style_name == "List Paragraph":
        return True

    pPr = paragraph._p.pPr
    if pPr is not None and pPr.numPr is not None:
        return True

    return False


html = []

html.append("""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Donald Gardner Resume</title>

<style>
html, body {
    margin: 0;
    padding: 0;
}

body {
    max-width: 850px;
    margin-left: auto;
    margin-right: auto;
    padding: 30px 25px;
    font-family: Arial, Helvetica, sans-serif;
}

p {
    padding: 0;
}

ul {
    padding-left: 28px;
}

li {
    padding: 0;
}

a {
    color: inherit;
}
</style>
</head>
<body>
""")

in_list = False

for paragraph in doc.paragraphs:
    text = paragraph.text

    # Preserve blank paragraphs from Word
    if not text.strip():
        if in_list:
            html.append("</ul>")
            in_list = False

        spacing = get_paragraph_spacing(paragraph)
        html.append(f'<p style="{spacing}">&nbsp;</p>')
        continue

    contents = render_paragraph_contents(paragraph)
    spacing = get_paragraph_spacing(paragraph)

    if is_list_paragraph(paragraph):
        if not in_list:
            html.append("<ul>")
            in_list = True

        html.append(f'<li style="{spacing}">{contents}</li>')
        continue

    if in_list:
        html.append("</ul>")
        in_list = False

    html.append(f'<p style="{spacing}">{contents}</p>')


if in_list:
    html.append("</ul>")

html.append("""
</body>
</html>
""")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(html))

print(f"Created {OUTPUT_FILE}")
