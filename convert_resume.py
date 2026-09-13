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


def get_alignment(paragraph):
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    return "left"


def paragraph_has_small_caps(paragraph):
    """
    Identify section titles by their Word formatting,
    rather than by the actual words in the title.
    """
    for run in paragraph.runs:
        if run.text.strip() and run.font.small_caps:
            return True

    return False


def get_paragraph_style(paragraph):
    fmt = paragraph.paragraph_format
    styles = []

    # Copy Word paragraph alignment
    styles.append(
        f"text-align: {get_alignment(paragraph)}"
    )

    # Copy Word spacing before paragraph
    if fmt.space_before is not None:
        styles.append(
            f"margin-top: {pt_to_px(fmt.space_before.pt)}px"
        )
    else:
        styles.append("margin-top: 0")

    # Copy Word spacing after paragraph initially
    if fmt.space_after is not None:
        styles.append(
            f"margin-bottom: {pt_to_px(fmt.space_after.pt)}px"
        )
    else:
        styles.append("margin-bottom: 0")

    # --------------------------------------------------
    # SECTION TITLE RULE
    #
    # Centered + Small Caps means a section title.
    #
    # The title itself is single spaced.
    # The space AFTER the title is 1.5 lines.
    # --------------------------------------------------

    if (
        paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and paragraph_has_small_caps(paragraph)
    ):
        styles.append("line-height: 1")
        styles.append("margin-bottom: 1.5em")

    else:
        # For all other paragraphs, preserve Word's
        # line-spacing information.
        if fmt.line_spacing is not None:

            if hasattr(fmt.line_spacing, "pt"):
                styles.append(
                    f"line-height: "
                    f"{pt_to_px(fmt.line_spacing.pt)}px"
                )

            elif isinstance(
                fmt.line_spacing,
                (int, float)
            ):
                styles.append(
                    f"line-height: {fmt.line_spacing}"
                )

    # Copy Word left indentation
    if fmt.left_indent is not None:
        styles.append(
            f"margin-left: "
            f"{pt_to_px(fmt.left_indent.pt)}px"
        )

    # Copy Word right indentation
    if fmt.right_indent is not None:
        styles.append(
            f"margin-right: "
            f"{pt_to_px(fmt.right_indent.pt)}px"
        )

    # Copy Word first-line/hanging indentation
    if fmt.first_line_indent is not None:
        styles.append(
            f"text-indent: "
            f"{pt_to_px(fmt.first_line_indent.pt)}px"
        )

    return "; ".join(styles)


def get_run_style(run):
    styles = []

    # Copy Word font size
    if run.font.size is not None:
        styles.append(
            f"font-size: {pt_to_px(run.font.size.pt)}px"
        )

    # Copy Word font
    if run.font.name:
        styles.append(
            f"font-family: '{escape(run.font.name)}'"
        )

    # Copy Word Small Caps
    if run.font.small_caps:
        styles.append(
            "font-variant: small-caps"
        )

    return "; ".join(styles)


def render_run(run):
    text = escape(run.text)

    if not text:
        return ""

    style = get_run_style(run)

    if style:
        text = (
            f'<span style="{style}">'
            f'{text}'
            f'</span>'
        )

    # Copy Word bold
    if run.bold:
        text = f"<strong>{text}</strong>"

    # Copy Word italics
    if run.italic:
        text = f"<em>{text}</em>"

    # Copy Word underline
    if run.underline:
        text = f"<u>{text}</u>"

    return text


def render_hyperlink(hyperlink, paragraph):
    text = ""

    for child in hyperlink.iter():
        if (
            child.tag == qn("w:t")
            and child.text
        ):
            text += child.text

    text = escape(text)

    rel_id = hyperlink.get(qn("r:id"))

    if not rel_id:
        return text

    try:
        url = paragraph.part.rels[
            rel_id
        ].target_ref

        return (
            f'<a href="{escape(url)}" '
            f'target="_blank">'
            f'{text}'
            f'</a>'
        )

    except KeyError:
        return text


def render_contents(paragraph):
    output = []

    for child in paragraph._p:

        if child.tag == qn("w:r"):

            for run in paragraph.runs:
                if run._r is child:
                    output.append(
                        render_run(run)
                    )
                    break

        elif child.tag == qn("w:hyperlink"):

            output.append(
                render_hyperlink(
                    child,
                    paragraph
                )
            )

    return "".join(output)


def is_list_paragraph(paragraph):

    if (
        paragraph.style
        and paragraph.style.name
        == "List Paragraph"
    ):
        return True

    pPr = paragraph._p.pPr

    if (
        pPr is not None
        and pPr.numPr is not None
    ):
        return True

    return False


html = []

html.append("""<!DOCTYPE html>
<html lang="en">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>Donald Gardner Resume</title>

<style>

html,
body {
    margin: 0;
    padding: 0;
}

body {
    max-width: 850px;
    margin-left: auto;
    margin-right: auto;
    padding: 18px 20px;

    font-family:
        Arial,
        Helvetica,
        sans-serif;
}

p {
    padding: 0;
}

.blank-line {
    height: 1em;
    margin: 0;
    padding: 0;
}

ul {
    margin-top: 0;
    margin-bottom: 0;
    padding-left: 28px;
}

li {
    margin-top: 0;
    margin-bottom: 0;
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

    text = paragraph.text.strip()

    # -----------------------------------------------
    # BLANK WORD PARAGRAPH
    # Preserve it as one blank line.
    # -----------------------------------------------

    if not text:

        if in_list:
            html.append("</ul>")
            in_list = False

        html.append(
            '<div class="blank-line"></div>'
        )

        continue


    contents = render_contents(paragraph)

    paragraph_style = get_paragraph_style(
        paragraph
    )


    # -----------------------------------------------
    # WORD BULLET / LIST
    # -----------------------------------------------

    if is_list_paragraph(paragraph):

        if not in_list:
            html.append("<ul>")
            in_list = True

        html.append(
            f'<li style="{paragraph_style}">'
            f'{contents}'
            f'</li>'
        )

        continue


    # Close list before next normal paragraph
    if in_list:
        html.append("</ul>")
        in_list = False


    # -----------------------------------------------
    # NORMAL WORD PARAGRAPH
    # -----------------------------------------------

    html.append(
        f'<p style="{paragraph_style}">'
        f'{contents}'
        f'</p>'
    )


if in_list:
    html.append("</ul>")


html.append("""
</body>
</html>
""")


with open(
    OUTPUT_FILE,
    "w",
    encoding="utf-8"
) as f:

    f.write(
        "\n".join(html)
    )


print(
    f"Created {OUTPUT_FILE}"
)
