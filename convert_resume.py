from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from html import escape

DOCX_FILE = "resume.docx"
OUTPUT_FILE = "site/index.html"

doc = Document(DOCX_FILE)


def paragraph_alignment(paragraph):
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    elif paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    return "left"


def render_run(run):
    text = escape(run.text)

    if not text:
        return ""

    if run.bold:
        text = f"<strong>{text}</strong>"

    if run.italic:
        text = f"<em>{text}</em>"

    if run.underline:
        text = f"<u>{text}</u>"

    return text


html_parts = []

html_parts.append("""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Donald Gardner Resume</title>

<style>
body {
    max-width: 850px;
    margin: 40px auto;
    padding: 0 25px;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 16px;
    line-height: 1.4;
}

p {
    margin-top: 6px;
    margin-bottom: 6px;
}

ul {
    margin-top: 5px;
    margin-bottom: 10px;
    padding-left: 28px;
}

li {
    margin-bottom: 5px;
}

.center {
    text-align: center;
}

.left {
    text-align: left;
}

.right {
    text-align: right;
}

.justify {
    text-align: justify;
}

.resume-heading {
    font-weight: bold;
    margin-top: 18px;
    margin-bottom: 10px;
}
</style>

</head>
<body>
""")

in_list = False

for paragraph in doc.paragraphs:

    text = paragraph.text.strip()

    if not text:
        if in_list:
            html_parts.append("</ul>")
            in_list = False

        html_parts.append("<br>")
        continue

    alignment = paragraph_alignment(paragraph)

    is_list = paragraph.style.name == "List Paragraph"

    rendered_runs = "".join(render_run(run) for run in paragraph.runs)

    if is_list:
        if not in_list:
            html_parts.append("<ul>")
            in_list = True

        html_parts.append(f"<li>{rendered_runs}</li>")
        continue

    if in_list:
        html_parts.append("</ul>")
        in_list = False

    all_bold = (
        paragraph.runs
        and all(
            run.bold is True
            for run in paragraph.runs
            if run.text.strip()
        )
    )

    extra_class = ""

    if all_bold:
        extra_class = " resume-heading"

    html_parts.append(
        f'<p class="{alignment}{extra_class}">{rendered_runs}</p>'
    )

if in_list:
    html_parts.append("</ul>")

html_parts.append("""
</body>
</html>
""")

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write("\n".join(html_parts))

print(f"Created {OUTPUT_FILE}")
