import io, re
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from docx import Document
from xhtml2pdf import pisa

def text_to_pdf(text):
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=A4)
    w, h = A4
    y = h - 45
    for line in text.splitlines() or [""]:
        if y < 45:
            c.showPage(); y = h - 45
        c.drawString(40, y, line[:110])
        y -= 14
    c.save()
    return out.getvalue()

def text_to_docx(text):
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

def text_to_html(text):
    escaped = (text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
    return "<!doctype html><html><body><pre>" + escaped + "</pre></body></html>"

def text_to_markdown(text):
    return text

def html_to_pdf(html_content):
    out = io.BytesIO()
    result = pisa.CreatePDF(io.BytesIO(html_content.encode("utf-8")), dest=out)
    if result.err:
        raise ValueError("Could not convert HTML to PDF.")
    return out.getvalue()

def html_to_text(html_content):
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html_content, flags=re.S|re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def markdown_to_html(md_text):
    lines = []
    for line in md_text.splitlines():
        if line.startswith("# "): lines.append(f"<h1>{line[2:]}</h1>")
        elif line.startswith("## "): lines.append(f"<h2>{line[3:]}</h2>")
        elif line.startswith("### "): lines.append(f"<h3>{line[4:]}</h3>")
        else: lines.append(f"<p>{line}</p>")
    return "<!doctype html><html><body>" + "\n".join(lines) + "</body></html>"
