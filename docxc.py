import io
from docx import Document
import txtc as tc

def docx_to_text(data):
    doc = Document(io.BytesIO(data))
    return "\n".join(p.text for p in doc.paragraphs)

def docx_to_markdown(data):
    return docx_to_text(data)

def docx_to_html(data):
    text = docx_to_text(data)
    return tc.text_to_html(text)

def docx_to_pdf(data):
    return tc.text_to_pdf(docx_to_text(data))
