import io, re
import fitz

def pdf_page_count(data):
    doc = fitz.open(stream=data, filetype="pdf")
    n = len(doc); doc.close()
    return n

def pdf_to_text(data):
    doc = fitz.open(stream=data, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    return text

def pdf_to_markdown(data):
    return pdf_to_text(data)

def pdf_to_html(data):
    return "<html><body><pre>" + pdf_to_text(data).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;") + "</pre></body></html>"

def pdf_to_csv(data):
    return pdf_to_text(data)

def pdf_to_images(data, dpi=150, fmt="png", page_range=None):
    doc = fitz.open(stream=data, filetype="pdf")
    pages = page_range if page_range is not None else range(len(doc))
    result = []
    for i in pages:
        page = doc[i]
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        result.append((i+1, pix.tobytes(fmt)))
    doc.close()
    return result

def pdf_to_docx(data, preserve_layout=True, extract_images=False):
    from docx import Document
    doc = Document()
    for page in fitz.open(stream=data, filetype="pdf"):
        text = page.get_text()
        for line in text.splitlines():
            doc.add_paragraph(line)
    out = io.BytesIO(); doc.save(out)
    return out.getvalue()

def images_to_pdf(images, page_size="Original", orientation="Portrait", margin_pt=0, fit="Fit"):
    from PIL import Image
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4, LETTER
    out = io.BytesIO()
    if page_size == "A4": default = A4
    elif page_size == "Letter": default = LETTER
    else: default = A4
    c = canvas.Canvas(out, pagesize=default)
    for data in images:
        img = Image.open(io.BytesIO(data))
        iw, ih = img.size
        pw, ph = default
        if orientation == "Landscape": pw, ph = ph, pw
        if page_size == "Original":
            # approximate points from pixels at 72 dpi
            pw, ph = iw, ih
            c.setPageSize((pw, ph))
        scale = min((pw-2*margin_pt)/iw, (ph-2*margin_pt)/ih)
        if fit == "Stretch":
            dw, dh = pw-2*margin_pt, ph-2*margin_pt
        else:
            dw, dh = iw*scale, ih*scale
        x, y = (pw-dw)/2, (ph-dh)/2
        img_buf = io.BytesIO(); img.convert("RGB").save(img_buf, "JPEG")
        img_buf.seek(0)
        from reportlab.lib.utils import ImageReader
        c.drawImage(ImageReader(img_buf), x, y, width=dw, height=dh)
        c.showPage()
    c.save()
    return out.getvalue()

def is_scanned_pdf(data):
    doc = fitz.open(stream=data, filetype="pdf")
    scanned = all(len(p.get_text().strip()) == 0 for p in doc) if len(doc) else False
    doc.close()
    return scanned

def _ocr_image_to_pdf_bytes(image_bytes, lang):
    import pytesseract
    from PIL import Image
    img = Image.open(io.BytesIO(image_bytes))
    pdf = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="pdf")
    return pdf

def searchable_pdf_from_ocr(data, lang="eng"):
    import pytesseract
    doc = fitz.open(stream=data, filetype="pdf")
    out = fitz.open()
    for page in doc:
        pix = page.get_pixmap(dpi=200, alpha=False)
        img = pix.tobytes("png")
        pdf_page = _ocr_image_to_pdf_bytes(img, lang)
        ocr_doc = fitz.open(stream=pdf_page, filetype="pdf")
        out.insert_pdf(ocr_doc)
        ocr_doc.close()
    result = out.tobytes()
    out.close(); doc.close()
    return result

def merge_pdfs(items):
    out = fitz.open()
    for data in items:
        d = fitz.open(stream=data, filetype="pdf"); out.insert_pdf(d); d.close()
    result = out.tobytes(); out.close()
    return result

def split_pdf(data, every=1):
    doc = fitz.open(stream=data, filetype="pdf")
    parts = []
    for start in range(0, len(doc), every):
        out = fitz.open()
        out.insert_pdf(doc, from_page=start, to_page=min(start+every-1, len(doc)-1))
        parts.append((f"part_{start//every+1}", out.tobytes()))
        out.close()
    doc.close()
    return parts

def compress_pdf(data, quality=60):
    doc = fitz.open(stream=data, filetype="pdf")
    out = doc.tobytes(garbage=4, deflate=True, clean=True)
    doc.close()
    return out

def rotate_pdf(data, degrees):
    doc = fitz.open(stream=data, filetype="pdf")
    for p in doc: p.set_rotation((p.rotation + degrees) % 360)
    out = doc.tobytes(); doc.close()
    return out

def extract_pages(data, indexes):
    doc = fitz.open(stream=data, filetype="pdf"); out = fitz.open()
    for i in indexes:
        if 0 <= i < len(doc): out.insert_pdf(doc, from_page=i, to_page=i)
    result = out.tobytes(); out.close(); doc.close()
    return result

def delete_pages(data, indexes):
    doc = fitz.open(stream=data, filetype="pdf")
    for i in sorted(set(indexes), reverse=True):
        if 0 <= i < len(doc): doc.delete_page(i)
    out = doc.tobytes(); doc.close()
    return out

def reorder_pages(data, order):
    doc = fitz.open(stream=data, filetype="pdf"); out = fitz.open()
    for i in order:
        if 0 <= i < len(doc): out.insert_pdf(doc, from_page=i, to_page=i)
    result = out.tobytes(); out.close(); doc.close()
    return result

def add_watermark(data, text, opacity=0.3):
    doc = fitz.open(stream=data, filetype="pdf")
    for p in doc:
        rect = p.rect
        p.insert_text((rect.width/2-80, rect.height/2), text, fontsize=30, rotate=45, color=(0.5,0.5,0.5), fill_opacity=opacity)
    out = doc.tobytes(); doc.close()
    return out

def add_page_numbers(data, start_at=1, pos="bottom-center"):
    doc = fitz.open(stream=data, filetype="pdf")
    for i, p in enumerate(doc):
        y = p.rect.height - 20
        if pos == "bottom-left": x = 30
        elif pos == "bottom-right": x = p.rect.width - 50
        else: x = p.rect.width/2 - 10
        p.insert_text((x,y), str(start_at+i), fontsize=10)
    out = doc.tobytes(); doc.close()
    return out
