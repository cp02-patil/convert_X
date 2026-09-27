"""
ConvertX - Free Universal File Converter
Convert Anything. Instantly.

A single-process Streamlit app. Every conversion below is real (uses
PyMuPDF / python-docx / openpyxl / python-pptx / Pillow / Tesseract /
reportlab / xhtml2pdf) - nothing here just renames a file extension.

No login. No payment. No subscriptions. No paid APIs. Files only ever
live in memory for the duration of your session.
"""
import io
import zipfile
import mimetypes
from datetime import datetime

import streamlit as st

import pdfc
import imgc
import txtc
import docxc
import xlsxc
import pptxc
import sec


# --------------------------------------------------------------------------
# Page setup
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="ConvertX - Convert Anything. Instantly.",
    page_icon="🔄",
    layout="wide",
)
# Render logo in the sidebar header above the navigation radio buttons
st.sidebar.image("logo.png", use_container_width=True)

MAX_FILE_MB = 50

FORMAT_TARGETS = {
    "pdf": ["docx", "jpg", "png", "txt", "html", "md", "csv", "pptx"],
    "docx": ["pdf", "txt", "html", "md"],
    "xlsx": ["pdf", "csv", "txt", "html"],
    "pptx": ["pdf", "jpg", "png", "txt"],
    "jpg": ["pdf", "png", "webp", "txt"],
    "jpeg": ["pdf", "png", "webp", "txt"],
    "png": ["pdf", "jpg", "webp", "txt"],
    "webp": ["pdf", "jpg", "png", "txt"],
    "txt": ["pdf", "docx", "html", "md"],
    "html": ["pdf", "txt"],
    "md": ["html"],
}

MIME_TYPES = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "png": "image/png",
    "webp": "image/webp",
    "txt": "text/plain",
    "html": "text/html",
    "md": "text/markdown",
    "csv": "text/csv",
    "zip": "application/zip",
}

OCR_LANGUAGES = {
    "English": "eng",
    "Hindi": "hin",
    "Marathi": "mar",
    "Gujarati": "guj",
    "Bengali": "ben",
    "Tamil": "tam",
    "Telugu": "tel",
    "Kannada": "kan",
    "Malayalam": "mal",
    "Punjabi": "pan",
}

# --------------------------------------------------------------------------
# Custom styling
# --------------------------------------------------------------------------
st.markdown("""
<style>
    .cx-hero {text-align:center; padding: 1.2rem 0 0.4rem 0;}
    .cx-hero h1 {font-size: 2.6rem; font-weight: 800; margin-bottom: 0.2rem;}
    .cx-tagline {color: #6b7280; font-size: 1.05rem; margin-bottom: 0.4rem;}
    .cx-badge {display:inline-block; background:#F0FDF4; color:#166534; padding:2px 10px;
               border-radius:999px; font-size:0.75rem; font-weight:600; margin:2px;}
    .cx-privacy {text-align:center; color:#9ca3af; font-size:0.85rem; margin-top:0.5rem;}
    div.stButton > button {border-radius: 10px; font-weight:600;}
    .cx-card {border:1px solid #e5e7eb; border-radius:14px; padding:1rem 1.2rem; margin-bottom:0.8rem;}
</style>
""", unsafe_allow_html=True)


def sizeof(n_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if n_bytes < 1024:
            return f"{n_bytes:.1f} {unit}"
        n_bytes /= 1024
    return f"{n_bytes:.1f} TB"


def error_box(msg="We couldn't convert this file. Please check the file and try again."):
    st.error(msg)


def run_progress(steps):
    bar = st.progress(0, text=steps[0])
    n = len(steps)
    for i, label in enumerate(steps):
        bar.progress(int((i + 1) / n * 100), text=label)
    return bar


# --------------------------------------------------------------------------
# Central conversion manager
# --------------------------------------------------------------------------
def convert(filename: str, file_bytes: bytes, src_fmt: str, target_fmt: str, options: dict):
    """Returns (output_bytes: bytes, output_ext: str) or raises ValueError with a message."""
    src_fmt = src_fmt.lower()
    target_fmt = target_fmt.lower()
    base = filename.rsplit(".", 1)[0]

    # ---------------- PDF as source ----------------
    if src_fmt == "pdf":
        if pdfc.is_scanned_pdf(file_bytes) and target_fmt in ("txt", "docx", "html", "md") and options.get("ocr_if_scanned", True):
            lang = OCR_LANGUAGES.get(options.get("ocr_lang", "English"), "eng")
            searchable = pdfc.searchable_pdf_from_ocr(file_bytes, lang=lang)
            file_bytes = searchable  # now has a real text layer

        if target_fmt == "docx":
            return pdfc.pdf_to_docx(file_bytes, options.get("preserve_layout", True), options.get("extract_images", False)), "docx"
        if target_fmt == "jpg":
            pages = pdfc.pdf_to_images(file_bytes, dpi=options.get("dpi", 150), fmt="jpg", page_range=options.get("page_range"))
            return pages, "jpg"  # list -> caller zips if >1
        if target_fmt == "png":
            pages = pdfc.pdf_to_images(file_bytes, dpi=options.get("dpi", 150), fmt="png", page_range=options.get("page_range"))
            return pages, "png"
        if target_fmt == "txt":
            return pdfc.pdf_to_text(file_bytes).encode("utf-8"), "txt"
        if target_fmt == "html":
            return pdfc.pdf_to_html(file_bytes).encode("utf-8"), "html"
        if target_fmt == "md":
            return pdfc.pdf_to_markdown(file_bytes).encode("utf-8"), "md"
        if target_fmt == "csv":
            return pdfc.pdf_to_csv(file_bytes).encode("utf-8"), "csv"
        if target_fmt == "pptx":
            return pdfc.pdf_to_pptx(file_bytes), "pptx"

    # ---------------- DOCX as source ----------------
    if src_fmt == "docx":
        if target_fmt == "pdf":
            return docxc.docx_to_pdf(file_bytes), "pdf"
        if target_fmt == "txt":
            return docxc.docx_to_text(file_bytes).encode("utf-8"), "txt"
        if target_fmt == "html":
            return docxc.docx_to_html(file_bytes).encode("utf-8"), "html"
        if target_fmt == "md":
            return docxc.docx_to_markdown(file_bytes).encode("utf-8"), "md"

    # ---------------- XLSX as source ----------------
    if src_fmt == "xlsx":
        if target_fmt == "pdf":
            return xlsxc.xlsx_to_pdf(file_bytes), "pdf"
        if target_fmt == "csv":
            return xlsxc.xlsx_to_csv(file_bytes).encode("utf-8"), "csv"
        if target_fmt == "txt":
            return xlsxc.xlsx_to_text(file_bytes).encode("utf-8"), "txt"
        if target_fmt == "html":
            return xlsxc.xlsx_to_html(file_bytes).encode("utf-8"), "html"

    # ---------------- PPTX as source ----------------
    if src_fmt == "pptx":
        if target_fmt == "pdf":
            return pptxc.pptx_to_pdf(file_bytes), "pdf"
        if target_fmt in ("jpg", "png"):
            return pptxc.pptx_to_images(file_bytes, fmt=target_fmt), target_fmt
        if target_fmt == "txt":
            return pptxc.pptx_to_text(file_bytes).encode("utf-8"), "txt"

    # ---------------- Images as source ----------------
    if src_fmt in ("jpg", "jpeg", "png", "webp"):
        if target_fmt == "pdf":
            return pdfc.images_to_pdf(
                [file_bytes], options.get("page_size", "Original"),
                options.get("orientation", "Portrait"), options.get("margin_pt", 0),
                options.get("fit", "Fit"),
            ), "pdf"
        if target_fmt in ("jpg", "png", "webp"):
            return imgc.convert_image(file_bytes, target_fmt, options.get("quality", 90)), target_fmt
        if target_fmt == "txt":
            lang = OCR_LANGUAGES.get(options.get("ocr_lang", "English"), "eng")
            return imgc.image_to_text_via_ocr(file_bytes, lang=lang).encode("utf-8"), "txt"

    # ---------------- Text as source ----------------
    if src_fmt == "txt":
        text = file_bytes.decode("utf-8", errors="replace")
        if target_fmt == "pdf":
            return txtc.text_to_pdf(text), "pdf"
        if target_fmt == "docx":
            return txtc.text_to_docx(text), "docx"
        if target_fmt == "html":
            return txtc.text_to_html(text).encode("utf-8"), "html"
        if target_fmt == "md":
            return txtc.text_to_markdown(text).encode("utf-8"), "md"

    # ---------------- HTML as source ----------------
    if src_fmt == "html":
        html_content = file_bytes.decode("utf-8", errors="replace")
        if target_fmt == "pdf":
            return txtc.html_to_pdf(html_content), "pdf"
        if target_fmt == "txt":
            return txtc.html_to_text(html_content).encode("utf-8"), "txt"

    # ---------------- Markdown as source ----------------
    if src_fmt == "md":
        md_text = file_bytes.decode("utf-8", errors="replace")
        if target_fmt == "html":
            return txtc.markdown_to_html(md_text).encode("utf-8"), "html"

    raise ValueError(f"Conversion from {src_fmt.upper()} to {target_fmt.upper()} isn't supported.")


def zip_pages(pages, base_name, ext):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for page_no, data in pages:
            zf.writestr(f"{base_name}_page{page_no}.{ext}", data)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Sidebar navigation
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# Sidebar navigation
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🔄 ConvertX")
    page = st.radio(
        "Navigate",
        ["🏠 Convert", "📄 PDF Tools", "🖼️ Batch Images", "🔤 OCR", "❓ FAQ / How it works"],
        label_visibility="collapsed",
    )
    st.markdown("---")

    # --- ADVERTISEMENT SECTION ---
    st.caption("ADVERTISEMENT")
    with st.container(border=True):
        st.markdown("**📢 Advertise Here**")
        st.caption("Reach active users looking to convert and edit files.")
        st.markdown("[Contact for Ads](mailto:patilchirantan71@gmail.com)")
# --------------------------------------------------------------------------
# PAGE: Convert (home)
# --------------------------------------------------------------------------
if page == "🏠 Convert":
    st.markdown(
        "<div class='cx-hero'><h1>Convert Anything. Instantly.</h1>"
        "<div class='cx-tagline'>Convert PDFs, documents, images, spreadsheets and more into the "
        "format you need. Fast, simple and completely free.</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<div style='text-align:center'>"
        "<span class='cx-badge'>PDF</span><span class='cx-badge'>DOCX</span>"
        "<span class='cx-badge'>XLSX</span><span class='cx-badge'>PPTX</span>"
        "<span class='cx-badge'>JPG</span><span class='cx-badge'>PNG</span>"
        "<span class='cx-badge'>WEBP</span><span class='cx-badge'>TXT</span>"
        "<span class='cx-badge'>HTML</span><span class='cx-badge'>Markdown</span>"
        "<span class='cx-badge'>CSV</span></div>",
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader("Drop your file here, or Choose File", key="main_upload")

    if uploaded:
        file_bytes = uploaded.getvalue()
        safe_name = sec.sanitize_filename(uploaded.name)
        ext = sec.get_extension(safe_name)

        if not sec.validate_file_size(file_bytes, MAX_FILE_MB):
            error_box(f"That file is over the {MAX_FILE_MB} MB limit.")
        elif not sec.validate_extension(safe_name):
            error_box("That file type isn't supported.")
        elif not sec.validate_magic_bytes(file_bytes, ext):
            error_box("This file doesn't look like a valid file of that type. It may be corrupted or renamed.")
        else:
            real_fmt = sec.detect_real_format(safe_name, file_bytes)
            col1, col2 = st.columns([1, 1])

            with col1:
                st.markdown("#### Your File")
                st.markdown(f"""
                <div class='cx-card'>
                📄 <b>{safe_name}</b><br>
                Type: {real_fmt.upper()}<br>
                Size: {sizeof(len(file_bytes))}
                </div>
                """, unsafe_allow_html=True)
                if real_fmt in ("jpg", "jpeg", "png", "webp"):
                    st.image(file_bytes, use_container_width=True)
                elif real_fmt == "pdf":
                    try:
                        thumb = pdfc.pdf_to_images(file_bytes, dpi=90, page_range=[0])[0][1]
                        st.image(thumb, caption=f"{pdfc.pdf_page_count(file_bytes)} page(s)", use_container_width=True)
                    except Exception:
                        pass

            with col2:
                st.markdown("#### Convert To")
                targets = FORMAT_TARGETS.get(real_fmt, [])
                if not targets:
                    st.warning("No supported conversions for this file type yet.")
                else:
                    target_fmt = st.selectbox("Output format", [t.upper() for t in targets]).lower()

                    options = {}
                    st.markdown("##### Options")
                    if real_fmt == "pdf" and target_fmt in ("jpg", "png"):
                        options["dpi"] = st.slider("DPI / image quality", 72, 300, 150, step=12)
                    if real_fmt == "pdf" and target_fmt == "docx":
                        options["preserve_layout"] = st.checkbox("Preserve layout", True)
                        options["extract_images"] = st.checkbox("Extract images", False)
                    if real_fmt == "pdf" and target_fmt in ("txt", "docx", "html", "md"):
                        options["ocr_if_scanned"] = st.checkbox("Auto-OCR if scanned document", True)
                        if options["ocr_if_scanned"]:
                            options["ocr_lang"] = st.selectbox("OCR language", list(OCR_LANGUAGES.keys()))
                    if real_fmt in ("jpg", "jpeg", "png", "webp") and target_fmt == "pdf":
                        options["page_size"] = st.selectbox("Page size", ["Original", "A4", "Letter"])
                        options["orientation"] = st.selectbox("Orientation", ["Portrait", "Landscape"])
                        options["fit"] = st.selectbox("Image fit", ["Fit", "Fill", "Stretch"])
                    if real_fmt in ("jpg", "jpeg", "png", "webp") and target_fmt in ("jpg", "webp"):
                        options["quality"] = st.slider("Quality", 30, 100, 90)
                    if real_fmt in ("jpg", "jpeg", "png", "webp") and target_fmt == "txt":
                        options["ocr_lang"] = st.selectbox("OCR language", list(OCR_LANGUAGES.keys()))

                    if st.button("Convert Now", type="primary", use_container_width=True):
                        steps = ["Uploading...", "Analyzing file...", "Converting...", "Finalizing..."]
                        run_progress(steps)
                        try:
                            with st.spinner("Converting..."):
                                result, out_ext = convert(safe_name, file_bytes, real_fmt, target_fmt, options)

                            if isinstance(result, list):  # multi-page image output
                                if len(result) == 1:
                                    out_bytes = result[0][1]
                                    out_name = f"{safe_name.rsplit('.', 1)[0]}.{out_ext}"
                                else:
                                    out_bytes = zip_pages(result, safe_name.rsplit(".", 1)[0], out_ext)
                                    out_name = f"{safe_name.rsplit('.', 1)[0]}_pages.zip"
                                    out_ext = "zip"
                            else:
                                out_bytes = result
                                out_name = f"{safe_name.rsplit('.', 1)[0]}.{out_ext}"

                            st.success("Conversion Complete ✓")
                            c1, c2 = st.columns(2)
                            c1.markdown(f"**Original:**\n\n{safe_name}")
                            c2.markdown(f"**Converted:**\n\n{out_name}")
                            st.download_button(
                                "⬇️ Download File", out_bytes, file_name=out_name,
                                mime=MIME_TYPES.get(out_ext, "application/octet-stream"),
                                use_container_width=True,
                            )
                        except ValueError as e:
                            error_box(str(e))
                        except Exception:
                            error_box()

    st.markdown("<div class='cx-privacy'>Your files are automatically deleted after processing.</div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# PAGE: PDF Tools
# --------------------------------------------------------------------------
elif page == "📄 PDF Tools":
    st.markdown("## 📄 PDF Tools")
    tool = st.selectbox(
        "Choose a tool",
        ["Merge PDF", "Split PDF", "Compress PDF", "Rotate PDF", "Extract Pages",
         "Delete Pages", "Reorder Pages", "Add Watermark", "Add Page Numbers"],
    )

    if tool == "Merge PDF":
        files = st.file_uploader("Upload PDFs to merge (in order)", type=["pdf"], accept_multiple_files=True)
        if files and len(files) >= 2 and st.button("Merge", type="primary"):
            try:
                out = pdfc.merge_pdfs([f.getvalue() for f in files])
                st.success("Merged ✓")
                st.download_button("⬇️ Download merged.pdf", out, "merged.pdf", "application/pdf")
            except Exception:
                error_box()
        elif files and len(files) < 2:
            st.info("Upload at least 2 PDFs.")

    elif tool == "Split PDF":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            every = st.number_input("Pages per split file", min_value=1, value=1)
            if st.button("Split", type="primary"):
                try:
                    parts = pdfc.split_pdf(f.getvalue(), int(every))
                    if len(parts) == 1:
                        st.download_button("⬇️ Download", parts[0][1], f"{parts[0][0]}.pdf", "application/pdf")
                    else:
                        buf = io.BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for label, data in parts:
                                zf.writestr(f"{label}.pdf", data)
                        st.success(f"Split into {len(parts)} files ✓")
                        st.download_button("⬇️ Download All (ZIP)", buf.getvalue(), "split_pages.zip", "application/zip")
                except Exception:
                    error_box()

    elif tool == "Compress PDF":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            q = st.slider("Image quality (lower = smaller file)", 10, 95, 60)
            if st.button("Compress", type="primary"):
                try:
                    orig_size = len(f.getvalue())
                    out = pdfc.compress_pdf(f.getvalue(), q)
                    st.success(f"Compressed ✓  {sizeof(orig_size)} → {sizeof(len(out))}")
                    st.download_button("⬇️ Download compressed.pdf", out, "compressed.pdf", "application/pdf")
                except Exception:
                    error_box()

    elif tool == "Rotate PDF":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            deg = st.selectbox("Rotate by", [90, 180, 270])
            if st.button("Rotate", type="primary"):
                try:
                    out = pdfc.rotate_pdf(f.getvalue(), deg)
                    st.success("Rotated ✓")
                    st.download_button("⬇️ Download rotated.pdf", out, "rotated.pdf", "application/pdf")
                except Exception:
                    error_box()

    elif tool in ("Extract Pages", "Delete Pages"):
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            n = pdfc.pdf_page_count(f.getvalue())
            st.caption(f"This PDF has {n} pages.")
            pages_str = st.text_input("Page numbers (e.g. 1,3,5-7)", "1")
            if st.button(tool, type="primary"):
                try:
                    idxs = []
                    for part in pages_str.split(","):
                        part = part.strip()
                        if "-" in part:
                            a, b = part.split("-")
                            idxs.extend(range(int(a) - 1, int(b)))
                        elif part:
                            idxs.append(int(part) - 1)
                    if tool == "Extract Pages":
                        out = pdfc.extract_pages(f.getvalue(), idxs)
                        fname = "extracted.pdf"
                    else:
                        out = pdfc.delete_pages(f.getvalue(), idxs)
                        fname = "pages_deleted.pdf"
                    st.success("Done ✓")
                    st.download_button(f"⬇️ Download {fname}", out, fname, "application/pdf")
                except Exception:
                    error_box()

    elif tool == "Reorder Pages":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            n = pdfc.pdf_page_count(f.getvalue())
            st.caption(f"This PDF has {n} pages. Enter the new order, e.g. 3,1,2")
            order_str = st.text_input("New page order", ",".join(str(i) for i in range(1, n + 1)))
            if st.button("Reorder", type="primary"):
                try:
                    new_order = [int(x.strip()) - 1 for x in order_str.split(",") if x.strip()]
                    out = pdfc.reorder_pages(f.getvalue(), new_order)
                    st.success("Reordered ✓")
                    st.download_button("⬇️ Download reordered.pdf", out, "reordered.pdf", "application/pdf")
                except Exception:
                    error_box()

    elif tool == "Add Watermark":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            text = st.text_input("Watermark text", "CONFIDENTIAL")
            opacity = st.slider("Opacity", 0.05, 1.0, 0.3)
            if st.button("Add Watermark", type="primary"):
                try:
                    out = pdfc.add_watermark(f.getvalue(), text, opacity)
                    st.success("Watermark added ✓")
                    st.download_button("⬇️ Download watermarked.pdf", out, "watermarked.pdf", "application/pdf")
                except Exception:
                    error_box()

    elif tool == "Add Page Numbers":
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            start_at = st.number_input("Start numbering at", min_value=1, value=1)
            pos = st.selectbox("Position", ["bottom-center", "bottom-right", "bottom-left"])
            if st.button("Add Page Numbers", type="primary"):
                try:
                    out = pdfc.add_page_numbers(f.getvalue(), int(start_at), pos)
                    st.success("Page numbers added ✓")
                    st.download_button("⬇️ Download numbered.pdf", out, "numbered.pdf", "application/pdf")
                except Exception:
                    error_box()

# --------------------------------------------------------------------------
# PAGE: Batch Images
# --------------------------------------------------------------------------
elif page == "🖼️ Batch Images":
    st.markdown("## 🖼️ Batch Image Conversion")
    mode = st.radio("Mode", ["Combine images into one PDF", "Convert each image independently"], horizontal=True)
    files = st.file_uploader("Upload images", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)

    if files:
        st.caption(f"{len(files)} image(s) uploaded.")
        names = [f.name for f in files]

        if mode == "Combine images into one PDF":
            order_str = st.text_input(
                "Order (comma-separated filenames, edit to reorder)",
                ", ".join(names),
            )
            page_size = st.selectbox("Page size", ["Original", "A4", "Letter"])
            orientation = st.selectbox("Orientation", ["Portrait", "Landscape"])
            fit = st.selectbox("Image fit", ["Fit", "Fill", "Stretch"])
            if st.button("Convert to PDF", type="primary"):
                try:
                    order = [n.strip() for n in order_str.split(",") if n.strip()]
                    by_name = {f.name: f.getvalue() for f in files}
                    ordered_bytes = [by_name[n] for n in order if n in by_name]
                    out = pdfc.images_to_pdf(ordered_bytes, page_size, orientation, 0, fit)
                    st.success("Converted ✓")
                    st.download_button("⬇️ Download converted_images.pdf", out, "converted_images.pdf", "application/pdf")
                except Exception:
                    error_box()

        else:
            target_fmt = st.selectbox("Convert all to", ["PDF", "JPG", "PNG", "WEBP"]).lower()
            if st.button("Convert All", type="primary"):
                try:
                    buf = io.BytesIO()
                    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in files:
                            base = f.name.rsplit(".", 1)[0]
                            if target_fmt == "pdf":
                                out = pdfc.images_to_pdf([f.getvalue()])
                            else:
                                out = imgc.convert_image(f.getvalue(), target_fmt)
                            zf.writestr(f"{base}.{target_fmt}", out)
                    st.success("Converted ✓")
                    st.download_button("⬇️ Download All (ZIP)", buf.getvalue(), "converted_images.zip", "application/zip")
                except Exception:
                    error_box()

# --------------------------------------------------------------------------
# PAGE: OCR
# --------------------------------------------------------------------------
elif page == "🔤 OCR":
    st.markdown("## 🔤 OCR - Extract Text from Images & Scanned PDFs")
    st.caption("Powered by Tesseract OCR (free & open-source).")
    ocr_mode = st.radio(
        "What do you want to do?",
        ["Image → Text", "Image → Searchable PDF", "Scanned PDF → Searchable PDF", "Scanned PDF → Text"],
    )
    lang_label = st.selectbox("Language", list(OCR_LANGUAGES.keys()))
    lang = OCR_LANGUAGES[lang_label]

    if ocr_mode.startswith("Image"):
        f = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png", "webp"])
        if f and st.button("Run OCR", type="primary"):
            try:
                with st.spinner("Running OCR..."):
                    if ocr_mode == "Image → Text":
                        text = imgc.image_to_text_via_ocr(f.getvalue(), lang)
                        st.success("Done ✓")
                        st.text_area("Extracted text", text, height=250)
                        st.download_button("⬇️ Download .txt", text.encode("utf-8"), "extracted.txt", "text/plain")
                    else:
                        pdf_bytes = pdfc._ocr_image_to_pdf_bytes(f.getvalue(), lang)
                        st.success("Done ✓")
                        st.download_button("⬇️ Download searchable.pdf", pdf_bytes, "searchable.pdf", "application/pdf")
            except Exception:
                error_box()
    else:
        f = st.file_uploader("Upload a PDF", type=["pdf"])
        if f:
            scanned = pdfc.is_scanned_pdf(f.getvalue())
            st.caption("🔍 Detected as a **scanned** PDF." if scanned else "ℹ️ This PDF already has selectable text, but OCR will still run if you continue.")
            if st.button("Run OCR", type="primary"):
                try:
                    with st.spinner("Running OCR across all pages..."):
                        out_pdf = pdfc.searchable_pdf_from_ocr(f.getvalue(), lang)
                    if ocr_mode == "Scanned PDF → Searchable PDF":
                        st.success("Done ✓")
                        st.download_button("⬇️ Download searchable.pdf", out_pdf, "searchable.pdf", "application/pdf")
                    else:
                        text = pdfc.pdf_to_text(out_pdf)
                        st.success("Done ✓")
                        st.text_area("Extracted text", text, height=250)
                        st.download_button("⬇️ Download .txt", text.encode("utf-8"), "extracted.txt", "text/plain")
                except Exception:
                    error_box()

# --------------------------------------------------------------------------
# PAGE: FAQ
# --------------------------------------------------------------------------
elif page == "❓ FAQ / How it works":
    st.markdown("## How ConvertX works")
    st.markdown("""
    1. **Upload** a file (drag & drop or choose a file).
    2. ConvertX **detects its real format** by reading its file signature - not just its extension.
    3. You **choose an output format** from the ones that are actually compatible.
    4. ConvertX **runs a real conversion** using open-source engines (PyMuPDF, LibreOffice-compatible
       tooling where available, Pillow, Tesseract OCR, reportlab).
    5. You **download the result**. Files are never stored beyond your session.
    """)
    st.markdown("## FAQ")
    with st.expander("Is ConvertX really free?"):
        st.write("Yes. No login, no payment, no subscriptions, no credits, no paid conversion APIs.")
    with st.expander("Do you store my files?"):
        st.write("No. Files are processed in memory for your session only and are never written to permanent storage or used to train AI models.")
    with st.expander("What's the maximum file size?"):
        st.write(f"{MAX_FILE_MB} MB per file by default. This is configurable by whoever deploys the app.")
    with st.expander("Why isn't my exact conversion listed?"):
        st.write("ConvertX only advertises conversions it can genuinely perform end-to-end. See the README for the full supported matrix.")
    with st.expander("How does OCR language support work?"):
        st.write("OCR is powered by Tesseract with language packs for English, Hindi, Marathi, Gujarati, Bengali, Tamil, Telugu, Kannada, Malayalam and Punjabi.")

st.markdown("---")

st.caption(
    "ConvertX · Convert Anything. Instantly. · No login · No payment · "
    "No subscriptions · Free & open-source engines"
)

import base64

with open("profile.png", "rb") as f:
    profile_image = base64.b64encode(f.read()).decode()

st.markdown(
    f"""
    <div style="text-align:center; margin-top:20px;">
        <img src="data:image/png;base64,{profile_image}"
             style="
                width:90px;
                height:90px;
                object-fit:cover;
                border-radius:50%;
                border:3px solid #ddd;
             ">
        <div style="color:#777; font-size:14px; margin-top:10px;">
            A free initiative to make digitalisation easier by
            <br>
            <b style="color:#444;">Chirantan Patil</b>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)
