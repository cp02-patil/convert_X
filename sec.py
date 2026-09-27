import os, re, zipfile, io

ALLOWED_EXTENSIONS = {
    "pdf","docx","xlsx","pptx","jpg","jpeg","png","webp","txt","html","md","csv"
}

MAGIC = {
    "pdf": b"%PDF",
    "zip": b"PK",
    "jpg": b"\xff\xd8\xff",
    "png": b"\x89PNG\r\n\x1a\n",
    "webp": b"RIFF",
}

def sanitize_filename(name):
    name = os.path.basename(name or "file")
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name)
    return name or "file"

def get_extension(name):
    return name.rsplit(".", 1)[1].lower() if "." in name else ""

def validate_file_size(data, max_mb):
    return len(data) <= max_mb * 1024 * 1024

def validate_extension(name):
    return get_extension(name) in ALLOWED_EXTENSIONS

def validate_magic_bytes(data, ext):
    ext = ext.lower()
    if ext in ("docx","xlsx","pptx"):
        return data[:2] == b"PK"
    if ext == "csv":
        return True
    if ext in ("txt","html","md"):
        return True
    magic = MAGIC.get(ext)
    if not magic:
        return True
    return data.startswith(magic)

def detect_real_format(name, data):
    ext = get_extension(name)
    if data.startswith(b"%PDF"):
        return "pdf"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return "webp"
    if data.startswith(b"PK"):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                names = z.namelist()
                if any(n.startswith("word/") for n in names): return "docx"
                if any(n.startswith("xl/") for n in names): return "xlsx"
                if any(n.startswith("ppt/") for n in names): return "pptx"
        except Exception:
            pass
    return ext
