import io
from PIL import Image

def convert_image(data, target_fmt, quality=90):
    img = Image.open(io.BytesIO(data))
    target_fmt = target_fmt.lower()
    out = io.BytesIO()
    if target_fmt in ("jpg","jpeg"):
        if img.mode in ("RGBA","LA","P"):
            bg = Image.new("RGB", img.size, "white")
            if img.mode != "RGB":
                img = img.convert("RGBA")
                bg.paste(img, mask=img.getchannel("A"))
                img = bg
        else:
            img = img.convert("RGB")
        img.save(out, format="JPEG", quality=quality, optimize=True)
    elif target_fmt == "png":
        img.save(out, format="PNG", optimize=True)
    elif target_fmt == "webp":
        img.save(out, format="WEBP", quality=quality)
    else:
        raise ValueError(f"Unsupported image output: {target_fmt}")
    return out.getvalue()

def image_to_text_via_ocr(data, lang="eng"):
    import pytesseract
    img = Image.open(io.BytesIO(data))
    return pytesseract.image_to_string(img, lang=lang)
