# -*- coding: utf-8 -*-
"""استخراج نص من صورة (سكرين شوت طلب) عبر Tesseract OCR — اختياري، لا يكسر البرنامج إن لم يتوفر."""

def ocr_image(path):
    """يرجّع (text, error). إذا OCR غير متاح، يرجّع نصاً فارغاً ورسالة توضيحية."""
    try:
        import pytesseract
        from PIL import Image
    except Exception:
        return "", ("OCR غير مثبّت. للتفعيل، نفّذ على الماك: "
                     "brew install tesseract tesseract-lang && "
                     "source .venv/bin/activate && pip install pytesseract pillow — "
                     "أو الصق نص الطلب يدوياً بالخانة أدناه (أدق وأسرع).")
    try:
        img = Image.open(path)
        text = pytesseract.image_to_string(img, lang="eng+tur+ara")
        if not text.strip():
            return "", "لم يُستخرج أي نص من الصورة — جرّب صورة أوضح أو الصق النص يدوياً."
        return text, None
    except Exception as e:
        msg = str(e)
        if "tesseract is not installed" in msg.lower() or "not found" in msg.lower():
            return "", ("محرّك Tesseract غير مثبّت على الجهاز. نفّذ: "
                         "brew install tesseract tesseract-lang — أو الصق نص الطلب يدوياً.")
        return "", f"فشل استخراج النص: {msg}"
