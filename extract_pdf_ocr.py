import pymupdf as fitz
import pytesseract
from PIL import Image
import io
import re

PDF_PATH = "assets/test_week2.pdf"

TEMPLATE_TITLES = {"title/course code", "equipment/name/use"}

def normalize(s):
    return re.sub(r"\s+", " ", s).strip().lower()

def split_entries(text):
    # New entry begins at any line containing "| REAL TIME" (tolerates OCR noise)
    parts = re.split(r"(?mi)^(?=.*\|\s*REAL\s*TIME)", text)
    return [p.strip() for p in parts if re.search(r"REAL\s*TIME", p, re.I)]

def is_template(entry):
    header = entry.splitlines()[0]
    title, _, time_part = header.partition("|")
    return (
        normalize(title) in TEMPLATE_TITLES
        or re.search(r"x{2}\s*[:;.]?\s*x{2}", time_part, re.I) is not None
    )

# 1. OCR every page into one string
pages_text = []
with fitz.open(PDF_PATH) as doc:
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        pages_text.append(pytesseract.image_to_string(img))

text = "\n".join(pages_text)

# 2. Split into entries and drop templates
entries = [e for e in split_entries(text) if not is_template(e)]

# 3. Show what's left
for e in entries:
    print(e)
    print("-" * 40)

print(f"\nKept {len(entries)} real entries.")