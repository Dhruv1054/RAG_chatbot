# Text cleaner before processing - Raw pdfs are messy
import re
import unicodedata
from src.config import DATA_DIR


def _normalise_unicode(text):
    return unicodedata.normalize("NFKD", text)


def _remove_page_numbers(text):
    text = re.sub(r"(?i)\bpage\s+\d+(\s+of\s+\d+)?\b", "", text)
    text = re.sub(r"[-~]{1,3}\s*\d+\s*[-~]{1,3}", "", text)
    text = re.sub(r"(?m)^\s*\d+\s*$", "", text)
    return text


def _remove_headers_footers(text):
    lines = text.split("\n")
    line_freq = {}
    for line in lines:
        stripped = line.strip()
        if stripped and len(stripped.split()) <= 8:
            line_freq[stripped] = line_freq.get(stripped, 0) + 1
    repeated = {line for line, count in line_freq.items() if count > 2}
    cleaned = [line for line in lines if line.strip() not in repeated]
    return "\n".join(cleaned)


def _collapse_whitespace(text):
    text = re.sub(r"[ \t]+$", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return text.strip()


def _extract_text_from_pdf_bytes(pdf_bytes):
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        pages = [page.get_text("text") for page in doc]
        doc.close()
        return "\n".join(pages)
    except ImportError:
        pass
    try:
        import io
        from pdfminer.high_level import extract_text as pdfminer_extract
        return pdfminer_extract(io.BytesIO(pdf_bytes))
    except ImportError:
        raise RuntimeError("No PDF parser found. Install pymupdf.")


def preprocess_text(raw_text):
    text = _normalise_unicode(raw_text)
    text = _remove_page_numbers(text)
    text = _remove_headers_footers(text)
    text = _collapse_whitespace(text)
    return text


def preprocess_pdf(pdf_bytes):     # Extract  raw text from pdf  using fitzz
    raw_text = _extract_text_from_pdf_bytes(pdf_bytes)
    return preprocess_text(raw_text) #clean the extracted text


def save_cleaned_text(text, filename="cleaned.txt"):
    dest = DATA_DIR / filename
    dest.write_text(text, encoding="utf-8")
    return dest


def load_cleaned_text(filename="cleaned.txt"):
    dest = DATA_DIR / filename
    if dest.exists():
        return dest.read_text(encoding="utf-8")
    return None
