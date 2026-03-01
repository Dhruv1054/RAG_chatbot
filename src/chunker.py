import json
from dataclasses import dataclass, asdict
from typing import List
from langchain.text_splitter import RecursiveCharacterTextSplitter # splits large text into chunks
from src.config import CHUNKS_DIR, CHUNK_SIZE, CHUNK_OVERLAP


@dataclass  # defines what one chunk contains
class Chunk:
    chunk_id:   int
    text:       str
    char_start: int
    char_end:   int
    word_count: int


def _build_splitter():  # Splitter object created 
    return RecursiveCharacterTextSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP,
        separators=["\n\n", "\n", ". ", "? ", "! ", ", ", " ", ""],  # Tells the langchain to split in this order, thats why recursive 
        length_function=len,
        is_separator_regex=False,
    )


def _compute_char_offsets(full_text, chunks):
    offsets = []
    cursor = 0
    for chunk in chunks:
        idx = full_text.find(chunk.strip()[:30], cursor)
        if idx == -1:
            idx = cursor
        start = idx
        end   = start + len(chunk)
        offsets.append((start, end))
        cursor = max(cursor + 1, start + CHUNK_SIZE - CHUNK_OVERLAP)
    return offsets


def split_into_chunks(text): # Main function (runs when document is uploaded)
    if not text or not text.strip():
        return []
    splitter = _build_splitter() #splits text using splitter
    raw_chunks = splitter.split_text(text)
    offsets = _compute_char_offsets(text, raw_chunks) 
    chunks = []  #
    for idx, (raw, (start, end)) in enumerate(zip(raw_chunks, offsets)):
        chunks.append(Chunk(
            chunk_id=idx,
            text=raw.strip(),
            char_start=start,
            char_end=end,
            word_count=len(raw.split()),
        ))
    return chunks


def save_chunks(chunks, doc_name="document"):
    doc_dir = CHUNKS_DIR / doc_name
    doc_dir.mkdir(parents=True, exist_ok=True)
    for old in doc_dir.glob("*"):
        old.unlink()
    manifest = []
    for chunk in chunks:
        chunk_file = doc_dir / f"chunk_{chunk.chunk_id:04d}.txt"
        chunk_file.write_text(chunk.text, encoding="utf-8")
        manifest.append(asdict(chunk))
    manifest_file = doc_dir / "manifest.json"
    manifest_file.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return doc_dir


def load_chunks_from_manifest(doc_name="document"):
    manifest_file = CHUNKS_DIR / doc_name / "manifest.json"
    if not manifest_file.exists():
        return []
    data = json.loads(manifest_file.read_text(encoding="utf-8"))
    return [Chunk(**item) for item in data]
