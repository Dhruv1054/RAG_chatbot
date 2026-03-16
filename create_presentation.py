"""
Script to generate a PowerPoint presentation for the RAG Chatbot project.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ── Color palette ──────────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x0D, 0x1B, 0x2A)   # deep navy
ACCENT     = RGBColor(0x00, 0xB4, 0xD8)   # electric cyan
ACCENT2    = RGBColor(0x90, 0xE0, 0xEF)   # light cyan
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xCC, 0xD6, 0xE0)
HIGHLIGHT  = RGBColor(0xFF, 0x6B, 0x6B)   # coral red for emphasis
GREEN      = RGBColor(0x06, 0xD6, 0xA0)   # mint green
CARD_BG    = RGBColor(0x1A, 0x2E, 0x45)   # slightly lighter navy for cards

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]  # fully blank layout


# ── Helpers ────────────────────────────────────────────────────────────────────

def add_bg(slide, color=DARK_BG):
    """Fill slide background with a solid color."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def box(slide, l, t, w, h, color, alpha=None):
    """Add a filled rectangle (no border)."""
    shape = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO_SHAPE,  # placeholder — overridden
        Inches(l), Inches(t), Inches(w), Inches(h)
    )
    # use add_shape with autoshape id 1 (rectangle)
    return shape


def rect(slide, l, t, w, h, fill_color, border_color=None, border_pt=0):
    """Add a rectangle with optional border."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    sp = slide.shapes.add_shape(1, Inches(l), Inches(t), Inches(w), Inches(h))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill_color
    if border_color and border_pt:
        sp.line.color.rgb = border_color
        sp.line.width = Pt(border_pt)
    else:
        sp.line.fill.background()
    return sp


def txt(slide, text, l, t, w, h,
        font_size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
        italic=False, wrap=True):
    """Add a text box."""
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb


def accent_bar(slide, t=0.55, h=0.06):
    """Horizontal accent bar below the title area."""
    rect(slide, 0, t, 13.33, h, ACCENT)


def slide_header(slide, title, subtitle=None):
    """Standard slide header with title + optional subtitle."""
    rect(slide, 0, 0, 13.33, 1.2, CARD_BG)
    accent_bar(slide, t=1.2, h=0.06)
    txt(slide, title, 0.5, 0.18, 12.3, 0.7,
        font_size=28, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        txt(slide, subtitle, 0.5, 0.78, 12.3, 0.4,
            font_size=14, color=ACCENT2, align=PP_ALIGN.LEFT)


def bullet_card(slide, items, l, t, w, h, icon="▶"):
    """A card containing a bullet list."""
    rect(slide, l, t, w, h, CARD_BG, ACCENT, 1)
    y = t + 0.15
    for item in items:
        txt(slide, f"{icon}  {item}", l + 0.18, y, w - 0.36, 0.38,
            font_size=14, color=LIGHT_GRAY)
        y += 0.42


def section_label(slide, text, l, t, w=3):
    """Small cyan label above a section."""
    txt(slide, text.upper(), l, t, w, 0.3,
        font_size=10, bold=True, color=ACCENT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title / Hero
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)

# gradient-like panel on left
rect(s, 0, 0, 5.8, 7.5, CARD_BG)
rect(s, 0, 0, 0.18, 7.5, ACCENT)   # cyan left stripe

txt(s, "RAG", 0.55, 1.1, 5.0, 1.4,
    font_size=90, bold=True, color=ACCENT, align=PP_ALIGN.LEFT)
txt(s, "CHATBOT", 0.55, 2.3, 5.0, 0.9,
    font_size=42, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
txt(s, "Retrieval-Augmented Generation\nfor Intelligent Document Q&A",
    0.55, 3.25, 5.0, 1.0,
    font_size=15, color=ACCENT2, align=PP_ALIGN.LEFT)

# right panel — key stats
for i, (val, lbl) in enumerate([
    ("Local LLM", "Powered by Ollama + LLaMA 3"),
    ("FAISS", "Vector Similarity Search"),
    ("LangGraph", "Orchestration Pipeline"),
    ("Streamlit", "Real-time Web UI"),
]):
    yt = 1.1 + i * 1.4
    rect(s, 6.3, yt, 6.4, 1.1, RGBColor(0x12, 0x25, 0x3A), ACCENT, 1)
    txt(s, val, 6.6, yt + 0.08, 5.8, 0.5,
        font_size=22, bold=True, color=ACCENT)
    txt(s, lbl, 6.6, yt + 0.55, 5.8, 0.4,
        font_size=12, color=LIGHT_GRAY)

txt(s, "Dhruv Rakhecha  •  March 2026", 0.55, 6.9, 5.0, 0.4,
    font_size=11, color=RGBColor(0x80, 0x99, 0xAA), align=PP_ALIGN.LEFT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Problem Statement
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Problem Statement",
             "Why do we need RAG?")

problems = [
    ("📚", "Knowledge Cutoff",
     "LLMs are trained on static data — they cannot answer\nquestions about new, private, or domain-specific documents."),
    ("🤯", "Hallucination",
     "Without grounding, models confidently generate plausible\nbut incorrect answers, eroding user trust."),
    ("🔒", "Privacy & Cost",
     "Sending sensitive documents to cloud APIs exposes data\nand incurs high per-token API costs at scale."),
    ("🔍", "No Source Citations",
     "Generic chatbots cannot tell you which part of a document\nbacked their answer, making verification impossible."),
]

for i, (icon, title, desc) in enumerate(problems):
    col = i % 2
    row = i // 2
    lx = 0.5 + col * 6.4
    ty = 1.55 + row * 2.6
    rect(s, lx, ty, 6.1, 2.3, CARD_BG, ACCENT if col == 0 else HIGHLIGHT, 1.5)
    txt(s, icon + "  " + title, lx + 0.2, ty + 0.15, 5.7, 0.5,
        font_size=17, bold=True, color=ACCENT if col == 0 else HIGHLIGHT)
    txt(s, desc, lx + 0.2, ty + 0.65, 5.7, 1.4,
        font_size=13, color=LIGHT_GRAY, wrap=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Our Solution
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Our Solution — RAG Chatbot",
             "Ground the LLM in your own documents, locally and privately")

txt(s,
    "We built a fully local Retrieval-Augmented Generation (RAG) system that "
    "ingests user-uploaded documents, converts them into semantic vector embeddings, "
    "and dynamically retrieves the most relevant context before generating an answer — "
    "eliminating hallucinations while keeping all data on-device.",
    0.5, 1.45, 12.3, 1.0,
    font_size=14, color=LIGHT_GRAY, wrap=True)

solutions = [
    ("✅ Grounded Answers",    "Every response is backed by real document chunks — no fabrication."),
    ("✅ Source Citations",    "Users see exactly which passage and similarity score drove the answer."),
    ("✅ 100% Local",          "Ollama + FAISS run entirely on-device — zero cloud API calls."),
    ("✅ Any Document",        "Upload PDFs or text files; the system indexes them in seconds."),
    ("✅ Streaming UI",        "Real-time token-by-token output via Streamlit for a fluid UX."),
    ("✅ Hallucination Guard", "System prompt enforces context-only answers with fallback message."),
]

for i, (title, desc) in enumerate(solutions):
    col = i % 2
    row = i // 2
    lx = 0.5 + col * 6.4
    ty = 2.65 + row * 1.5
    rect(s, lx, ty, 6.1, 1.3, CARD_BG, GREEN, 1)
    txt(s, title, lx + 0.2, ty + 0.1, 5.7, 0.4,
        font_size=15, bold=True, color=GREEN)
    txt(s, desc, lx + 0.2, ty + 0.55, 5.7, 0.6,
        font_size=12, color=LIGHT_GRAY)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Architecture Overview
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "System Architecture", "How all components fit together")

# --- Upload pipeline (left column) ---
rect(s, 0.4, 1.4, 3.5, 5.7, CARD_BG, ACCENT, 1)
txt(s, "📥  INGESTION PIPELINE", 0.6, 1.55, 3.1, 0.4,
    font_size=12, bold=True, color=ACCENT)

for i, step in enumerate([
    "User uploads PDF / TXT",
    "Text extraction (PyMuPDF)",
    "Preprocessing & cleaning",
    "Chunking  (512 chars, 128 overlap)",
    "Sentence-Transformer embedding",
    "FAISS vector index stored to disk",
]):
    ty = 2.1 + i * 0.78
    rect(s, 0.6, ty, 3.1, 0.6, RGBColor(0x0D, 0x23, 0x38))
    txt(s, f"{i+1}. {step}", 0.75, ty + 0.12, 2.9, 0.38,
        font_size=11, color=WHITE)

# arrow between columns
txt(s, "⟷", 4.1, 3.9, 0.7, 0.5, font_size=28, color=ACCENT, align=PP_ALIGN.CENTER)

# --- Query pipeline (right column) ---
rect(s, 5.0, 1.4, 7.9, 5.7, CARD_BG, ACCENT2, 1)
txt(s, "💬  QUERY PIPELINE", 5.2, 1.55, 7.5, 0.4,
    font_size=12, bold=True, color=ACCENT2)

query_steps = [
    ("User types a question",           ACCENT2),
    ("Query → 384-dim embedding",       ACCENT2),
    ("FAISS top-K retrieval (K=4)",     ACCENT2),
    ("Context builder (≤3000 tokens)",  ACCENT2),
    ("Prompt builder injects context",  ACCENT2),
    ("Ollama LLaMA 3 generates answer", GREEN),
    ("Streamlit streams tokens + sources", GREEN),
]

for i, (step, col) in enumerate(query_steps):
    ty = 2.1 + i * 0.72
    rect(s, 5.2, ty, 7.5, 0.55, RGBColor(0x0D, 0x23, 0x38))
    txt(s, f"{i+1}. {step}", 5.35, ty + 0.1, 7.3, 0.38,
        font_size=11, color=col)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Tech Stack
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Technology Stack", "Libraries and frameworks used")

categories = [
    ("Frontend",        ACCENT,    ["Streamlit 1.40", "Real-time streaming", "File upload widget", "Chat history UI"]),
    ("Backend API",     ACCENT2,   ["FastAPI 0.115", "Uvicorn ASGI server", "REST endpoints", "Pydantic validation"]),
    ("RAG Orchestration", GREEN,   ["LangGraph 0.2", "LangChain 0.3", "State-machine pipeline", "Streaming support"]),
    ("Embeddings",      HIGHLIGHT, ["Sentence-Transformers", "all-MiniLM-L6-v2", "384-dim vectors", "CPU-friendly"]),
    ("Vector DB",       RGBColor(0xFF,0xB7,0x00), ["FAISS 1.9", "L2 similarity search", "Persistent index", "Metadata pickle store"]),
    ("LLM Runtime",     RGBColor(0xC7,0x7D,0xFF), ["Ollama local server", "LLaMA 3 8B Q4", "HTTP streaming API", "Temperature = 0.1"]),
]

cols = 3
for i, (cat, color, items) in enumerate(categories):
    col = i % cols
    row = i // cols
    lx = 0.35 + col * 4.32
    ty = 1.5 + row * 2.75
    rect(s, lx, ty, 4.1, 2.5, CARD_BG, color, 1.5)
    txt(s, cat, lx + 0.18, ty + 0.12, 3.74, 0.42,
        font_size=15, bold=True, color=color)
    for j, item in enumerate(items):
        txt(s, "• " + item, lx + 0.18, ty + 0.6 + j * 0.44, 3.74, 0.38,
            font_size=12, color=LIGHT_GRAY)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — RAG Pipeline Deep Dive
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "RAG Pipeline — Deep Dive", "LangGraph state machine with three nodes")

# Draw three pipeline nodes as connected boxes
nodes = [
    ("RETRIEVE", "① RETRIEVE NODE",
     ["Embed user query (384-dim)", "Search FAISS index", "Return top-4 chunks", "Include similarity scores"],
     ACCENT, 0.4),
    ("CONTEXT", "② CONTEXT NODE",
     ["Receive retrieved chunks", "Count estimated tokens", "Truncate to ≤3000 tokens", "Preserve most relevant first"],
     GREEN, 4.55),
    ("GENERATE", "③ GENERATE NODE",
     ["Build system + user prompt", "Inject context window", "Call Ollama LLaMA 3", "Stream tokens to UI"],
     HIGHLIGHT, 8.7),
]

for node_id, title, bullets, color, lx in nodes:
    rect(s, lx, 1.5, 3.9, 5.5, CARD_BG, color, 2)
    rect(s, lx, 1.5, 3.9, 0.7, color)   # colored header bar
    txt(s, title, lx + 0.15, 1.6, 3.6, 0.5,
        font_size=14, bold=True, color=DARK_BG)
    for j, b in enumerate(bullets):
        txt(s, "▸  " + b, lx + 0.2, 2.4 + j * 0.75, 3.5, 0.6,
            font_size=13, color=LIGHT_GRAY)

# Arrows between nodes
for ax in [4.35, 8.5]:
    txt(s, "➤", ax, 3.95, 0.35, 0.5,
        font_size=22, color=ACCENT, align=PP_ALIGN.CENTER)

# State label
txt(s, "LangGraph StateGraph  —  nodes connected by directed edges, supports streaming",
    0.5, 7.1, 12.3, 0.35,
    font_size=12, italic=True, color=RGBColor(0x88, 0xAA, 0xBB), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Key Features
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Key Features", "What makes this chatbot stand out")

features = [
    ("🔄 Streaming Responses",
     "Tokens stream in real-time to the Streamlit UI,\ngiving a ChatGPT-like interactive feel.",
     ACCENT),
    ("📎 Source Citations",
     "Each answer includes chunk previews and\ncosine similarity scores for full transparency.",
     GREEN),
    ("🧩 Modular Architecture",
     "Every component (embedder, retriever, LLM, UI)\nis independently configurable via config.py.",
     ACCENT2),
    ("🛡️ Hallucination Guard",
     "System prompt instructs the model to answer only\nfrom context, with a safe fallback phrase.",
     HIGHLIGHT),
    ("💾 Persistent Index",
     "FAISS index + metadata saved to disk so the\nvector store survives server restarts.",
     RGBColor(0xFF,0xB7,0x00)),
    ("🖥️ Fully Local",
     "No external API calls — Ollama runs LLaMA 3\non your machine; data never leaves the device.",
     RGBColor(0xC7,0x7D,0xFF)),
]

for i, (title, desc, color) in enumerate(features):
    col = i % 2
    row = i // 2
    lx = 0.4 + col * 6.5
    ty = 1.5 + row * 1.9
    rect(s, lx, ty, 6.2, 1.75, CARD_BG, color, 1.5)
    txt(s, title, lx + 0.2, ty + 0.12, 5.8, 0.45,
        font_size=15, bold=True, color=color)
    txt(s, desc, lx + 0.2, ty + 0.65, 5.8, 0.9,
        font_size=12, color=LIGHT_GRAY, wrap=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Demo Flow / User Journey
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "User Journey", "From document upload to grounded answer")

steps = [
    ("1", "Upload Document",
     "User uploads a PDF or TXT file through the Streamlit sidebar.",
     ACCENT),
    ("2", "Automatic Indexing",
     "System preprocesses, chunks, embeds, and stores vectors in FAISS — all in seconds.",
     ACCENT2),
    ("3", "Ask a Question",
     "User types any question in the chat input about the uploaded document.",
     GREEN),
    ("4", "Semantic Search",
     "Query is embedded and top-4 most relevant chunks are retrieved from FAISS.",
     RGBColor(0xFF,0xB7,0x00)),
    ("5", "LLM Generation",
     "LLaMA 3 generates a grounded answer using only the retrieved context.",
     HIGHLIGHT),
    ("6", "Cited Response",
     "Answer streams to the UI along with source chunks and similarity scores.",
     RGBColor(0xC7,0x7D,0xFF)),
]

for i, (num, title, desc, color) in enumerate(steps):
    row = i // 3
    col = i % 3
    lx = 0.35 + col * 4.32
    ty = 1.5 + row * 2.6

    rect(s, lx, ty, 4.1, 2.35, CARD_BG, color, 1)
    # number circle (simulated with bold text)
    rect(s, lx + 0.15, ty + 0.12, 0.55, 0.55, color)
    txt(s, num, lx + 0.15, ty + 0.1, 0.55, 0.55,
        font_size=18, bold=True, color=DARK_BG, align=PP_ALIGN.CENTER)
    txt(s, title, lx + 0.85, ty + 0.15, 3.1, 0.45,
        font_size=14, bold=True, color=color)
    txt(s, desc, lx + 0.2, ty + 0.72, 3.7, 1.45,
        font_size=12, color=LIGHT_GRAY, wrap=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — Technical Specifications
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Technical Specifications", "Configuration and performance parameters")

specs = [
    ("Embedding Model",   "all-MiniLM-L6-v2",              "384-dimensional dense vectors"),
    ("Chunk Size",        "512 characters",                 "With 128-character overlap"),
    ("Vector Search",     "FAISS L2 distance",              "Top-K = 4 chunks retrieved"),
    ("Context Window",    "≤ 3,000 tokens",                 "Intelligent truncation preserves best chunks"),
    ("LLM Model",         "LLaMA 3 8B (Q4_K_M quant)",     "Via Ollama at localhost:11434"),
    ("Temperature",       "0.1",                            "Low randomness for factual answers"),
    ("Max Output",        "1,024 tokens",                   "Per response generation"),
    ("API Server",        "FastAPI on 0.0.0.0:8000",        "REST + streaming endpoints"),
    ("Document Formats",  "PDF + Plain Text",               "PyMuPDF for PDF extraction"),
    ("Persistence",       "FAISS index + metadata.pkl",     "Survives server restarts"),
]

rect(s, 0.4, 1.42, 12.5, 0.42, ACCENT)
txt(s, "Parameter", 0.55, 1.47, 3.0, 0.32, font_size=13, bold=True, color=DARK_BG)
txt(s, "Value", 3.6, 1.47, 3.5, 0.32, font_size=13, bold=True, color=DARK_BG)
txt(s, "Notes", 7.2, 1.47, 5.5, 0.32, font_size=13, bold=True, color=DARK_BG)

for i, (param, val, note) in enumerate(specs):
    bg = CARD_BG if i % 2 == 0 else RGBColor(0x12, 0x25, 0x3A)
    ty = 1.92 + i * 0.5
    rect(s, 0.4, ty, 12.5, 0.48, bg)
    txt(s, param, 0.55, ty + 0.07, 3.0, 0.36, font_size=12, color=ACCENT2)
    txt(s, val,   3.6,  ty + 0.07, 3.5, 0.36, font_size=12, bold=True, color=WHITE)
    txt(s, note,  7.2,  ty + 0.07, 5.5, 0.36, font_size=11, color=LIGHT_GRAY)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — Challenges & Solutions
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Challenges & Solutions",
             "Engineering decisions made during development")

pairs = [
    ("Challenge: Token Limit",
     "LLMs have a fixed context window; too many chunks would exceed it.",
     "Solution",
     "Context node estimates tokens and truncates intelligently, keeping the highest-scoring chunks first."),
    ("Challenge: Hallucination",
     "Models tend to generate confident but wrong answers without grounding.",
     "Solution",
     "System prompt strictly instructs the model to answer only from retrieved context, with a polite fallback."),
    ("Challenge: PDF Noise",
     "PDFs contain page numbers, headers, footers, and Unicode artifacts.",
     "Solution",
     "preprocessing.py uses regex patterns to strip noise before chunking, improving retrieval quality."),
    ("Challenge: Persistence",
     "Re-indexing on every restart is slow and wasteful.",
     "Solution",
     "FAISS index and metadata are serialized to disk (index.faiss + metadata.pkl) for instant reload."),
]

for i, (ch_title, ch_body, sol_title, sol_body) in enumerate(pairs):
    row = i // 2
    col = i % 2
    lx = 0.4 + col * 6.5
    ty = 1.5 + row * 2.7

    # challenge box
    rect(s, lx, ty, 3.0, 2.4, CARD_BG, HIGHLIGHT, 1)
    txt(s, ch_title, lx + 0.15, ty + 0.1, 2.7, 0.42, font_size=12, bold=True, color=HIGHLIGHT)
    txt(s, ch_body,  lx + 0.15, ty + 0.6, 2.7, 1.6,  font_size=11, color=LIGHT_GRAY, wrap=True)

    # solution box
    rect(s, lx + 3.1, ty, 3.0, 2.4, CARD_BG, GREEN, 1)
    txt(s, sol_title, lx + 3.25, ty + 0.1, 2.7, 0.42, font_size=12, bold=True, color=GREEN)
    txt(s, sol_body,  lx + 3.25, ty + 0.6, 2.7, 1.6,  font_size=11, color=LIGHT_GRAY, wrap=True)

    txt(s, "→", lx + 2.98, ty + 0.9, 0.3, 0.5, font_size=18, color=ACCENT, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — Future Enhancements
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)
slide_header(s, "Future Enhancements", "Roadmap for next iterations")

enhancements = [
    ("🔍 Hybrid Search",
     "Combine BM25 keyword search with dense vector retrieval for better coverage of exact-match queries."),
    ("🔄 Multi-Document Sessions",
     "Allow multiple documents to be indexed together and queried simultaneously with per-source attribution."),
    ("🌐 Web Scraping Ingestion",
     "Add a URL input mode that scrapes and indexes web pages alongside uploaded documents."),
    ("🗂️ Conversation Memory",
     "Maintain multi-turn chat context so the model can answer follow-up questions referencing prior exchanges."),
    ("📊 Evaluation Dashboard",
     "Integrate RAG evaluation metrics (faithfulness, answer relevance) using RAGAS or TruLens."),
    ("☁️ Cloud Deployment",
     "Containerize with Docker and deploy to a cloud provider with GPU support for faster inference."),
]

for i, (title, desc) in enumerate(enhancements):
    col = i % 2
    row = i // 2
    lx = 0.4 + col * 6.5
    ty = 1.5 + row * 1.85
    rect(s, lx, ty, 6.2, 1.7, CARD_BG, ACCENT2, 1)
    txt(s, title, lx + 0.2, ty + 0.12, 5.8, 0.45,
        font_size=14, bold=True, color=ACCENT2)
    txt(s, desc, lx + 0.2, ty + 0.62, 5.8, 0.9,
        font_size=12, color=LIGHT_GRAY, wrap=True)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 12 — Thank You / Q&A
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_bg(s)

rect(s, 0, 0, 13.33, 7.5, DARK_BG)
rect(s, 0, 0, 0.25, 7.5, ACCENT)
rect(s, 13.08, 0, 0.25, 7.5, ACCENT)
rect(s, 0, 0, 13.33, 0.18, ACCENT)
rect(s, 0, 7.32, 13.33, 0.18, ACCENT)

txt(s, "Thank You", 0, 1.8, 13.33, 1.5,
    font_size=72, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
txt(s, "Questions & Discussion", 0, 3.3, 13.33, 0.6,
    font_size=24, color=ACCENT2, align=PP_ALIGN.CENTER)

txt(s, "RAG Chatbot  •  Local • Private • Grounded",
    0, 4.2, 13.33, 0.45,
    font_size=16, color=ACCENT, align=PP_ALIGN.CENTER)

txt(s, "Built with  FastAPI · Streamlit · LangGraph · FAISS · Sentence-Transformers · Ollama LLaMA 3",
    0, 4.85, 13.33, 0.4,
    font_size=12, color=LIGHT_GRAY, align=PP_ALIGN.CENTER)

txt(s, "Dhruv Rakhecha  •  2026",
    0, 6.8, 13.33, 0.4,
    font_size=12, color=RGBColor(0x80, 0x99, 0xAA), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/RAG_chatbot/RAG_Chatbot_Presentation.pptx"
prs.save(output_path)
print(f"Presentation saved to: {output_path}")
print(f"Total slides: {len(prs.slides)}")
