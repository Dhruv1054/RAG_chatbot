## This file is like control panel.
import os
from pathlib import Path

BASE_DIR   = Path(__file__).resolve().parent.parent # Root folder
DATA_DIR   = BASE_DIR / "data"  # where the cleaned text is saved
CHUNKS_DIR = BASE_DIR / "chunks" # where the chunks are stored
VECTOR_DIR = BASE_DIR / "vectordb" # where FAISS database is stored

for _d in (DATA_DIR, CHUNKS_DIR, VECTOR_DIR):
    _d.mkdir(parents=True, exist_ok=True)

CHUNK_SIZE    = 512 # each chunk = max 512 characters
CHUNK_OVERLAP = 128 # characters overlap by 128 chars so context is not lost at edges.

EMBEDDING_MODEL = "all-MiniLM-L6-v2" # AI model used which converts text into numbers
EMBEDDING_DIM   = 384

TOP_K = 4 # retrieve top 4 relevant chunks
 
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434") # where ollama runs locally
LLM_MODEL       = "llama3:8b-instruct-q4_K_M" #AI model used
LLM_TEMPERATURE = 0.1
LLM_TIMEOUT     = 120

API_HOST = os.getenv("API_HOST", "0.0.0.0")
API_PORT = int(os.getenv("API_PORT", 8000))

FAISS_INDEX_FILE = VECTOR_DIR / "index.faiss"  # stores actual vector embeddings
FAISS_META_FILE  = VECTOR_DIR / "metadata.pkl" # stores extra information about each chunk(original chunk, page numbers, )
