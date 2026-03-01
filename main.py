import io
import json
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from src.chunker import save_chunks, split_into_chunks
from src.config import API_HOST, API_PORT
from src.embeddings import encode_texts
from src.llm import check_ollama_health
from src.preprocessing import preprocess_pdf, preprocess_text
from src.rag_graph import run_graph, run_graph_stream
from src.retriever import format_retrieved_for_response, retrieve
from src.vector_store import get_store, reset_store

app = FastAPI(title="RAG Chatbot API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class QueryRequest(BaseModel):
    question: str
    stream:   bool = False


@app.get("/health")
def health():
    store = get_store()
    ollama_ok, ollama_msg = check_ollama_health()
    return {
        "status":        "ok",
        "chunks_loaded": store.num_chunks,
        "ollama_ok":     ollama_ok,
        "ollama_msg":    ollama_msg,
    }


@app.post("/upload")
async def upload_document(file: UploadFile = File(...)):
    content = await file.read()
    filename = file.filename or "upload"

    if filename.lower().endswith(".pdf"):
        cleaned = preprocess_pdf(content)
    else:
        try:
            raw = content.decode("utf-8")
        except UnicodeDecodeError:
            raw = content.decode("latin-1")
        cleaned = preprocess_text(raw)

    chunks = split_into_chunks(cleaned)
    if not chunks:
        raise HTTPException(status_code=422, detail="No text could be extracted from the file.")

    doc_name = filename.rsplit(".", 1)[0]
    save_chunks(chunks, doc_name=doc_name)

    texts      = [c.text for c in chunks]
    embeddings = encode_texts(texts)

    store = reset_store()
    store.add_documents(chunks, embeddings)
    store.save()

    return {
        "message":    f"Indexed {len(chunks)} chunks from '{filename}'.",
        "num_chunks": len(chunks),
        "doc_name":   doc_name,
    }


@app.post("/query")
def query(req: QueryRequest):
    store = get_store()
    if store.is_empty:
        raise HTTPException(status_code=400, detail="No document indexed. Upload a file first.")

    if not req.question.strip():
        raise HTTPException(status_code=422, detail="Question must not be empty.")

    if req.stream:
        def event_stream():
            for event in run_graph_stream(req.question):
                yield f"data: {json.dumps(event)}\n\n"
            yield "data: [DONE]\n\n"
        return StreamingResponse(event_stream(), media_type="text/event-stream")

    result = run_graph(req.question)
    return {
        "answer":      result["response"],
        "sources":     result["sources"],
        "elapsed_ms":  result["elapsed_ms"],
    }


@app.get("/chunks")
def list_chunks():
    store = get_store()
    return {
        "num_chunks": store.num_chunks,
        "chunks": [
            {"chunk_id": c.chunk_id, "word_count": c.word_count, "preview": c.text[:120]}
            for c in store.get_all_chunks()
        ],
    }


@app.delete("/reset")
def reset():
    reset_store()
    return {"message": "Vector store cleared."}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host=API_HOST, port=API_PORT, reload=True)
