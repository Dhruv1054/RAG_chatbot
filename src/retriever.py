from src.config import TOP_K
from src.embeddings import encode_query
from src.vector_store import get_store


def retrieve(question, k=TOP_K):
    store = get_store()
    if store.is_empty:
        return []
    query_vec = encode_query(question)
    return store.search(query_vec, k=k)


def format_retrieved_for_response(results):
    return [
        {
            "chunk_id":   r.chunk.chunk_id,
            "text":       r.chunk.text,
            "score":      round(r.score, 4),
            "rank":       r.rank,
            "word_count": r.chunk.word_count,
        }
        for r in results
    ]
