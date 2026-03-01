import pickle
from dataclasses import dataclass
from typing import List
import faiss
import numpy as np
from src.config import EMBEDDING_DIM, FAISS_INDEX_FILE, FAISS_META_FILE, TOP_K
from src.chunker import Chunk


@dataclass
class SearchResult:
    chunk: Chunk
    score: float
    rank:  int


class VectorStore:
    def __init__(self):
        self._index  = faiss.IndexFlatL2(EMBEDDING_DIM)
        self._chunks = []

    def add_documents(self, chunks, embeddings):
        if len(chunks) != embeddings.shape[0]:
            raise ValueError("Mismatch between chunks and embeddings count.")
        self._index  = faiss.IndexFlatL2(EMBEDDING_DIM)
        self._chunks = []
        self._index.add(embeddings.astype(np.float32))
        self._chunks = list(chunks)

    def search(self, query_embedding, k=TOP_K):
        if self._index.ntotal == 0:
            return []
        k = min(k, self._index.ntotal)
        distances, indices = self._index.search(query_embedding.astype(np.float32), k)
        results = []
        for rank, (dist, idx) in enumerate(zip(distances[0], indices[0])):
            if idx == -1:
                continue
            results.append(SearchResult(chunk=self._chunks[idx], score=float(dist), rank=rank))
        return results

    def save(self):
        faiss.write_index(self._index, str(FAISS_INDEX_FILE))
        with open(FAISS_META_FILE, "wb") as f:
            pickle.dump(self._chunks, f)

    def load(self):
        if not FAISS_INDEX_FILE.exists() or not FAISS_META_FILE.exists():
            return False
        self._index  = faiss.read_index(str(FAISS_INDEX_FILE))
        with open(FAISS_META_FILE, "rb") as f:
            self._chunks = pickle.load(f)
        return True

    @property
    def num_chunks(self):
        return self._index.ntotal

    @property
    def is_empty(self):
        return self._index.ntotal == 0

    def get_all_chunks(self):
        return list(self._chunks)


_store = None


def get_store():
    global _store
    if _store is None:
        _store = VectorStore()
        _store.load()
    return _store


def reset_store():
    global _store
    _store = VectorStore()
    return _store
