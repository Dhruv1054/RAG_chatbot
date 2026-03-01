import numpy as np
from sentence_transformers import SentenceTransformer
from src.config import EMBEDDING_MODEL, EMBEDDING_DIM

_model = None


def _get_model():
    global _model
    if _model is None:
        _model = SentenceTransformer(EMBEDDING_MODEL)
    return _model


def encode_texts(texts, batch_size=64):
    if not texts:
        return np.empty((0, EMBEDDING_DIM), dtype=np.float32)
    model = _get_model()
    embeddings = model.encode(
        texts,
        batch_size=batch_size,
        show_progress_bar=len(texts) > 100,
        normalize_embeddings=False,
        convert_to_numpy=True,
    )
    return embeddings.astype(np.float32)


def encode_query(query):
    model = _get_model()
    vec = model.encode([query], normalize_embeddings=False, convert_to_numpy=True)
    return vec.astype(np.float32)


def embedding_dim():
    return EMBEDDING_DIM
