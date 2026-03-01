from src.vector_store import SearchResult

_SYSTEM_TEMPLATE = """You are an AI assistant answering questions based strictly on the provided context.

Rules:
- Only use information from the context.
- If the answer is not in the context, say: "I don't have enough information from the provided document."
- Do not fabricate details.
- Keep answers concise but complete.

Context:
{context}

Question:
{question}

Answer:"""


def build_context(results):
    if not results:
        return "(No relevant context was found in the document.)"
    parts = []
    for r in results:
        header = f"[Excerpt {r.rank + 1} | similarity score: {r.score:.4f}]"
        parts.append(f"{header}\n{r.chunk.text.strip()}")
    return "\n\n---\n\n".join(parts)


def build_prompt(question, results):
    context = build_context(results)
    return _SYSTEM_TEMPLATE.format(context=context, question=question.strip())


def estimate_tokens(prompt):
    return len(prompt) // 4


def truncate_context_to_fit(results, question, max_tokens=3000):
    selected = []
    for r in results:
        candidate = selected + [r]
        prompt = build_prompt(question, candidate)
        if estimate_tokens(prompt) > max_tokens:
            break
        selected = candidate
    return selected if selected else results[:1]
