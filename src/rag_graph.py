import time
from typing import List, TypedDict
from langgraph.graph import END, StateGraph
from src.llm import stream_response
from src.prompt_builder import build_prompt, truncate_context_to_fit
from src.retriever import format_retrieved_for_response, retrieve
from src.vector_store import SearchResult


class RAGState(TypedDict):
    question:         str
    retrieved_chunks: List[SearchResult]
    context:          str
    response:         str
    sources:          List[dict]
    elapsed_ms:       float


def retrieve_node(state):
    results = retrieve(state["question"])
    return {"retrieved_chunks": results, "sources": format_retrieved_for_response(results)}


def context_node(state):
    results = truncate_context_to_fit(state["retrieved_chunks"], state["question"], max_tokens=3000)
    from src.prompt_builder import build_context
    return {"context": build_context(results)}


def generate_node(state):
    prompt   = build_prompt(state["question"], state["retrieved_chunks"])
    response = "".join(stream_response(prompt))
    return {"response": response}


def _build_graph():
    graph = StateGraph(RAGState)
    graph.add_node("retrieve",      retrieve_node)
    graph.add_node("build_context", context_node)
    graph.add_node("generate",      generate_node)
    graph.set_entry_point("retrieve")
    graph.add_edge("retrieve",      "build_context")
    graph.add_edge("build_context", "generate")
    graph.add_edge("generate",      END)
    return graph



_compiled_graph = _build_graph().compile()


def run_graph(question):
    t0 = time.time()
    initial_state = {
        "question": question, "retrieved_chunks": [],
        "context": "", "response": "", "sources": [], "elapsed_ms": 0.0,
    }
    final_state = _compiled_graph.invoke(initial_state)
    final_state["elapsed_ms"] = round((time.time() - t0) * 1000, 1)
    return final_state


def run_graph_stream(question):
    t0      = time.time()
    results = retrieve(question)
    sources = format_retrieved_for_response(results)
    results = truncate_context_to_fit(results, question, max_tokens=3000)
    prompt  = build_prompt(question, results)
    for token in stream_response(prompt):
        yield {"type": "token", "data": token}
    yield {"type": "sources", "data": sources}
    yield {"type": "done",    "data": {"elapsed_ms": round((time.time() - t0) * 1000, 1)}}
