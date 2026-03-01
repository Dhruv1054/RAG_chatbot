import json
import requests
import streamlit as st

API_BASE = "http://localhost:8000"

st.set_page_config(page_title="RAG Chatbot", page_icon="📄", layout="wide")
st.title("📄 RAG Chatbot")
st.caption("Upload a PDF or text file, then ask questions about it.")

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Document Upload")
    uploaded = st.file_uploader("Choose a PDF or .txt file", type=["pdf", "txt"])

    if uploaded and st.button("Index Document", type="primary"):
        with st.spinner("Preprocessing & indexing…"):
            resp = requests.post(
                f"{API_BASE}/upload",
                files={"file": (uploaded.name, uploaded.getvalue(), uploaded.type)},
                timeout=120,
            )
        if resp.ok:
            data = resp.json()
            st.success(f"Indexed **{data['num_chunks']}** chunks from `{uploaded.name}`.")
            st.session_state["doc_ready"] = True
        else:
            st.error(f"Upload failed: {resp.text}")

    st.divider()

    try:
        health = requests.get(f"{API_BASE}/health", timeout=5).json()
        st.metric("Chunks in store", health.get("chunks_loaded", 0))
        ollama_ok = health.get("ollama_ok", False)
        st.markdown(
            f"Ollama: {'✅ connected' if ollama_ok else '❌ ' + health.get('ollama_msg', 'not running')}"
        )
    except Exception:
        st.warning("API not reachable. Is `main.py` running?")

    if st.button("Clear vector store"):
        requests.delete(f"{API_BASE}/reset", timeout=10)
        st.session_state.pop("doc_ready", None)
        st.rerun()

# ── Chat ─────────────────────────────────────────────────────────────────────
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        if msg.get("sources"):
            with st.expander("Sources"):
                for src in msg["sources"]:
                    st.markdown(
                        f"**Chunk {src['chunk_id']}** | score `{src['score']}` | "
                        f"{src['word_count']} words\n\n> {src['text'][:300]}…"
                    )

if prompt := st.chat_input("Ask a question about the document…"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        placeholder = st.empty()
        sources_box = st.container()
        full_text   = ""
        sources     = []

        try:
            with requests.post(
                f"{API_BASE}/query",
                json={"question": prompt, "stream": True},
                stream=True,
                timeout=180,
            ) as resp:
                if not resp.ok:
                    st.error(f"Error: {resp.text}")
                else:
                    for line in resp.iter_lines():
                        if not line:
                            continue
                        line = line.decode("utf-8") if isinstance(line, bytes) else line
                        if not line.startswith("data:"):
                            continue
                        payload = line[5:].strip()
                        if payload == "[DONE]":
                            break
                        try:
                            event = json.loads(payload)
                        except json.JSONDecodeError:
                            continue
                        if event["type"] == "token":
                            full_text += event["data"]
                            placeholder.markdown(full_text + "▌")
                        elif event["type"] == "sources":
                            sources = event["data"]
                        elif event["type"] == "done":
                            elapsed = event["data"].get("elapsed_ms", 0)
                            placeholder.markdown(full_text)
                            st.caption(f"Answered in {elapsed:.0f} ms")
        except requests.exceptions.ConnectionError:
            st.error("Cannot reach the API. Make sure `main.py` is running.")

        if sources:
            with sources_box.expander("Sources"):
                for src in sources:
                    st.markdown(
                        f"**Chunk {src['chunk_id']}** | score `{src['score']}` | "
                        f"{src['word_count']} words\n\n> {src['text'][:300]}…"
                    )

    st.session_state.messages.append({
        "role": "assistant", "content": full_text, "sources": sources
    })
