# RAG Chatbot Project

This repository contains a Retrieval-Augmented Generation (RAG) chatbot built with Python. The system ingests documents, preprocesses them into vector embeddings, and uses a language model to answer queries with context from the stored knowledge base.

## 📁 Project Architecture & Flow

1. **Data Ingestion**
   - Documents are placed in `chunks/` or sourced from external files.
   - `src/chunker.py` helps split large files into smaller chunks if needed.

2. **Preprocessing**
   - `src/preprocessing.py` cleans text, removes noise, and prepares it for embedding.
   - Output is stored in `data/` or passed directly to the embedding step.

3. **Embedding Creation**
   - Use `src/embeddings.py` to convert textual chunks into vector representations.
   - Embeddings are stored in a vector store (`vectordb/index.faiss` by default) handled by `src/vector_store.py`.

4. **Retrieval**
   - `src/retriever.py` queries the vector store with a user prompt to fetch relevant chunks.

5. **RAG Pipeline**
   - `src/llm.py` and `src/prompt_builder.py` construct the final prompt combining retrieved context and user question.
   - `src/rag_graph.py` orchestrates the flow from retrieval to generation.

6. **Chatbot Interface**
   - `app.py` or `main.py` runs a simple API/CLI that accepts user queries and returns streamed responses.
   - Streaming is supported via response generators for real-time output.

## ⚙️ Steps to Run the Workflow

1. **Install requirements**
   ```bash
   pip install -r requirements.txt
   ```

2. **Preprocess your data**
   ```bash
   python src/preprocessing.py --input <path-to-documents> --output data/cleaned.json
   ```

3. **Create embeddings**
   ```bash
   python src/embeddings.py --input data/cleaned.json --output vectordb/index.faiss
   ```

4. **Build the RAG pipeline**
   ```bash
   python src/rag_graph.py --vector-store vectordb/index.faiss
   ```

## 🤖 Model and Embedding Choices

- **Language Model**: Default is [OpenAI GPT-4](https://openai.com) or an equivalent local model through `llm.py`. The choice should balance latency and cost. You can configure another model by adjusting `config.py`.
- **Embeddings**: Uses OpenAI embeddings (`text-embedding-3-small`) by default, but `embeddings.py` is modular for switching to other providers (e.g. `sentence-transformers`, `huggingface` models).

> _Tip_: For large datasets, consider higher-dimensional embeddings and a FAISS index with HNSW or IVF for faster retrieval.

## 🚀 Running the Chatbot with Streaming Response

1. Start the application:
   ```bash
   python app.py
   # or
   python main.py
   ```

2. In a separate terminal, send queries via curl or a web interface:
   ```bash
   curl -N localhost:8000/chat -d '{"query": "Explain CPU scheduling algorithms"}'
   ```
   The `-N` flag keeps the connection open for streaming tokens.

3. The server will stream partial outputs as they are generated, enabling a responsive chat experience.

## Demo video link
> 📺 [Demo Video](https://drive.google.com/drive/folders/1MQLUPPd0oObxO9xRm8KigH5AsWXsMaPR?usp=sharing) 

---
