from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin   = Cm(3.0)
    section.right_margin  = Cm(2.5)

# ── Helpers ───────────────────────────────────────────────────
def heading(text, level=1, color=None):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.size = Pt(16 if level == 1 else 13)
        run.bold = True
        if color:
            run.font.color.rgb = RGBColor(*color)
    return p

def body(text, bold=False, italic=False, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    return p

def bullet(text, size=11):
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p

def numbered(text, size=11):
    p = doc.add_paragraph(style='List Number')
    run = p.add_run(text)
    run.font.size = Pt(size)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p

def shade_cell(cell, hex_color="1F4E79"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_text(cell, text, bold=False, size=10, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

# ══════════════════════════════════════════════════════════════
#  TITLE PAGE
# ══════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()

tp = doc.add_paragraph()
tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("PROJECT REPORT")
r.font.size = Pt(22)
r.bold = True
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

tp2 = doc.add_paragraph()
tp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = tp2.add_run("RAG Chatbot – Retrieval-Augmented Generation\nfor Intelligent Document Q&A")
r2.font.size = Pt(16)
r2.bold = True

doc.add_paragraph()
doc.add_paragraph()

info = [
    ("Submitted By", "Dhruv1054"),
    ("Branch",       "claude/create-project-presentation-PnPdW"),
    ("Technology",   "Python · FastAPI · Streamlit · FAISS · LangChain · Ollama"),
    ("Date",         "March 2026"),
]
for label, val in info:
    pi = doc.add_paragraph()
    pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl = pi.add_run(f"{label}: ")
    rl.bold = True
    rl.font.size = Pt(12)
    rv = pi.add_run(val)
    rv.font.size = Pt(12)

doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  1. INTRODUCTION
# ══════════════════════════════════════════════════════════════
heading("1. INTRODUCTION")
body(
    "In recent years, the exponential growth of digital information has created an urgent need for "
    "intelligent systems capable of understanding and retrieving relevant knowledge from large document "
    "repositories. Traditional keyword-based search engines, while effective for simple queries, often "
    "fail to capture the semantic intent behind user questions and return results that are syntactically "
    "matching but contextually irrelevant."
)
body(
    "Retrieval-Augmented Generation (RAG) is an emerging paradigm in Natural Language Processing (NLP) "
    "that bridges the gap between information retrieval and generative language modelling. By combining "
    "dense vector retrieval with powerful Large Language Models (LLMs), RAG systems can answer complex "
    "questions with accurate, grounded, and context-aware responses — dramatically reducing hallucination "
    "and improving factual accuracy."
)
body(
    "This project presents a fully functional RAG Chatbot built using Python, FastAPI, Streamlit, "
    "FAISS, LangChain, LangGraph, and Ollama. Users can upload their own PDF or text documents, which "
    "are automatically chunked, embedded, and indexed. At query time, the most relevant chunks are "
    "retrieved via semantic similarity search and passed as context to a locally-hosted Llama 3 LLM, "
    "which then generates a precise, context-grounded answer."
)
body(
    "The system is designed with a fully modular architecture, offering both a Streamlit web interface "
    "and a RESTful API, making it accessible for both end-users and developers. All inference runs "
    "locally via Ollama, ensuring data privacy and eliminating dependency on external cloud APIs."
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  2. PROBLEM STATEMENT
# ══════════════════════════════════════════════════════════════
heading("2. PROBLEM STATEMENT")
body(
    "Modern organisations and individuals deal with massive volumes of documents — research papers, "
    "manuals, reports, legal contracts, and academic notes. Extracting precise information from these "
    "documents manually is time-consuming, error-prone, and impractical at scale."
)
body(
    "Existing solutions present several critical limitations:"
)
bullet("Traditional search engines rely on keyword matching and cannot understand semantic meaning.")
bullet("General-purpose LLMs (GPT-4, Llama, etc.) hallucinate information not grounded in source documents.")
bullet("Cloud-based AI systems raise data privacy and security concerns for sensitive documents.")
bullet("Off-the-shelf chatbots do not allow users to query against their own custom document sets.")
bullet("Static FAQ bots cannot dynamically adapt to new documents without full retraining.")
body(
    "The core problem addressed by this project is: How can we build an intelligent, scalable, and "
    "privacy-preserving question-answering system that retrieves accurate, verifiable answers "
    "exclusively from user-provided documents — without hallucination and without sending data to "
    "external servers?"
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  3. OBJECTIVES
# ══════════════════════════════════════════════════════════════
heading("3. OBJECTIVES")
body("The primary objectives of this project are:")
numbered("Design and implement a complete Retrieval-Augmented Generation (RAG) pipeline from document ingestion to answer generation.")
numbered("Develop a robust document preprocessing module that cleans, normalises, and chunks PDF and text documents for efficient indexing.")
numbered("Implement a semantic embedding engine using sentence-transformers (all-MiniLM-L6-v2) to generate 384-dimensional dense vector representations.")
numbered("Build a FAISS-based vector store supporting persistent storage and efficient cosine/L2 similarity search over large document corpora.")
numbered("Integrate a locally-hosted LLM (Llama 3 via Ollama) for context-grounded answer generation with streaming support.")
numbered("Create a modular, maintainable codebase following separation-of-concerns principles with dedicated modules for each pipeline stage.")
numbered("Develop a FastAPI-based RESTful backend with endpoints for document upload, querying, health checks, and index management.")
numbered("Build an interactive Streamlit frontend that allows non-technical users to upload documents and query them conversationally.")
numbered("Ensure data privacy by keeping all inference and storage entirely local — no external API calls.")
numbered("Evaluate system performance in terms of retrieval accuracy, response latency, and answer relevance.")
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  4. SCOPE OF THE PROJECT
# ══════════════════════════════════════════════════════════════
heading("4. SCOPE OF THE PROJECT")
body(
    "The scope of this project encompasses the complete design, development, and deployment of a "
    "RAG-based intelligent chatbot system. The following boundaries define what the project includes "
    "and excludes."
)
body("Included in Scope:", bold=True)
bullet("Support for PDF (.pdf) and plain-text (.txt) document formats.")
bullet("Document chunking with configurable chunk size (512 characters) and overlap (128 characters).")
bullet("Sentence-transformer-based embedding and FAISS-based vector indexing.")
bullet("Local LLM inference using Ollama with Llama 3 8B quantised model.")
bullet("Streaming response delivery via Server-Sent Events (SSE).")
bullet("FastAPI RESTful API with five core endpoints (health, upload, query, chunks, reset).")
bullet("Streamlit-based conversational web UI with document upload and source visualisation.")
bullet("Persistent storage of document chunks and FAISS index across server restarts.")
bullet("LangGraph-based orchestration of the multi-step RAG pipeline.")
body("Excluded from Scope:", bold=True)
bullet("Support for image, video, or audio content within documents.")
bullet("Multi-user authentication and access control.")
bullet("Cloud deployment and horizontal scaling infrastructure.")
bullet("Fine-tuning or retraining of embedding or LLM models.")
bullet("Multi-lingual document support (English only in current version).")
bullet("Real-time collaborative document editing or annotation.")
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  5. PROPOSED SYSTEM
# ══════════════════════════════════════════════════════════════
heading("5. PROPOSED SYSTEM")
body(
    "The proposed system is a locally-deployable, privacy-preserving RAG Chatbot that enables "
    "end-users to upload documents and obtain accurate, context-grounded answers. Unlike traditional "
    "chatbots that rely on pre-trained knowledge, the proposed system grounds every answer in the "
    "content of the uploaded documents, effectively eliminating hallucination."
)
body("Key architectural decisions of the proposed system:", bold=True)
bullet(
    "Modular pipeline: The system is decomposed into discrete, independently testable modules — "
    "preprocessing, chunking, embedding, vector storage, retrieval, prompt building, and LLM generation."
)
bullet(
    "Local-first design: All components run on the user's machine. Document content never leaves the "
    "local environment, making the system suitable for sensitive or confidential documents."
)
bullet(
    "Dual interface: Both a developer-facing REST API (FastAPI) and a user-friendly web UI (Streamlit) "
    "are provided, catering to diverse use cases."
)
bullet(
    "LangGraph orchestration: The RAG pipeline is modelled as a directed graph using LangGraph, "
    "enabling clear state transitions and easy extensibility."
)
bullet(
    "Streaming responses: The LLM generates answers token-by-token and streams them to the client "
    "in real time, providing a responsive user experience."
)
body(
    "The proposed system addresses all limitations of existing solutions: it eliminates hallucination "
    "through context-only answer generation, protects privacy by running locally, supports any custom "
    "document set, and adapts dynamically to newly uploaded documents without retraining."
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  6. LITERATURE SURVEY  (15+ entries, APA table format)
# ══════════════════════════════════════════════════════════════
heading("6. LITERATURE SURVEY")
body(
    "Follow APA Format. The table below presents a review of 15 research papers relevant to the "
    "domains of Retrieval-Augmented Generation, semantic search, vector databases, large language "
    "models, document processing, and conversational AI systems."
)
doc.add_paragraph()

papers = [
    {
        "sno": "1",
        "title": "Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., … Kiela, D. (2020). Retrieval-Augmented Generation for knowledge-intensive NLP tasks. Advances in Neural Information Processing Systems, 33, 9459–9474.",
        "merits": "Introduced the foundational RAG framework; demonstrates significant reduction in hallucination; achieves state-of-the-art on Open-Domain QA benchmarks; combines retrieval and generation elegantly in a single end-to-end trainable architecture.",
        "demerits": "Requires a large pre-built Wikipedia index; retrieval latency can be high at inference; marginal improvement on tasks not requiring external knowledge; computationally expensive to train end-to-end.",
    },
    {
        "sno": "2",
        "title": "Karpukhin, V., Oğuz, B., Min, S., Lewis, P., Wu, L., Edunov, S., … Yih, W. (2020). Dense Passage Retrieval for open-domain question answering. Proceedings of EMNLP 2020, 6769–6781.",
        "merits": "Dense retrieval significantly outperforms BM25 sparse retrieval; dual-encoder architecture is efficient at inference; shows strong transfer to new QA datasets; forms the retrieval backbone for many RAG systems.",
        "demerits": "Requires large labelled QA datasets for training the retriever; does not generalise as well to domain-specific corpora without fine-tuning; index must be fully re-built when documents change.",
    },
    {
        "sno": "3",
        "title": "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence embeddings using Siamese BERT-Networks. Proceedings of EMNLP-IJCNLP 2019, 3982–3992.",
        "merits": "Produces semantically rich sentence embeddings 9× faster than BERT cross-encoders; supports cosine similarity comparison at scale; widely adopted as an embedding backbone; pre-trained models available for diverse domains.",
        "demerits": "Embedding quality can degrade for very long documents; requires fine-tuning for specialised domains; English-centric pre-trained models may underperform on multilingual tasks.",
    },
    {
        "sno": "4",
        "title": "Johnson, J., Douze, M., & Jégou, H. (2019). Billion-scale similarity search with GPUs. IEEE Transactions on Big Data, 7(3), 535–547.",
        "merits": "FAISS enables billion-scale nearest-neighbour search; GPU-accelerated index for ultra-low latency; multiple index types (Flat, IVF, PQ) for accuracy–speed trade-offs; open-source and widely used in production RAG systems.",
        "demerits": "Flat index does not scale to billions of vectors without quantisation; IVF indexes require parameter tuning (nlist, nprobe); GPU version requires significant VRAM; not natively distributed.",
    },
    {
        "sno": "5",
        "title": "Touvron, H., Martin, L., Stone, K., Albert, P., Almahairi, A., Babaei, Y., … Scialom, T. (2023). Llama 2: Open foundation and fine-tuned chat models. arXiv preprint arXiv:2307.09288.",
        "merits": "Open-weight model enabling local, privacy-preserving deployment; competitive with GPT-3.5 on many benchmarks; RLHF-trained chat version follows instructions reliably; broad community adoption and tooling support.",
        "demerits": "8B model requires 16 GB RAM; smaller models sacrifice accuracy; context window limited to 4096 tokens in base version; not optimised for retrieval-augmented settings out of the box.",
    },
    {
        "sno": "6",
        "title": "Chase, H. (2022). LangChain: Building applications with LLMs through composability [Software]. https://github.com/langchain-ai/langchain",
        "merits": "Provides high-level abstractions for document loaders, text splitters, vector stores, and LLM chains; dramatically reduces boilerplate code; active community with extensive integrations; supports streaming natively.",
        "demerits": "Abstractions can obscure underlying behaviour, making debugging difficult; rapidly evolving API introduces breaking changes; overhead from abstraction layers can impact performance in latency-critical applications.",
    },
    {
        "sno": "7",
        "title": "Gao, Y., Xiong, Y., Gao, X., Jia, K., Pan, J., Bi, Y., … Wang, H. (2023). Retrieval-Augmented Generation for large language models: A survey. arXiv preprint arXiv:2312.10997.",
        "merits": "Comprehensive taxonomy of RAG variants (Naive, Advanced, Modular); identifies open research challenges; reviews 100+ papers; provides a roadmap for RAG system design.",
        "demerits": "Survey nature means no novel empirical results; rapidly outdated given the speed of LLM research; limited coverage of deployment and production considerations.",
    },
    {
        "sno": "8",
        "title": "Shi, W., Min, S., Yasunaga, M., Seo, M., James, R., Lewis, M., … Zettlemoyer, L. (2023). REPLUG: Retrieval-Augmented Language Model Pre-Training. arXiv preprint arXiv:2301.12652.",
        "merits": "Treats the LLM as a black box; retriever is trained to improve LLM perplexity; achieves strong performance on language modelling and QA; compatible with any frozen LLM.",
        "demerits": "Pre-training the retriever is computationally expensive; improvements on downstream tasks are incremental over simpler RAG baselines; requires access to large text corpora.",
    },
    {
        "sno": "9",
        "title": "Ram, O., Levine, Y., Dalmedigos, I., Muhlgay, D., Shashua, A., Leyton-Brown, K., & Shoham, Y. (2023). In-context retrieval-augmented language models. Transactions of the Association for Computational Linguistics, 11, 1316–1331.",
        "merits": "No retriever training required; leverages in-context learning of frozen LLMs; applicable to any autoregressive LM; strong performance on long-document tasks.",
        "demerits": "Performance bounded by context window size; concatenating many retrieved passages increases inference cost; sensitivity to passage ordering in the prompt.",
    },
    {
        "sno": "10",
        "title": "Fan, A., Grave, E., & Joulin, A. (2021). Reducing transformer depth on demand with structured dropout. Proceedings of ICLR 2021.",
        "merits": "LayerDrop reduces inference cost at deployment without retraining; enables adaptive depth at query time; improves deployment flexibility for resource-constrained environments.",
        "demerits": "Minor accuracy degradation with high dropout rates; requires modified training procedure; not widely adopted in RAG-specific architectures.",
    },
    {
        "sno": "11",
        "title": "Izacard, G., & Grave, E. (2021). Leveraging passage retrieval with generative models for open domain question answering. Proceedings of EACL 2021, 874–880.",
        "merits": "Fusion-in-Decoder (FiD) fuses multiple retrieved passages at the decoder level, outperforming single-passage approaches; scalable to many retrieved passages; sets NaturalQuestions and TriviaQA SOTA.",
        "demerits": "Decoder computation scales linearly with number of passages; requires a T5-based seq2seq model; not directly applicable to decoder-only LLMs like Llama.",
    },
    {
        "sno": "12",
        "title": "Shi, F., Chen, X., Misra, K., Scales, N., Dohan, D., Chi, E., … Zhou, D. (2023). Large language models can be easily distracted by irrelevant context. Proceedings of ICML 2023.",
        "merits": "Identifies a critical failure mode of RAG systems (irrelevant context distraction); provides benchmark datasets; motivates better retrieval filtering and re-ranking strategies.",
        "demerits": "Study is primarily analytical rather than prescriptive; mitigation strategies explored are limited; findings may not generalise across all LLM families.",
    },
    {
        "sno": "13",
        "title": "Jiang, Z., Xu, F. F., Gao, L., Sun, Z., Liu, Q., Dwivedi-Yu, J., … Neubig, G. (2023). Active retrieval augmented generation. Proceedings of EMNLP 2023, 7969–7992.",
        "merits": "FLARE performs retrieval only when needed (forward-looking active retrieval), reducing unnecessary API calls; improves efficiency; demonstrates better calibration of retrieval decisions.",
        "demerits": "Triggering heuristics are approximate; adds latency at generation time; harder to implement than passive RAG pipelines; requires a capable base LLM.",
    },
    {
        "sno": "14",
        "title": "Wang, L., Yang, N., Huang, X., Jiao, B., Yang, L., Jiang, D., … Wei, F. (2022). Text embeddings by weakly-supervised contrastive pre-training. arXiv preprint arXiv:2212.03533.",
        "merits": "E5 embeddings achieve top performance on BEIR and MTEB benchmarks; weakly supervised training on large web corpora reduces labelled data requirements; multi-lingual variant available.",
        "demerits": "Large model variants require significant compute; smaller variants underperform on domain-specific tasks; not specifically optimised for short-chunk retrieval as used in this project.",
    },
    {
        "sno": "15",
        "title": "Zhao, W. X., Zhou, K., Li, J., Tang, T., Wang, X., Hou, Y., … Wen, J.-R. (2023). A survey of large language models. arXiv preprint arXiv:2303.18223.",
        "merits": "Comprehensive review of LLM architectures, training strategies, alignment methods, and applications; 500+ references; identifies evaluation benchmarks and open research challenges.",
        "demerits": "Breadth over depth — individual topics are covered superficially; rapidly outdated given the pace of LLM research; limited discussion of RAG-specific LLM considerations.",
    },
    {
        "sno": "16",
        "title": "Zhu, Y., Yuan, H., Wang, S., Liu, J., Liu, W., Deng, C., … Wen, J.-R. (2023). Large language models for information retrieval: A survey. arXiv preprint arXiv:2308.07107.",
        "merits": "Surveys LLM applications across query rewriting, document ranking, and knowledge augmentation; identifies integration challenges; provides a unified taxonomy.",
        "demerits": "Primarily theoretical; lacks empirical comparisons between methods; coverage of production deployment challenges is limited.",
    },
    {
        "sno": "17",
        "title": "Besta, M., Blach, N., Kubicek, A., Gerstenberger, R., Gianinazzi, L., Gajda, J., … Hoefler, T. (2024). Graph of Thoughts: Solving elaborate problems with large language models. Proceedings of AAAI 2024.",
        "merits": "Graph-of-Thoughts extends chain-of-thought prompting to DAG structures; enables non-linear reasoning; improves performance on complex multi-step problems.",
        "demerits": "Significantly increases prompt complexity and token usage; requires careful graph design; overkill for simple document Q&A tasks addressed in this project.",
    },
]

# Build table
col_widths = [Cm(1.0), Cm(6.5), Cm(5.5), Cm(5.5)]
table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Header row
hdr_cells = table.rows[0].cells
headers = ["S.NO", "TITLE (APA Format)", "MERITS", "DEMERITS"]
for i, (cell, hdr) in enumerate(zip(hdr_cells, headers)):
    shade_cell(cell, "1F4E79")
    set_cell_text(cell, hdr, bold=True, size=10, color=(255, 255, 255), align=WD_ALIGN_PARAGRAPH.CENTER)
    cell.width = col_widths[i]

# Data rows
fill_colors = ["D9E1F2", "EBF3FB"]
for idx, p in enumerate(papers):
    row_cells = table.add_row().cells
    shade_cell(row_cells[0], fill_colors[idx % 2])
    shade_cell(row_cells[1], fill_colors[idx % 2])
    shade_cell(row_cells[2], "E2EFDA" if idx % 2 == 0 else "F0FFF0")
    shade_cell(row_cells[3], "FCE4D6" if idx % 2 == 0 else "FFF2CC")
    set_cell_text(row_cells[0], p["sno"], bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(row_cells[1], p["title"], size=9)
    set_cell_text(row_cells[2], p["merits"], size=9)
    set_cell_text(row_cells[3], p["demerits"], size=9)
    for c, w in zip(row_cells, col_widths):
        c.width = w

doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  7. FINDINGS IN LITERATURE SURVEY
# ══════════════════════════════════════════════════════════════
heading("7. FINDINGS IN LITERATURE SURVEY")
body(
    "A systematic review of the 17 papers in the literature survey yields the following key findings "
    "that directly informed the design and implementation of the RAG Chatbot:"
)

body("7.1 RAG Is the Established Paradigm for Grounded QA", bold=True)
body(
    "Lewis et al. (2020) established that RAG significantly reduces hallucination compared to "
    "purely parametric LLMs. This finding justified the core architectural decision to ground "
    "every system response in retrieved document chunks rather than relying on the LLM's "
    "internal knowledge."
)

body("7.2 Dense Retrieval Outperforms Sparse Methods", bold=True)
body(
    "Karpukhin et al. (2020) and Wang et al. (2022) confirm that dense vector retrieval using "
    "bi-encoder architectures substantially outperforms traditional BM25 keyword-matching for "
    "semantic question answering. This validated the choice of sentence-transformers and FAISS "
    "over simple inverted-index approaches."
)

body("7.3 Sentence-BERT Embeddings Provide Optimal Speed–Quality Trade-off", bold=True)
body(
    "Reimers & Gurevych (2019) demonstrated that Sentence-BERT embeddings (all-MiniLM-L6-v2) "
    "achieve near-SOTA semantic similarity performance while being 9× faster than cross-encoder "
    "approaches. The 384-dimensional embedding space provides compact yet expressive representations "
    "suitable for real-time retrieval."
)

body("7.4 FAISS Enables Scalable Billion-Scale Similarity Search", bold=True)
body(
    "Johnson et al. (2019) showed FAISS can perform billion-scale similarity search with "
    "sub-millisecond latency per query on GPU. Even on CPU, the IndexFlatL2 index used in this "
    "project supports fast exact search over thousands of document chunks."
)

body("7.5 Local LLMs Offer Viable Privacy-Preserving Inference", bold=True)
body(
    "Touvron et al. (2023) demonstrated that open-weight models like Llama 2/3 can match "
    "proprietary API models on many benchmarks. This enabled the decision to use Ollama for "
    "fully local LLM inference, eliminating data privacy concerns."
)

body("7.6 Irrelevant Context Remains a Critical Challenge", bold=True)
body(
    "Shi et al. (2023) identified that LLMs are easily distracted by irrelevant retrieved passages. "
    "This finding motivated the system's context truncation strategy (max_tokens=3000) and the "
    "system prompt that explicitly instructs the LLM to answer only from provided context."
)

body("7.7 Chunking Strategy Significantly Impacts Retrieval Quality", bold=True)
body(
    "Across multiple papers, proper document segmentation was identified as a critical factor. "
    "Based on these findings, the system uses a 512-character chunk size with 128-character overlap "
    "via RecursiveCharacterTextSplitter, balancing context preservation with retrieval granularity."
)

body("7.8 LangChain Reduces Development Complexity", bold=True)
body(
    "Chase (2022) established LangChain as the de-facto orchestration framework for LLM applications. "
    "LangGraph (an extension) provides graph-based pipeline management, which was used to implement "
    "the clean, stateful RAG pipeline with well-defined state transitions."
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  8. METHODOLOGY
# ══════════════════════════════════════════════════════════════
heading("8. METHODOLOGY")
body(
    "The system was developed following an iterative, module-by-module approach. The methodology "
    "comprises the following phases:"
)

body("Phase 1 – Requirements Analysis", bold=True)
bullet("Identify functional and non-functional requirements through literature review and stakeholder analysis.")
bullet("Define system boundaries, data flow, and interface contracts between modules.")
bullet("Select technology stack based on literature findings (FAISS, sentence-transformers, Ollama, LangChain).")

body("Phase 2 – System Design", bold=True)
bullet("Design modular architecture with eight discrete pipeline stages.")
bullet("Define API contracts using Pydantic models for all request/response schemas.")
bullet("Design FAISS index structure and chunk metadata schema.")
bullet("Create use case diagrams, class diagrams, and sequence diagrams.")

body("Phase 3 – Document Ingestion Pipeline", bold=True)
bullet("Step 1 – Preprocessing: Accept PDF (via PyMuPDF) or TXT files; extract raw text; remove page numbers, headers, footers using regex; normalise Unicode (NFKD); collapse whitespace.")
bullet("Step 2 – Chunking: Apply RecursiveCharacterTextSplitter with chunk_size=512 and chunk_overlap=128 to produce overlapping text segments.")
bullet("Step 3 – Embedding: Encode each chunk using all-MiniLM-L6-v2 to produce 384-dimensional float32 vectors.")
bullet("Step 4 – Indexing: Add embeddings to FAISS IndexFlatL2; persist index and chunk metadata pickle to disk.")

body("Phase 4 – Query Processing Pipeline", bold=True)
bullet("Step 5 – Query Embedding: Encode the user's natural-language question using the same embedding model.")
bullet("Step 6 – Retrieval: Search FAISS index for top-4 most similar chunks using L2 distance.")
bullet("Step 7 – Context Assembly: Concatenate retrieved chunks; truncate to fit LLM context window (max_tokens=3000).")
bullet("Step 8 – Prompt Construction: Build structured prompt with system instruction, retrieved context, and user question.")
bullet("Step 9 – LLM Generation: Send prompt to Ollama (Llama 3 8B); stream response token-by-token via SSE.")

body("Phase 5 – API & UI Development", bold=True)
bullet("Implement FastAPI endpoints: /health, /upload, /query (streaming), /chunks, /reset.")
bullet("Implement Streamlit UI with file uploader, chat history, and source snippet display.")

body("Phase 6 – Testing & Evaluation", bold=True)
bullet("Unit test each module (preprocessor, chunker, embedder, vector store, retriever, prompt builder).")
bullet("Integration test the full RAG pipeline end-to-end.")
bullet("Evaluate retrieval quality with precision@4 on held-out document QA pairs.")
bullet("Measure end-to-end response latency.")
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  9. SOFTWARE REQUIREMENTS
# ══════════════════════════════════════════════════════════════
heading("9. SOFTWARE REQUIREMENTS")

heading("9.1 Functional Requirements", level=2)

body("FR-1: Document Upload", bold=True)
bullet("The system shall accept PDF and TXT file uploads via the /upload REST endpoint and the Streamlit UI.")
bullet("The system shall extract text from PDFs using PyMuPDF and decode TXT files with UTF-8/Latin-1 fallback.")
bullet("The system shall preprocess, chunk, embed, and index uploaded documents automatically.")

body("FR-2: Query Processing", bold=True)
bullet("The system shall accept natural-language queries via the /query REST endpoint and the Streamlit chat interface.")
bullet("The system shall retrieve the top-4 most semantically similar document chunks for each query.")
bullet("The system shall construct a context-grounded prompt and generate an answer using the local LLM.")
bullet("The system shall support real-time token-streaming responses via Server-Sent Events.")

body("FR-3: Vector Store Management", bold=True)
bullet("The system shall persist FAISS index and chunk metadata across server restarts.")
bullet("The system shall expose a /reset endpoint to clear all stored data.")
bullet("The system shall expose a /chunks endpoint to list all stored chunks with previews.")

body("FR-4: Health Monitoring", bold=True)
bullet("The system shall expose a /health endpoint reporting chunk count, Ollama connectivity, and embedding model status.")

heading("9.2 Non-Functional Requirements", level=2)

body("NFR-1: Performance", bold=True)
bullet("Query-to-first-token latency shall be ≤ 5 seconds on standard hardware (8 GB RAM, 4-core CPU).")
bullet("Document indexing shall complete within 30 seconds for a 50-page PDF.")
bullet("FAISS retrieval shall complete within 100 ms for a corpus of 10,000 chunks.")

body("NFR-2: Reliability", bold=True)
bullet("The system shall handle malformed or encrypted PDFs gracefully, returning a descriptive error message.")
bullet("The system shall remain available if Ollama is temporarily unreachable, returning a 503 status with diagnostic information.")

body("NFR-3: Security", bold=True)
bullet("All document processing and LLM inference shall be performed locally; no document content shall be transmitted to external services.")
bullet("The API shall validate all file uploads for type and size before processing.")

body("NFR-4: Maintainability", bold=True)
bullet("The codebase shall follow a modular architecture with a single responsibility per module.")
bullet("All module interfaces shall be defined using Pydantic data models.")
bullet("Configuration parameters shall be centralised in src/config.py.")

body("NFR-5: Usability", bold=True)
bullet("The Streamlit UI shall require no technical knowledge to operate — file upload and chat interaction shall be self-explanatory.")
bullet("API documentation shall be auto-generated and accessible at /docs (Swagger UI).")

body("NFR-6: Portability", bold=True)
bullet("The system shall run on Linux, macOS, and Windows (via WSL2 for Ollama).")
bullet("All dependencies shall be declared in requirements.txt with pinned versions.")
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  10. SYSTEM ARCHITECTURE
# ══════════════════════════════════════════════════════════════
heading("10. SYSTEM ARCHITECTURE")
body(
    "The system follows a layered, modular architecture. Figure 1 shows the high-level data flow "
    "from document upload to answer generation. The architecture is composed of five layers: "
    "Interface Layer, API Layer, Processing Layer, Storage Layer, and Inference Layer."
)

body("10.1 High-Level Architecture", bold=True)
body(
    "┌───────────────────────────────────────────────────────────────────┐\n"
    "│                     INTERFACE LAYER                               │\n"
    "│      Streamlit Web UI  │  REST API Clients  │  curl / SDK         │\n"
    "├───────────────────────────────────────────────────────────────────┤\n"
    "│                       API LAYER  (FastAPI)                        │\n"
    "│  /upload │ /query (SSE) │ /chunks │ /reset │ /health              │\n"
    "├───────────────────────────────────────────────────────────────────┤\n"
    "│                     PROCESSING LAYER                              │\n"
    "│  Preprocessor → Chunker → Embedder → VectorStore → Retriever     │\n"
    "│               PromptBuilder → LangGraph RAG Pipeline              │\n"
    "├───────────────────────────────────────────────────────────────────┤\n"
    "│                      STORAGE LAYER                                │\n"
    "│   FAISS Index (vectordb/)  │  Chunk Metadata (chunks/)           │\n"
    "├───────────────────────────────────────────────────────────────────┤\n"
    "│                     INFERENCE LAYER                               │\n"
    "│          Ollama Server  →  Llama 3 8B-Instruct (Q4_K_M)         │\n"
    "└───────────────────────────────────────────────────────────────────┘"
)

body("10.2 Use Case Diagram", bold=True)
body(
    "The following textual representation describes the use cases. (UML diagrams are embedded below "
    "in ASCII notation for inclusion in document; a visual diagram tool such as draw.io may be used "
    "to render the formal UML version.)"
)
body(
    "┌─────────────────────────────────────────────────────────────────┐\n"
    "│                        RAG Chatbot System                       │\n"
    "│                                                                  │\n"
    "│  ┌──────────┐   upload document ──────────► [Upload Document]  │\n"
    "│  │          │   query document  ──────────► [Submit Query]      │\n"
    "│  │   User   │   view sources    ──────────► [View Source Chunks]│\n"
    "│  │          │   reset index     ──────────► [Reset Vector Store]│\n"
    "│  └──────────┘   check health   ──────────► [Monitor Health]    │\n"
    "│                                                                  │\n"
    "│  ┌────────────┐                                                  │\n"
    "│  │ Admin/Dev  │   manage config ──────────► [Configure System]  │\n"
    "│  └────────────┘                                                  │\n"
    "└─────────────────────────────────────────────────────────────────┘"
)

body("Use Case Descriptions:", bold=True)
uc_data = [
    ("UC-01", "Upload Document", "User", "User uploads PDF or TXT file; system preprocesses, chunks, embeds, and indexes it."),
    ("UC-02", "Submit Query",    "User", "User types a question; system retrieves relevant chunks and streams LLM-generated answer."),
    ("UC-03", "View Source Chunks", "User", "User browses indexed document chunks with previews via /chunks endpoint."),
    ("UC-04", "Reset Vector Store", "User/Admin", "Actor clears all indexed documents and resets FAISS index."),
    ("UC-05", "Monitor Health",  "Admin/Dev", "Actor queries /health for system status, chunk count, and Ollama connectivity."),
    ("UC-06", "Configure System","Admin/Dev", "Admin modifies src/config.py to change chunk size, model, or retrieval parameters."),
]
uc_table = doc.add_table(rows=1, cols=4)
uc_table.style = 'Table Grid'
for cell, hdr in zip(uc_table.rows[0].cells, ["ID", "Use Case", "Actor", "Description"]):
    shade_cell(cell, "2E75B6")
    set_cell_text(cell, hdr, bold=True, size=10, color=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
for uc in uc_data:
    row = uc_table.add_row().cells
    for i, val in enumerate(uc):
        set_cell_text(row[i], val, size=10)

doc.add_paragraph()
body("10.3 Class Diagram", bold=True)
body(
    "The class diagram below shows the primary classes and their relationships in the system:"
)
body(
    "┌──────────────────┐       ┌──────────────────┐       ┌──────────────────┐\n"
    "│  DocumentLoader  │──────►│  Preprocessor    │──────►│    Chunker       │\n"
    "│ +load(path)      │       │ +clean(text)     │       │ +split(text)     │\n"
    "│ +extract_pdf()   │       │ +normalize()     │       │ chunk_size: int  │\n"
    "│ +extract_txt()   │       │ +remove_headers()│       │ overlap: int     │\n"
    "└──────────────────┘       └──────────────────┘       └────────┬─────────┘\n"
    "                                                                │\n"
    "                                                      ┌─────────▼─────────┐\n"
    "┌──────────────────┐       ┌──────────────────┐       │   EmbeddingModel  │\n"
    "│  VectorStore     │◄──────│   Embeddings     │◄──────│ +encode(chunks)   │\n"
    "│ +add(vectors)    │       │ +embed(text)     │       │ model: str        │\n"
    "│ +search(q, k)    │       │ +load_model()    │       │ dim: int = 384    │\n"
    "│ +persist()       │       └──────────────────┘       └───────────────────┘\n"
    "│ +load()          │\n"
    "└────────┬─────────┘\n"
    "         │\n"
    "┌────────▼─────────┐       ┌──────────────────┐       ┌──────────────────┐\n"
    "│   Retriever      │──────►│  PromptBuilder   │──────►│   LLMInterface   │\n"
    "│ +retrieve(q, k)  │       │ +build(q, chunks)│       │ +generate(prompt)│\n"
    "│ top_k: int = 4   │       │ +truncate()      │       │ +stream(prompt)  │\n"
    "└──────────────────┘       └──────────────────┘       │ model: str       │\n"
    "                                                       └──────────────────┘\n"
    "         ┌──────────────────────────────────────────────────────────┐\n"
    "         │                    RAGGraph (LangGraph)                   │\n"
    "         │ nodes: [retrieve, build_prompt, generate]                │\n"
    "         │ state: RAGState {query, chunks, prompt, answer}          │\n"
    "         │ +run(query) → answer                                     │\n"
    "         └──────────────────────────────────────────────────────────┘"
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  11. IMPLEMENTATION
# ══════════════════════════════════════════════════════════════
heading("11. IMPLEMENTATION")
body(
    "This section details the implementation of each module in the RAG pipeline, including "
    "key design decisions, code architecture, and sample logic."
)

body("11.1 Project Structure", bold=True)
body(
    "RAG_chatbot/\n"
    "├── main.py            # FastAPI server (5 endpoints)\n"
    "├── app.py             # Streamlit web UI\n"
    "├── requirements.txt   # Pinned dependencies\n"
    "├── src/\n"
    "│   ├── config.py      # Central configuration constants\n"
    "│   ├── preprocessing.py  # Document cleaning & normalisation\n"
    "│   ├── chunker.py        # RecursiveCharacterTextSplitter wrapper\n"
    "│   ├── embeddings.py     # Sentence-transformer encoder\n"
    "│   ├── vector_store.py   # FAISS index management\n"
    "│   ├── retriever.py      # Top-K semantic retrieval\n"
    "│   ├── llm.py            # Ollama streaming interface\n"
    "│   ├── prompt_builder.py # Context-aware prompt assembly\n"
    "│   └── rag_graph.py      # LangGraph pipeline orchestrator\n"
    "├── chunks/            # Persisted document chunks\n"
    "└── vectordb/          # Persisted FAISS index"
)

body("11.2 Configuration (src/config.py)", bold=True)
body(
    "All system parameters are centralised in config.py to enable easy tuning without modifying "
    "module logic. Key constants include CHUNK_SIZE=512, CHUNK_OVERLAP=128, EMBEDDING_MODEL="
    "'all-MiniLM-L6-v2', TOP_K=4, MAX_CONTEXT_TOKENS=3000, OLLAMA_MODEL='llama3:8b-instruct-q4_K_M', "
    "and TEMPERATURE=0.1."
)

body("11.3 Preprocessing (src/preprocessing.py)", bold=True)
body(
    "The preprocessor applies a sequence of cleaning operations to raw document text: "
    "(1) Unicode NFKD normalisation, (2) regex-based removal of page number patterns "
    "(e.g., 'Page 3 of 10'), (3) detection and removal of repeated header/footer lines "
    "using frequency analysis, and (4) whitespace collapsing. The result is clean, "
    "prose-only text ready for chunking."
)

body("11.4 Chunking (src/chunker.py)", bold=True)
body(
    "LangChain's RecursiveCharacterTextSplitter is used with separators=['\n\n', '\n', '. ', ' ']. "
    "This hierarchy ensures chunks prefer natural paragraph and sentence boundaries over arbitrary "
    "character splits. The 128-character overlap preserves context across chunk boundaries, "
    "preventing loss of information for queries that span two adjacent chunks."
)

body("11.5 Embedding (src/embeddings.py)", bold=True)
body(
    "The SentenceTransformer('all-MiniLM-L6-v2') model produces 384-dimensional L2-normalised "
    "vectors. The model is loaded once at startup and reused for all encode() calls, avoiding "
    "repeated model loading overhead. Batch encoding is used for efficient indexing of large "
    "document sets."
)

body("11.6 Vector Store (src/vector_store.py)", bold=True)
body(
    "FAISS IndexFlatL2 performs exact nearest-neighbour search using L2 distance. The index is "
    "serialised to disk using faiss.write_index() after each upload, and loaded at startup if "
    "it exists. Chunk metadata (text, document name, position) is stored in a parallel Python "
    "pickle file, allowing the retriever to return text chunks alongside their similarity scores."
)

body("11.7 Retriever (src/retriever.py)", bold=True)
body(
    "The retriever embeds the user query, calls faiss_index.search(query_vector, k=4), and returns "
    "the top-4 chunks with their L2 distances. Distance scores are inverted and normalised to "
    "produce a relevance score in [0, 1] for display in the Streamlit UI."
)

body("11.8 LLM Interface (src/llm.py)", bold=True)
body(
    "The LLM module posts to the Ollama /api/generate endpoint with stream=True. Responses are "
    "yielded token-by-token using Python's requests library with stream=True, enabling real-time "
    "Server-Sent Events delivery through FastAPI's StreamingResponse."
)

body("11.9 Prompt Builder (src/prompt_builder.py)", bold=True)
body(
    "The prompt builder assembles: (1) a system instruction ('Answer only from the provided context. "
    "If the answer is not in the context, say so.'), (2) numbered context passages from retrieved "
    "chunks, and (3) the user question. Token estimation (4 characters per token heuristic) ensures "
    "the total prompt does not exceed MAX_CONTEXT_TOKENS=3000."
)

body("11.10 RAG Graph (src/rag_graph.py)", bold=True)
body(
    "LangGraph models the pipeline as a directed graph with nodes: retrieve → build_prompt → generate. "
    "Each node receives a typed RAGState object and returns an updated state. This design enables "
    "easy insertion of new pipeline stages (e.g., query rewriting, re-ranking) without modifying "
    "existing nodes."
)

body("11.11 FastAPI Server (main.py)", bold=True)
body(
    "The FastAPI server exposes five endpoints. The /upload endpoint accepts a multipart file, "
    "runs the full ingestion pipeline, and returns the number of chunks indexed. The /query endpoint "
    "accepts a JSON body and returns a StreamingResponse with SSE-formatted tokens. The /health "
    "endpoint performs a lightweight Ollama ping and returns system diagnostics."
)

body("11.12 Streamlit UI (app.py)", bold=True)
body(
    "The Streamlit frontend provides: a file uploader for PDF/TXT upload, a chat input field, "
    "a scrollable message history, and an expandable section showing retrieved source chunks "
    "for each answer. The UI calls the FastAPI backend directly, consuming the SSE stream and "
    "rendering tokens progressively using st.write_stream()."
)
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  12. SUMMARY
# ══════════════════════════════════════════════════════════════
heading("12. SUMMARY")
body(
    "This report presented the complete design, development, and implementation of a "
    "Retrieval-Augmented Generation (RAG) Chatbot system. The project successfully addresses "
    "the critical limitations of traditional document search and general-purpose LLMs by "
    "combining semantic vector retrieval with locally-hosted generative AI."
)
body("Key Achievements:", bold=True)
bullet("A fully functional, privacy-preserving RAG pipeline was implemented using sentence-transformers, FAISS, LangChain, LangGraph, and Ollama.")
bullet("The system supports PDF and TXT document ingestion with intelligent chunking, embedding, and persistent vector indexing.")
bullet("Real-time streaming responses are delivered via Server-Sent Events, providing a responsive user experience.")
bullet("A dual interface — Streamlit web UI and FastAPI REST API — makes the system accessible to both end-users and developers.")
bullet("All inference runs locally via Ollama (Llama 3 8B), ensuring complete data privacy.")
bullet("The modular architecture (8 distinct modules) enables independent testing, easy maintenance, and straightforward extensibility.")
body("Conclusions:", bold=True)
body(
    "The literature survey confirmed that RAG is the state-of-the-art approach for grounded question "
    "answering, and that dense vector retrieval with sentence-transformer embeddings provides the "
    "optimal speed–quality trade-off for real-time systems. The implemented system validates these "
    "findings in practice, demonstrating accurate, context-grounded answers with low latency."
)
body("Future Work:", bold=True)
bullet("Implement hybrid retrieval combining FAISS dense search with BM25 sparse retrieval (reciprocal rank fusion).")
bullet("Add query rewriting and hypothetical document embedding (HyDE) to improve retrieval recall.")
bullet("Support multi-modal documents (tables, figures) via OCR and visual embeddings.")
bullet("Implement multi-user authentication and per-user document namespacing.")
bullet("Evaluate and fine-tune the embedding model on domain-specific corpora.")
bullet("Develop a Docker-based deployment package for one-click installation.")
doc.add_page_break()

# ══════════════════════════════════════════════════════════════
#  13. REFERENCES  (17 APA entries)
# ══════════════════════════════════════════════════════════════
heading("13. REFERENCES")
body("(APA 7th Edition Format)", italic=True)
doc.add_paragraph()

refs = [
    "Besta, M., Blach, N., Kubicek, A., Gerstenberger, R., Gianinazzi, L., Gajda, J., … Hoefler, T. (2024). Graph of Thoughts: Solving elaborate problems with large language models. *Proceedings of the 38th AAAI Conference on Artificial Intelligence*. https://doi.org/10.1609/aaai.v38i16.29720",
    "Chase, H. (2022). *LangChain: Building applications with LLMs through composability* [Software]. GitHub. https://github.com/langchain-ai/langchain",
    "Fan, A., Grave, E., & Joulin, A. (2021). Reducing transformer depth on demand with structured dropout. *Proceedings of the 9th International Conference on Learning Representations (ICLR 2021)*.",
    "Gao, Y., Xiong, Y., Gao, X., Jia, K., Pan, J., Bi, Y., … Wang, H. (2023). Retrieval-Augmented Generation for large language models: A survey. *arXiv preprint arXiv:2312.10997*. https://arxiv.org/abs/2312.10997",
    "Izacard, G., & Grave, E. (2021). Leveraging passage retrieval with generative models for open domain question answering. *Proceedings of the 16th Conference of the European Chapter of the Association for Computational Linguistics (EACL 2021)*, 874–880. https://doi.org/10.18653/v1/2021.eacl-main.74",
    "Jiang, Z., Xu, F. F., Gao, L., Sun, Z., Liu, Q., Dwivedi-Yu, J., … Neubig, G. (2023). Active retrieval augmented generation. *Proceedings of the 2023 Conference on Empirical Methods in Natural Language Processing (EMNLP 2023)*, 7969–7992. https://doi.org/10.18653/v1/2023.emnlp-main.495",
    "Johnson, J., Douze, M., & Jégou, H. (2019). Billion-scale similarity search with GPUs. *IEEE Transactions on Big Data*, *7*(3), 535–547. https://doi.org/10.1109/TBDATA.2019.2921572",
    "Karpukhin, V., Oğuz, B., Min, S., Lewis, P., Wu, L., Edunov, S., … Yih, W. (2020). Dense Passage Retrieval for open-domain question answering. *Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing (EMNLP 2020)*, 6769–6781. https://doi.org/10.18653/v1/2020.emnlp-main.550",
    "Lewis, P., Perez, E., Piktus, A., Petroni, F., Karpukhin, V., Goyal, N., … Kiela, D. (2020). Retrieval-Augmented Generation for knowledge-intensive NLP tasks. *Advances in Neural Information Processing Systems*, *33*, 9459–9474.",
    "Ram, O., Levine, Y., Dalmedigos, I., Muhlgay, D., Shashua, A., Leyton-Brown, K., & Shoham, Y. (2023). In-context retrieval-augmented language models. *Transactions of the Association for Computational Linguistics*, *11*, 1316–1331. https://doi.org/10.1162/tacl_a_00605",
    "Reimers, N., & Gurevych, I. (2019). Sentence-BERT: Sentence embeddings using Siamese BERT-Networks. *Proceedings of the 2019 Conference on Empirical Methods in Natural Language Processing and the 9th International Joint Conference on Natural Language Processing (EMNLP-IJCNLP 2019)*, 3982–3992. https://doi.org/10.18653/v1/D19-1410",
    "Shi, F., Chen, X., Misra, K., Scales, N., Dohan, D., Chi, E., … Zhou, D. (2023). Large language models can be easily distracted by irrelevant context. *Proceedings of the 40th International Conference on Machine Learning (ICML 2023)*.",
    "Shi, W., Min, S., Yasunaga, M., Seo, M., James, R., Lewis, M., … Zettlemoyer, L. (2023). REPLUG: Retrieval-Augmented Language Model Pre-Training. *arXiv preprint arXiv:2301.12652*. https://arxiv.org/abs/2301.12652",
    "Touvron, H., Martin, L., Stone, K., Albert, P., Almahairi, A., Babaei, Y., … Scialom, T. (2023). Llama 2: Open foundation and fine-tuned chat models. *arXiv preprint arXiv:2307.09288*. https://arxiv.org/abs/2307.09288",
    "Wang, L., Yang, N., Huang, X., Jiao, B., Yang, L., Jiang, D., … Wei, F. (2022). Text embeddings by weakly-supervised contrastive pre-training. *arXiv preprint arXiv:2212.03533*. https://arxiv.org/abs/2212.03533",
    "Zhao, W. X., Zhou, K., Li, J., Tang, T., Wang, X., Hou, Y., … Wen, J.-R. (2023). A survey of large language models. *arXiv preprint arXiv:2303.18223*. https://arxiv.org/abs/2303.18223",
    "Zhu, Y., Yuan, H., Wang, S., Liu, J., Liu, W., Deng, C., … Wen, J.-R. (2023). Large language models for information retrieval: A survey. *arXiv preprint arXiv:2308.07107*. https://arxiv.org/abs/2308.07107",
]

for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph(style='List Number')
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # Detect italic markers (*...*) and render accordingly
    parts = ref.split('*')
    run_italic = False
    for part in parts:
        run = p.add_run(part)
        run.font.size = Pt(11)
        run.italic = run_italic
        run_italic = not run_italic

# ── Save ─────────────────────────────────────────────────────
output_path = "/home/user/RAG_chatbot/RAG_Chatbot_Report.docx"
doc.save(output_path)
print(f"Document saved: {output_path}")
