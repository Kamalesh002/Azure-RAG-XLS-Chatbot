# 🤖 Intelligent Excel Chatbot with RAG & Azure AI

A sophisticated, web-based Q&A chatbot that allows you to have conversations with your Excel data. This project leverages a **Retrieval-Augmented Generation (RAG)** architecture powered by **Azure OpenAI** and **Azure AI Search** to provide accurate, context-aware answers from your documents.



## ✨ Key Features

* **Chat with Your Data**: Upload any `.xlsx` or `.xls` file and start asking questions about its content in natural language.
* **Retrieval-Augmented Generation (RAG)**: Combines the power of Large Language Models (LLMs) with factual data from your files, reducing hallucinations and providing answers grounded in your specific context.
* **Advanced Semantic Search**: Uses state-of-the-art text embeddings and **Azure AI Search's vector search** capabilities to find the most relevant information to answer your questions.
* **High-Performance & Efficient**: Features an intelligent **embedding cache** that dramatically speeds up the indexing of previously seen or unchanged data, minimizing costs and processing time.
* **Cloud-Native & Scalable**: Built entirely on the robust and scalable **Microsoft Azure** ecosystem.
* **Turnkey Web Interface**: A clean, modern, and responsive UI built with Flask, HTML, and vanilla JavaScript, allowing for easy interaction and file uploads.
* **Built-in Observability**: Integrated with **Azure Application Insights** for enterprise-grade logging, tracing, and performance monitoring.

---

## 🛠️ Tech Stack & Architecture

This project uses a modern, cloud-native stack to deliver a seamless and intelligent experience.

* **Backend**: **Python** with **Flask**
* **AI Services**:
    * **Azure OpenAI Service**: For generating embeddings (`text-embedding-ada-002` or similar) and chat completions (`gpt-4`).
    * **Azure AI Search**: For creating a robust, searchable index with vector search capabilities.
* **Data Processing**: **Pandas** for efficient Excel file manipulation.
* **Frontend**: **HTML**, **CSS**, **JavaScript** (served directly from Flask).
* **Cloud & DevOps**:
    * **Microsoft Azure**
    * **Azure Application Insights**: For logging and performance tracing.
    * **Environment-based Configuration**: Securely manages keys and endpoints via a `.env` file.

### 🏛️ Architectural Flow

The application follows a two-stage RAG process:

1.  **Indexing Flow (Data Ingestion)**
    * A user uploads an Excel file through the web UI.
    * The Flask backend processes the file, reading each row.
    * The content of each row is converted into a vector embedding using Azure OpenAI.
    * To optimize performance and cost, the application first checks a local cache. If an embedding for that exact row already exists (verified via an MD5 hash), the cached version is used, skipping a new API call.
    * The row content and its corresponding embedding are then uploaded to an Azure AI Search index.

2.  **Query Flow (Answering Questions)**
    * A user asks a question in the chat interface.
    * The question is converted into a query vector using the same Azure OpenAI embedding model.
    * This vector is used to perform a similarity search against the Azure AI Search index, retrieving the top most relevant rows (the "context").
    * The retrieved context and the original question are passed to an Azure OpenAI Chat model (GPT-4) with a carefully crafted prompt.
    * The model generates a final, human-readable answer based *only* on the provided context.
    * The answer is streamed back and displayed in the UI.

---
