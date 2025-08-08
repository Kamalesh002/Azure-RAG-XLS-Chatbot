import logging, os, json
import time
from rag_utils import (
    file_md5, row_md5, allowed_file, embed, index_excel, EMBED_CACHE_DIR
)
import cProfile
import pstats
import io as sysio
from opencensus.ext.azure.log_exporter import AzureLogHandler
from opencensus.ext.azure.trace_exporter import AzureExporter
from opencensus.trace.samplers import ProbabilitySampler
from opencensus.trace.tracer import Tracer
import pandas as pd
from flask import Flask, request, jsonify, render_template_string
from openai import AzureOpenAI
from azure.search.documents import SearchClient
from azure.search.documents.indexes import SearchIndexClient
from azure.search.documents.indexes.models import (
    SearchIndex, SearchField, SearchFieldDataType, SimpleField, VectorSearch, VectorSearchAlgorithmConfiguration
)
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv


load_dotenv()

# Set up Application Insights logging
if "APPINSIGHTS_INSTRUMENTATIONKEY" in os.environ:
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    logger.addHandler(
        AzureLogHandler(
            connection_string=f'InstrumentationKey={os.environ["APPINSIGHTS_INSTRUMENTATIONKEY"]}'
        )
    )
    tracer = Tracer(
        exporter=AzureExporter(
            connection_string=f'InstrumentationKey={os.environ["APPINSIGHTS_INSTRUMENTATIONKEY"]}'
        ),
        sampler=ProbabilitySampler(1.0),
    )
else:
    logger = logging.getLogger(__name__)
    tracer = None

from werkzeug.utils import secure_filename

UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

#def file_md5(file_path):
#def row_md5(row):


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route("/upload", methods=["POST"])
def upload_excel():
    start_time = time.time()
    if 'file' not in request.files:
        logger.warning("No file part in the request")
        return jsonify({"error": "No file part in the request"}), 400
    file = request.files['file']
    if file.filename == '':
        logger.warning("No selected file")
        return jsonify({"error": "No selected file"}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        logger.info(f"File uploaded: {filename}")
        try:
            index_start = time.time()
            index_excel(file_path, client, search_client, AZURE_OPENAI_EMBED_DEPLOYMENT, logger=logger)
            index_end = time.time()
            logger.info(f"File indexed: {filename} (Indexing time: {index_end - index_start:.2f}s)")
            total_time = time.time() - start_time
            logger.info(f"Upload+Index total time for {filename}: {total_time:.2f}s")
            return jsonify({"message": "File uploaded and indexed successfully."})
        except Exception as e:
            logger.error(f"Indexing failed for {filename}: {e}")
            return jsonify({"error": str(e)}), 500
    else:
        logger.warning("Invalid file type upload attempt")
        return jsonify({"error": "Invalid file type. Only .xlsx and .xls allowed."}), 400

AZURE_SEARCH_ENDPOINT = os.environ["AZURE_SEARCH_ENDPOINT"]
AZURE_SEARCH_KEY = os.environ["AZURE_SEARCH_KEY"]
AZURE_SEARCH_INDEX = "excel-index"

AZURE_OPENAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"]
AZURE_OPENAI_KEY = os.environ["AZURE_OPENAI_KEY"]
AZURE_OPENAI_DEPLOYMENT = os.environ["AZURE_OPENAI_DEPLOYMENT"]
AZURE_OPENAI_EMBED_DEPLOYMENT = os.environ["AZURE_OPENAI_EMBED_DEPLOYMENT"]

client = AzureOpenAI(
    api_key=AZURE_OPENAI_KEY,
    api_version="2024-12-01-preview",
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
)


search_client = SearchClient(
    endpoint=AZURE_SEARCH_ENDPOINT,
    index_name=AZURE_SEARCH_INDEX,
    credential=AzureKeyCredential(AZURE_SEARCH_KEY)
)



# Ensure the index exists at startup (after function definition)
def create_index():
    fields = [
        SimpleField(name="id", type=SearchFieldDataType.String, key=True),
        SimpleField(name="content", type=SearchFieldDataType.String, filterable=False, searchable=True),
        SearchField(name="embedding", type=SearchFieldDataType.Collection(SearchFieldDataType.Double),
                    vector_search_dimensions=1536, vector_search_configuration="default")
    ]
    index = SearchIndex(
        name=AZURE_SEARCH_INDEX,
        fields=fields,
        vector_search=VectorSearch(
            algorithm_configurations=[
                VectorSearchAlgorithmConfiguration(
                    name="default",
                    kind="hnsw",
                    parameters={"m": 4, "efConstruction": 400}
                )
            ]
        )
    )
    SearchIndexClient(AZURE_SEARCH_ENDPOINT, AzureKeyCredential(AZURE_SEARCH_KEY)).create_or_update_index(index)

try:
    create_index()
except Exception as e:
    logging.warning(f"Could not create index at startup: {e}")







# Endpoint to list unique project names (if present in the Excel)
@app.route("/projects", methods=["GET"])
def list_projects():
    try:
        # Try to get all docs and extract unique project names
        results = search_client.search(search_text="*", top=1000)
        project_names = set()
        for doc in results:
            if "project_name" in doc:
                project_names.add(doc["project_name"])
        return jsonify({"projects": sorted(project_names)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def search(query):
    query_vector = embed(query, client, AZURE_OPENAI_EMBED_DEPLOYMENT, logger=logger)
    try:
        results = search_client.search(search_text=None, vector=query_vector, top=3, vector_fields="embedding")
        return " ".join([doc["content"] for doc in results])
    except TypeError as e:
        # fallback for older SDKs or if vector search fails
        results = search_client.search(search_text=query, top=3)
        return " ".join([doc["content"] for doc in results])

def generate_answer(context, question):
    messages = [
        {"role": "system", "content": "You are a helpful assistant who answers using the provided context."},
        {"role": "user", "content": f"Context: {context}\n\nQuestion: {question}"}
    ]
    response = client.chat.completions.create(
        model=AZURE_OPENAI_DEPLOYMENT,
        messages=messages
    )
    return response.choices[0].message.content


# Inline chatbot UI with Excel upload
@app.route("/")
def index():
    return render_template_string('''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>RAG Q&A Chatbot</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', Arial, sans-serif; background: #f6f8fa; margin: 0; padding: 0; }
        .container { max-width: 520px; margin: 48px auto; background: #fff; border-radius: 18px; box-shadow: 0 4px 32px #0002; padding: 32px 28px 24px 28px; }
        h2 { text-align: center; color: #1a237e; font-weight: 600; margin-bottom: 18px; letter-spacing: 1px; }
        #upload-form { display: flex; align-items: center; gap: 10px; margin-bottom: 22px; justify-content: center; }
        #file { margin-right: 0; font-size: 15px; }
        #upload-form button { background: #1a237e; color: #fff; border: none; border-radius: 6px; padding: 7px 18px; font-size: 15px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
        #upload-form button:hover { background: #3949ab; }
        #upload-status { font-size: 14px; margin-left: 10px; }
        #chat-box { height: 370px; overflow-y: auto; border: 1.5px solid #e3e6f0; border-radius: 12px; padding: 18px 14px; background: #f9fafc; margin-bottom: 18px; box-shadow: 0 1px 4px #0001; }
        .msg { margin: 12px 0; display: flex; }
        .bubble { padding: 11px 18px; border-radius: 18px; max-width: 80%; font-size: 15px; line-height: 1.6; box-shadow: 0 1px 4px #0001; }
        .bubble.user { background: linear-gradient(90deg, #e3f0ff 60%, #c5cae9 100%); color: #1a237e; margin-left: auto; border-bottom-right-radius: 4px; }
        .bubble.bot { background: linear-gradient(90deg, #e8f5e9 60%, #c8e6c9 100%); color: #256029; margin-right: auto; border-bottom-left-radius: 4px; }
        .sender-label { font-size: 12px; font-weight: 600; margin-bottom: 2px; opacity: 0.7; }
        #question-form { display: flex; gap: 10px; margin-top: 0; }
        #question { flex: 1; padding: 10px 12px; border-radius: 8px; border: 1.5px solid #bdbdbd; font-size: 15px; transition: border 0.2s; }
        #question:focus { border: 1.5px solid #1a237e; outline: none; }
        #send-btn { padding: 10px 22px; border: none; border-radius: 8px; background: #1a237e; color: #fff; font-size: 15px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
        #send-btn:hover { background: #3949ab; }
        @media (max-width: 600px) {
            .container { max-width: 98vw; padding: 12px 2vw 12px 2vw; }
            #chat-box { height: 260px; padding: 10px 4px; }
        }
        .clearfix { clear: both; }
    </style>
</head>
<body>
    <div class="container">
        <h2>RAG Q&amp;A Chatbot</h2>
        <form id="upload-form" enctype="multipart/form-data">
            <input type="file" id="file" name="file" accept=".xlsx,.xls" required>
            <button type="submit">Upload Excel</button>
            <span id="upload-status"></span>
        </form>
        <div id="chat-box"></div>
        <form id="question-form">
            <input type="text" id="question" placeholder="Type your question..." autocomplete="off" required />
            <button type="submit" id="send-btn">Send</button>
        </form>
    </div>
    <script>
        const chatBox = document.getElementById('chat-box');
        const questionForm = document.getElementById('question-form');
        const questionInput = document.getElementById('question');
        const uploadForm = document.getElementById('upload-form');
        const uploadStatus = document.getElementById('upload-status');

        function appendMessage(sender, text) {
            const msgDiv = document.createElement('div');
            msgDiv.className = 'msg clearfix';
            const bubble = document.createElement('div');
            bubble.className = 'bubble ' + sender;
            // Add sender label for clarity
            const label = document.createElement('div');
            label.className = 'sender-label';
            label.innerText = sender === 'user' ? 'You' : 'Bot';
            bubble.appendChild(label);
            const textDiv = document.createElement('div');
            textDiv.innerText = text;
            bubble.appendChild(textDiv);
            msgDiv.appendChild(bubble);
            chatBox.appendChild(msgDiv);
            chatBox.scrollTop = chatBox.scrollHeight;
        }

        questionForm.addEventListener('submit', function(e) {
            e.preventDefault();
            const question = questionInput.value.trim();
            if (!question) return;
            appendMessage('user', question);
            questionInput.value = '';
            appendMessage('bot', '...');
            fetch('/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ question })
            })
            .then(res => res.json())
            .then(data => {
                // Remove the last '...' bot message
                chatBox.removeChild(chatBox.lastChild);
                if (data.answer) {
                    appendMessage('bot', data.answer);
                } else if (data.error) {
                    appendMessage('bot', '[Error] ' + data.error);
                }
            })
            .catch(() => {
                chatBox.removeChild(chatBox.lastChild);
                appendMessage('bot', '[Error] Failed to get response.');
            });
        });

        uploadForm.addEventListener('submit', function(e) {
            e.preventDefault();
            const fileInput = document.getElementById('file');
            const file = fileInput.files[0];
            if (!file) return;
            uploadStatus.textContent = 'Uploading...';
            const formData = new FormData();
            formData.append('file', file);
            fetch('/upload', {
                method: 'POST',
                body: formData
            })
            .then(res => res.json())
            .then(data => {
                if (data.message) {
                    uploadStatus.textContent = '✔ ' + data.message;
                } else if (data.error) {
                    uploadStatus.textContent = '✖ ' + data.error;
                }
            })
            .catch(() => {
                uploadStatus.textContent = '✖ Upload failed.';
            });
        });
    </script>
</body>
</html>
    ''')

@app.route("/chat", methods=["POST"])
def chat():
    try:
        data = request.get_json()
        question = data.get("question")
        if not question:
            return jsonify({"error": "Missing 'question' field"}), 400

        context = search(question)
        answer = generate_answer(context, question)
        return jsonify({"answer": answer})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def profile_app():
    pr = cProfile.Profile()
    pr.enable()
    app.run(debug=True)
    pr.disable()
    s = sysio.StringIO()
    ps = pstats.Stats(pr, stream=s).sort_stats('cumtime')
    ps.print_stats(30)  # Top 30 lines
    logger.info("\n" + s.getvalue())

if __name__ == "__main__":
    # To profile, run this file as main. Logs will include cProfile output after shutdown.
    profile_app()
