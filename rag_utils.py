import os
import hashlib
import pickle
import time
import pandas as pd
from openai import AzureOpenAI
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential

# Embedding cache directory
EMBED_CACHE_DIR = os.path.join(os.getcwd(), 'embed_cache')
os.makedirs(EMBED_CACHE_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def file_md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def row_md5(row):
    return hashlib.md5(str(row.values).encode('utf-8')).hexdigest()

def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def embed(text, client, model, logger=None, cache=None, row_key=None):
    if cache is not None and row_key is not None and row_key in cache:
        if logger:
            logger.info(f"Embedding cache hit for row {row_key}")
        return cache[row_key]
    if logger:
        logger.info(f"Embedding generation started for text of length {len(text)}")
    start = time.time()
    try:
        response = client.embeddings.create(
            input=[text],
            model=model
        )
        if hasattr(response, 'data') and response.data and hasattr(response.data[0], 'embedding'):
            duration = time.time() - start
            if logger:
                logger.info(f"Embedding generated in {duration:.2f}s")
            emb = response.data[0].embedding
            if cache is not None and row_key is not None:
                cache[row_key] = emb
            return emb
        else:
            if logger:
                logger.error("Embedding response missing data.")
            raise ValueError("Embedding response missing data.")
    except Exception as e:
        if logger:
            logger.error(f"Embedding generation failed: {e}")
        raise

def index_excel(file_path, client, search_client, embed_model, logger=None):
    if logger:
        logger.info(f"Indexing started for file: {file_path}")
    start = time.time()
    if not os.path.exists(file_path):
        if logger:
            logger.error(f"File not found: {file_path}")
        raise FileNotFoundError(f"File not found: {file_path}")
    df = pd.read_excel(file_path)
    if df.empty:
        if logger:
            logger.error("Excel file is empty.")
        raise ValueError("Excel file is empty.")
    docs = []
    file_hash = file_md5(file_path)
    cache_path = os.path.join(EMBED_CACHE_DIR, f"{file_hash}.pkl")
    if os.path.exists(cache_path):
        with open(cache_path, 'rb') as f:
            embed_cache = pickle.load(f)
        if logger:
            logger.info(f"Loaded embedding cache for file {file_path}")
    else:
        embed_cache = {}
    updated = False
    for i, row in df.iterrows():
        content = " ".join([str(val) for val in row.values])
        row_key = row_md5(row)
        try:
            emb_start = time.time()
            embedding = embed(content, client, embed_model, logger=logger, cache=embed_cache, row_key=row_key)
            emb_end = time.time()
            if logger:
                logger.info(f"Row {i}: Embedding time: {emb_end - emb_start:.2f}s")
            if row_key not in embed_cache:
                updated = True
        except Exception as e:
            if logger:
                logger.error(f"Row {i}: Embedding failed: {e}")
            continue
        doc = {
            "id": str(i),
            "content": content,
            "embedding": embedding
        }
        if 'project_name' in row.index:
            doc['project_name'] = str(row['project_name'])
        docs.append(doc)
    if updated:
        with open(cache_path, 'wb') as f:
            pickle.dump(embed_cache, f)
        if logger:
            logger.info(f"Updated embedding cache for file {file_path}")
    if docs:
        upload_start = time.time()
        search_client.upload_documents(documents=docs)
        upload_end = time.time()
        if logger:
            logger.info(f"Indexed {len(docs)} documents in {upload_end - upload_start:.2f}s")
    else:
        if logger:
            logger.warning("No documents to index.")
        raise ValueError("No documents to index.")
    total = time.time() - start
    if logger:
        logger.info(f"Indexing completed for file: {file_path} in {total:.2f}s")
