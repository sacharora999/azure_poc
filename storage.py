"""Blob-backed durable data exchange using a storage connection string."""
import os, json, hashlib
from azure.storage.blob import BlobServiceClient, ContentSettings

_service = None
def service():
    global _service
    if _service is None:
        conn = os.environ['RECON_STORAGE_CONNECTION_STRING']
        _service = BlobServiceClient.from_connection_string(conn)
    return _service

def get(container, key):
    return service().get_blob_client(container, key).download_blob(max_concurrency=1).readall()

def put(container, key, data, content_type='application/octet-stream'):
    service().get_blob_client(container,key).upload_blob(data,overwrite=True,content_settings=ContentSettings(content_type=content_type))
    return key

def stream_lines(container,key):
    """Read line-oriented JSONL incrementally without materializing full blob."""
    downloader=service().get_blob_client(container,key).download_blob(max_concurrency=1)
    pending=b''
    for chunk in downloader.chunks():
        pending+=chunk
        while b'\n' in pending:
            line,pending=pending.split(b'\n',1)
            if line: yield json.loads(line)
    if pending.strip():yield json.loads(pending)

def sha256(data):return hashlib.sha256(data).hexdigest()
