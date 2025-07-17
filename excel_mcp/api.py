from fastapi import FastAPI, Form
import requests
import os
import paramiko
import uuid
import shutil

app = FastAPI()
TEMP_DIR = "/tmp"

def process_file(input_path, output_path):
    # 这里写你的文件处理逻辑，当前为简单复制
    shutil.copy(input_path, output_path)

def upload_file_to_server(local_path, remote_path, hostname, username, password):
    transport = paramiko.Transport((hostname, 22))
    transport.connect(username=username, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    if sftp is not None:
        sftp.put(local_path, remote_path)
        sftp.close()
    if transport is not None:
        transport.close()

@app.post("/process_and_upload")
def process_and_upload(url: str = Form(...)):
    filename = str(uuid.uuid4()) + "_" + url.split("/")[-1]
    local_path = os.path.join(TEMP_DIR, filename)
    r = requests.get(url, stream=True)
    with open(local_path, 'wb') as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    processed_filename = "processed_" + filename
    processed_path = os.path.join(TEMP_DIR, processed_filename)
    process_file(local_path, processed_path)
    remote_path = f"/root/files/{processed_filename}"
    upload_file_to_server(
        processed_path, remote_path,
        "8.156.74.79", "root", "zfsZBC123"
    )
    download_url = f"http://8.156.74.79:8001/{processed_filename}"
    return {"download_url": download_url} 