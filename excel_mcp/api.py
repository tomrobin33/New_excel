from fastapi import FastAPI, Form
import requests
import os
import uuid
import paramiko

app = FastAPI()
TEMP_DIR = "/tmp"

@app.post("/process_excel")
def process_excel(url: str = Form(...)):
    # 1. 本地下载 Excel
    excel_filename = str(uuid.uuid4()) + ".xlsx"
    excel_path = os.path.join(TEMP_DIR, excel_filename)
    r = requests.get(url, stream=True)
    with open(excel_path, 'wb') as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    # 2. 本地处理 Excel（如 pandas 读取、分析、生成结果文件）
    # 这里留空，交由其他MCP处理Word生成
    # 假设已生成 result.docx
    result_filename = "result_" + str(uuid.uuid4()) + ".docx"
    result_path = os.path.join(TEMP_DIR, result_filename)
    # TODO: 由其他MCP生成Word报告并保存到 result_path
    # 这里只做占位
    with open(result_path, 'w') as f:
        f.write('Word报告内容由其他MCP生成')
    # 3. 上传 result.docx 到服务器
    transport = paramiko.Transport(("8.156.74.79", 22))
    transport.connect(username="root", password="zfsZBC123")
    sftp = paramiko.SFTPClient.from_transport(transport)
    if sftp is not None:
        sftp.put(result_path, f"/root/files/{result_filename}")
        sftp.close()
    if transport is not None:
        transport.close()
    # 4. 删除本地临时文件
    os.remove(excel_path)
    os.remove(result_path)
    # 5. 返回公网下载链接
    return {"download_url": f"http://8.156.74.79:8001/{result_filename}"} 