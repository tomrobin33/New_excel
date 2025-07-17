from fastapi import FastAPI, Form
import requests
import os
import uuid
import pandas as pd

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
    # 2. 读取为DataFrame
    df = pd.read_excel(excel_path)
    # 3. 删除本地临时文件
    os.remove(excel_path)
    # 4. 返回原始数据（json格式，便于大模型处理）
    return {"data": df.to_dict(orient="records")} 