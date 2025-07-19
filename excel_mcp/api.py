from fastapi import FastAPI, Form, Request
from fastapi.responses import JSONResponse
import requests
import os
import uuid

app = FastAPI()
TEMP_DIR = "/tmp"

ALLOWED_PATHS = ["/process_excel"]

@app.middleware("http")
async def block_other_paths(request: Request, call_next):
    if request.url.path not in ALLOWED_PATHS:
        return JSONResponse(status_code=403, content={"error": "接口不允许"})
    return await call_next(request)

# 如无其它业务逻辑，process_excel 接口已移除 