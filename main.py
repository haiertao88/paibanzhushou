#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS 规格书助手 - 网页版
基于 FastAPI + HTML/CSS/JS 构建，保留原桌面版所有核心功能
功能包括：背景图片注入、产品名称注入、AI文案生成、翻译、图片管理、表格处理等
"""

import os
import sys
import json
import re
import base64
import hashlib
import uuid
import time
import threading
import asyncio
from typing import Optional, List, Dict, Any
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import uvicorn
from pydantic import BaseModel
import aiohttp
import aiofiles
from bs4 import BeautifulSoup
import requests
from PIL import Image
import io
import tempfile
import shutil
from urllib.parse import urlparse

# ========================================================
# 配置
# ========================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(APP_DIR, "static")
TEMPLATES_DIR = os.path.join(APP_DIR, "templates")
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
GALLERY_FILE = os.path.join(APP_DIR, "gallery.json")

# 创建必要的目录
for d in [STATIC_DIR, TEMPLATES_DIR, UPLOAD_DIR]:
    os.makedirs(d, exist_ok=True)

# ========================================================
# 数据模型
# ========================================================
class Settings(BaseModel):
    zhipu_key: str = ""
    zhipu_model: str = "glm-4.7-flash"
    zhipu_thinking: bool = False
    baidu_id: str = ""
    baidu_key: str = ""
    std_cover: str = ""
    std_body: str = ""
    neu_cover: str = ""
    neu_body: str = ""
    prod_name: str = "产品名称"

class GalleryItem(BaseModel):
    id: str
    filename: str
    data: str  # base64
    thumbnail: str  # base64 thumbnail

class TranslateRequest(BaseModel):
    texts: List[str]
    src: str
    tgt: str
    engine: str = "baidu"

class AIWriterRequest(BaseModel):
    lang: str
    raw_text: str = ""
    sections: Dict[str, bool] = {}
    custom_sections: str = ""
    custom_prompt: bool = False
    url: str = ""
    user_prompt: str = ""
    bilingual_table: bool = False

class AIRewriteRequest(BaseModel):
    text: str
    style: str
    global_context: str = ""

class WebSearchRequest(BaseModel):
    keyword: str
    page: int = 0

# ========================================================
# 全局状态
# ========================================================
settings = Settings()
gallery: List[Dict] = []
config_lock = threading.Lock()

# 加载配置
def load_config():
    global settings
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for key, value in data.items():
                    if hasattr(settings, key):
                        setattr(settings, key, value)
        except:
            pass

def save_config():
    with config_lock:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(settings.dict(), f, ensure_ascii=False, indent=2)

# 加载图库
def load_gallery():
    global gallery
    if os.path.exists(GALLERY_FILE):
        try:
            with open(GALLERY_FILE, "r", encoding="utf-8") as f:
                gallery = json.load(f)
        except:
            gallery = []

def save_gallery():
    with open(GALLERY_FILE, "w", encoding="utf-8") as f:
        json.dump(gallery, f, ensure_ascii=False, indent=2)

load_config()
load_gallery()

# ========================================================
# 辅助函数
# ========================================================
def get_machine_code():
    """生成机器码（基于IP和hostname）"""
    import socket
    hostname = socket.gethostname()
    return hashlib.md5(hostname.encode()).hexdigest()[:16].upper()

def verify_license(key: str):
    """验证注册码"""
    if not key:
        return False, "请输入注册码"
    try:
        decoded = base64.b64decode(key.encode()).decode()
        parts = decoded.split('|')
        if len(parts) != 3:
            return False, "注册码格式错误"
        mac, exp_ts, sign = parts
        expected_sign = hashlib.sha256((f"{mac}|{exp_ts}" + "PRO_ASSISTANT_SECRET_2026_FINAL").encode()).hexdigest()[:16]
        if sign != expected_sign:
            return False, "注册码无效或被篡改"
        if mac != get_machine_code():
            return False, "此注册码与当前机器不匹配"
        if float(exp_ts) < time.time():
            return False, "授权已过期"
        return True, "验证通过"
    except Exception:
        return False, "注册码解析失败"

async def fetch_web_content(url: str, max_chars: int = 8000):
    """抓取网页内容"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers, timeout=30) as resp:
                html = await resp.text()
                
                soup = BeautifulSoup(html, 'html.parser')
                
                # 提取表格
                tables = []
                for tbl in soup.find_all('table'):
                    md_rows = []
                    for row in tbl.find_all('tr'):
                        cells = [c.get_text(strip=True).replace('|', '/') for c in row.find_all(['td', 'th'])]
                        if cells:
                            md_rows.append("| " + " | ".join(cells) + " |")
                    if len(md_rows) >= 2:
                        col_count = md_rows[0].count('|') - 1
                        tables.append(md_rows[0] + "\n|" + " | ".join(["---"] * max(col_count, 1)) + "|\n" + "\n".join(md_rows[1:]))
                
                # 提取正文
                for tag in soup(['script', 'style', 'nav', 'footer', 'header', 'iframe']):
                    tag.decompose()
                
                text = soup.get_text(separator='\n', strip=True)
                lines = [l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 2]
                text = '\n'.join(lines)[:max_chars]
                
                return text, '\n'.join(tables)[:3000], ""
    except Exception as e:
        return "", "", str(e)

async def baidu_image_search(keyword: str, page: int = 0, rn: int = 20):
    """百度图片搜索"""
    results = []
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Referer': 'https://image.baidu.com/'
        }
        params = {
            'tn': 'resultjson_com',
            'word': keyword,
            'pn': str(page * rn),
            'rn': str(rn),
        }
        async with aiohttp.ClientSession() as session:
            async with session.get('https://image.baidu.com/search/acjson', params=params, headers=headers, timeout=15) as resp:
                data = await resp.json()
                for item in data.get('data', []):
                    if not isinstance(item, dict):
                        continue
                    thumb = item.get('thumbURL', '') or item.get('middleURL', '')
                    if thumb and thumb.startswith('http'):
                        results.append({
                            'url': thumb,
                            'desc': item.get('fromPageTitleEnc', '')[:25]
                        })
    except Exception as e:
        pass
    return results

async def download_image(url: str) -> bytes:
    """下载图片"""
    headers = {'User-Agent': 'Mozilla/5.0'}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers, timeout=30) as resp:
            if resp.status == 200:
                return await resp.read()
    raise Exception("下载失败")

def resize_image(data: bytes, max_size: int = 200) -> str:
    """调整图片大小并返回base64"""
    try:
        img = Image.open(io.BytesIO(data))
        img = img.convert('RGB')
        ratio = max_size / max(img.size)
        new_size = (int(img.size[0] * ratio), int(img.size[1] * ratio))
        img = img.resize(new_size, Image.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format='JPEG', quality=70)
        return base64.b64encode(buffer.getvalue()).decode()
    except:
        return ""

async def call_ai(prompt: str, system: str = "你是资深的技术文档编辑", temperature: float = 0.7, stream: bool = False):
    """调用智谱AI"""
    if not settings.zhipu_key:
        raise Exception("未配置智谱 API Key")
    
    headers = {
        "Authorization": f"Bearer {settings.zhipu_key}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": settings.zhipu_model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt}
        ],
        "temperature": temperature,
        "stream": stream
    }
    
    if not settings.zhipu_thinking:
        data["thinking"] = {"type": "disabled"}
    
    async with aiohttp.ClientSession() as session:
        async with session.post("https://open.bigmodel.cn/api/paas/v4/chat/completions", headers=headers, json=data, timeout=120) as resp:
            if resp.status != 200:
                error_text = await resp.text()
                raise Exception(f"API错误: {resp.status} - {error_text}")
            result = await resp.json()
            return result["choices"][0]["message"]["content"]

async def call_ai_stream(prompt: str, system: str = "你是资深的技术文档编辑"):
    """流式调用智谱AI"""
    if not settings.zhipu_key:
        raise Exception("未配置智谱 API Key")
    
    headers = {
        "Authorization": f"Bearer {settings.zhipu_key}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": settings.zhipu_model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt}
        ],
        "stream": True
    }
    
    if not settings.zhipu_thinking:
        data["thinking"] = {"type": "disabled"}
    
    async with aiohttp.ClientSession() as session:
        async with session.post("https://open.bigmodel.cn/api/paas/v4/chat/completions", headers=headers, json=data, timeout=120) as resp:
            async for line in resp.content:
                if line:
                    line_str = line.decode().strip()
                    if line_str.startswith("data: "):
                        data_str = line_str[6:]
                        if data_str == "[DONE]":
                            break
                        try:
                            chunk = json.loads(data_str)
                            if chunk["choices"][0]["delta"].get("content"):
                                yield chunk["choices"][0]["delta"]["content"]
                        except:
                            pass

async def call_baidu_translate(texts: List[str], src: str, tgt: str) -> List[str]:
    """百度翻译API"""
    if not settings.baidu_id or not settings.baidu_key:
        raise Exception("未配置百度翻译 API")
    
    import random
    
    results = []
    MAX_BYTES = 4000
    batches = []
    current_batch = []
    current_len = 0
    
    for t in texts:
        t_len = len(t.encode('utf-8'))
        if current_len + t_len > MAX_BYTES and current_batch:
            batches.append(current_batch)
            current_batch = [t]
            current_len = t_len
        else:
            current_batch.append(t)
            current_len += t_len + 1
    
    if current_batch:
        batches.append(current_batch)
    
    for batch in batches:
        query = '\n'.join(batch)
        salt = str(random.randint(32768, 65536))
        sign_str = settings.baidu_id + query + salt + settings.baidu_key
        sign = hashlib.md5(sign_str.encode()).hexdigest()
        
        async with aiohttp.ClientSession() as session:
            async with session.post("https://fanyi-api.baidu.com/api/trans/vip/translate", 
                                   data={"q": query, "from": src, "to": tgt, "appid": settings.baidu_id, "salt": salt, "sign": sign},
                                   timeout=30) as resp:
                data = await resp.json()
                if "error_code" in data:
                    raise Exception(f"翻译错误: {data.get('error_msg', '未知错误')}")
                for item in data.get("trans_result", []):
                    results.append(item.get("dst", ""))
                await asyncio.sleep(0.5)
    
    # 补齐长度
    while len(results) < len(texts):
        results.append("")
    return results[:len(texts)]

async def call_ai_translate(texts: List[str], src: str, tgt: str) -> List[str]:
    """AI翻译"""
    tgt_name = "英文" if tgt == "en" else "中文" if tgt == "zh" else tgt
    
    results = []
    batch_size = 15
    
    for batch_idx in range(0, len(texts), batch_size):
        batch_texts = texts[batch_idx:batch_idx + batch_size]
        input_json = {}
        prefixes = []
        
        for i, text in enumerate(batch_texts):
            prefix_match = re.match(r'^([-*•●■▶◆○—]\s*|\d+[\.、]\s*|[①-⑳]\s*)(.*)', text)
            if prefix_match:
                prefixes.append(prefix_match.group(1))
                content = prefix_match.group(2)
            else:
                prefixes.append("")
                content = text
            
            if content.strip():
                input_json[str(i)] = content
        
        if not input_json:
            for text in batch_texts:
                results.append(text)
            continue
        
        prompt = f"Translate EVERY string value in the following JSON object into {tgt_name}. Requirements:\n1. Maintain precise technical terminology.\n2. Output strictly valid JSON without any markdown formatting.\n3. Keep the exact same JSON keys (0, 1, 2...).\n\n{json.dumps(input_json, ensure_ascii=False)}"
        
        try:
            response = await call_ai(prompt, "You are a professional technical document translator. Output strictly valid JSON.", temperature=0.3)
            
            # 提取JSON
            match = re.search(r'\{.*\}', response, re.DOTALL)
            if match:
                response = match.group(0)
            response = response.replace("```json", "").replace("```", "").strip()
            
            res_json = json.loads(response)
            
            for i, text in enumerate(batch_texts):
                if str(i) in res_json:
                    translated = str(res_json[str(i)]).strip()
                    # 清理外层引号
                    if (translated.startswith('"') and translated.endswith('"')) or \
                       (translated.startswith("'") and translated.endswith("'")):
                        translated = translated[1:-1]
                    results.append(prefixes[i] + translated)
                else:
                    results.append(text)
            
            await asyncio.sleep(0.5)
        except Exception as e:
            for text in batch_texts:
                results.append(text)
    
    return results

# ========================================================
# FastAPI 应用
# ========================================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    # 启动时
    print("服务启动...")
    yield
    # 关闭时
    save_config()
    save_gallery()
    print("服务关闭")

app = FastAPI(title="WPS规格书助手", lifespan=lifespan)

# 静态文件和服务
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """主页面"""
    html_content = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WPS 规格书助手 - 网页版</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Microsoft YaHei', 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            color: #e0e0e0;
            overflow: hidden;
            height: 100vh;
        }
        
        .app-container {
            display: flex;
            height: 100vh;
            gap: 1px;
            background-color: #2c3e50;
        }
        
        /* 左侧面板 - 文案助手 */
        .left-panel {
            flex: 1;
            background: #1e2a3a;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            box-shadow: 2px 0 10px rgba(0,0,0,0.3);
        }
        
        /* 右侧面板 - 图片控制台 */
        .right-panel {
            flex: 1;
            background: #1e2a3a;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            box-shadow: -2px 0 10px rgba(0,0,0,0.3);
        }
        
        /* 标题栏 */
        .title-bar {
            background: linear-gradient(90deg, #2c3e50, #1a252f);
            padding: 8px 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #3498db;
            cursor: move;
            user-select: none;
        }
        
        .title-bar h2 {
            font-size: 14px;
            font-weight: 600;
            color: #3498db;
        }
        
        .title-bar-actions {
            display: flex;
            gap: 8px;
        }
        
        .title-btn {
            background: #34495e;
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            transition: all 0.2s;
        }
        
        .title-btn:hover {
            background: #e74c3c;
        }
        
        /* 标签页 */
        .tabs {
            display: flex;
            background: #2c3e50;
            border-bottom: 1px solid #34495e;
        }
        
        .tab {
            padding: 10px 20px;
            cursor: pointer;
            background: transparent;
            border: none;
            color: #bdc3c7;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.2s;
        }
        
        .tab.active {
            background: #3498db;
            color: white;
        }
        
        .tab:hover:not(.active) {
            background: #34495e;
        }
        
        /* 内容区域 */
        .content-area {
            flex: 1;
            overflow-y: auto;
            padding: 12px;
        }
        
        /* 可折叠面板 */
        .collapsible {
            background: #1a252f;
            border-radius: 8px;
            margin-bottom: 12px;
            overflow: hidden;
        }
        
        .collapsible-header {
            background: #2c3e50;
            padding: 10px 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 600;
            font-size: 13px;
            transition: background 0.2s;
        }
        
        .collapsible-header:hover {
            background: #34495e;
        }
        
        .collapsible-header .icon {
            font-size: 16px;
        }
        
        .collapsible-content {
            padding: 12px;
            display: none;
        }
        
        .collapsible-content.open {
            display: block;
        }
        
        /* 表单元素 */
        .form-group {
            margin-bottom: 12px;
        }
        
        .form-row {
            display: flex;
            gap: 8px;
            margin-bottom: 8px;
            flex-wrap: wrap;
        }
        
        label {
            font-size: 12px;
            color: #bdc3c7;
            margin-bottom: 4px;
            display: block;
        }
        
        input, select, textarea {
            background: #2c3e50;
            border: 1px solid #34495e;
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 13px;
            outline: none;
            transition: all 0.2s;
        }
        
        input:focus, select:focus, textarea:focus {
            border-color: #3498db;
        }
        
        button {
            background: #3498db;
            border: none;
            color: white;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.2s;
        }
        
        button:hover {
            background: #2980b9;
            transform: translateY(-1px);
        }
        
        button.danger {
            background: #e74c3c;
        }
        
        button.danger:hover {
            background: #c0392b;
        }
        
        button.success {
            background: #27ae60;
        }
        
        button.success:hover {
            background: #219a52;
        }
        
        button.warning {
            background: #f39c12;
        }
        
        /* 微调工具栏 */
        .toolbar-side {
            position: absolute;
            left: 0;
            top: 50%;
            transform: translateY(-50%);
            background: #1a252f;
            border-radius: 0 8px 8px 0;
            padding: 8px 4px;
            display: flex;
            flex-direction: column;
            gap: 6px;
            z-index: 100;
        }
        
        .tool-btn {
            width: 44px;
            padding: 8px 0;
            font-size: 11px;
            text-align: center;
            background: #2c3e50;
            border-radius: 6px;
        }
        
        /* 图库网格 */
        .gallery-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 12px;
        }
        
        .gallery-item {
            background: #2c3e50;
            border-radius: 8px;
            overflow: hidden;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
        }
        
        .gallery-item.selected {
            border: 2px solid #e74c3c;
        }
        
        .gallery-item img {
            width: 100%;
            height: 120px;
            object-fit: cover;
        }
        
        .gallery-item .desc {
            padding: 6px;
            font-size: 11px;
            text-align: center;
            background: #1a252f;
        }
        
        /* 进度条 */
        .progress-bar {
            height: 4px;
            background: #2c3e50;
            border-radius: 2px;
            overflow: hidden;
            margin: 8px 0;
        }
        
        .progress-fill {
            height: 100%;
            background: #3498db;
            width: 0%;
            transition: width 0.3s;
        }
        
        /* 日志区域 */
        .log-area {
            background: #0f172a;
            border-radius: 8px;
            padding: 8px;
            font-family: 'Consolas', monospace;
            font-size: 11px;
            height: 120px;
            overflow-y: auto;
            margin-top: 12px;
        }
        
        /* 状态栏 */
        .status-bar {
            background: #1a252f;
            padding: 6px 12px;
            font-size: 11px;
            color: #bdc3c7;
            border-top: 1px solid #2c3e50;
        }
        
        /* 滚动条 */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        
        ::-webkit-scrollbar-track {
            background: #1a252f;
        }
        
        ::-webkit-scrollbar-thumb {
            background: #34495e;
            border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: #3498db;
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .app-container {
                flex-direction: column;
            }
        }
        
        /* 加载动画 */
        .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 0.6s linear infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        /* MD编辑器 */
        .md-editor {
            width: 100%;
            height: 200px;
            font-family: 'Consolas', monospace;
            resize: vertical;
        }
        
        /* 预览区 */
        .preview-area {
            background: #0f172a;
            border-radius: 8px;
            padding: 12px;
            margin-top: 12px;
            max-height: 300px;
            overflow-y: auto;
        }
        
        .preview-area h3 {
            font-size: 12px;
            color: #3498db;
            margin-bottom: 8px;
        }
        
        .preview-content {
            font-size: 12px;
            line-height: 1.5;
        }
        
        .preview-content table {
            border-collapse: collapse;
            width: 100%;
            margin: 8px 0;
        }
        
        .preview-content th, .preview-content td {
            border: 1px solid #34495e;
            padding: 6px;
        }
        
        .preview-content th {
            background: #2c3e50;
        }
    </style>
</head>
<body>
    <div class="app-container">
        <!-- 左侧面板 -->
        <div class="left-panel">
            <div class="title-bar">
                <h2>📝 文案助手</h2>
                <div class="title-bar-actions">
                    <button class="title-btn" onclick="swapWindows()">⇄</button>
                    <button class="title-btn" onclick="openSettings()">⚙️</button>
                </div>
            </div>
            
            <!-- 微调工具栏 -->
            <div class="toolbar-side" id="toolbarSide">
                <button class="tool-btn" onclick="removeDescriptions()" title="删除冒号后说明">✂️<br>去说明</button>
                <button class="tool-btn" onclick="replaceNumbers()" title="序号替换为•">🔄<br>去序号</button>
                <button class="tool-btn" onclick="removeBlankLines()" title="删除空行">🗑️<br>去空行</button>
                <button class="tool-btn" onclick="boldPrefix()" title="冒号前加粗">🅱️<br>前加粗</button>
            </div>
            
            <!-- 标签页 -->
            <div class="tabs">
                <button class="tab active" data-tab="core">🛠️ 核心操作</button>
                <button class="tab" data-tab="md">Ⓜ️ MD编辑器</button>
                <button class="tab" data-tab="ai">🤖 AI文案</button>
                <button class="tab" data-tab="trans">🌐 翻译</button>
            </div>
            
            <!-- 核心操作内容 -->
            <div class="content-area" id="tab-core">
                <!-- 背景图片注入 -->
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">🖼️</span> 背景/封面图片注入
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <div class="form-row">
                            <input type="text" id="stdCover" placeholder="标准封面图片URL" style="flex:2">
                            <button onclick="selectFile('stdCover')">📂 选择</button>
                            <button onclick="injectBg('stdCover', true)">注入</button>
                        </div>
                        <div class="form-row">
                            <input type="text" id="stdBody" placeholder="标准正文图片URL" style="flex:2">
                            <button onclick="selectFile('stdBody')">📂 选择</button>
                            <button onclick="injectBg('stdBody', false)">注入</button>
                        </div>
                        <div class="form-row">
                            <input type="text" id="neuCover" placeholder="中性封面图片URL" style="flex:2">
                            <button onclick="selectFile('neuCover')">📂 选择</button>
                            <button onclick="injectBg('neuCover', true)">注入</button>
                        </div>
                        <div class="form-row">
                            <input type="text" id="neuBody" placeholder="中性正文图片URL" style="flex:2">
                            <button onclick="selectFile('neuBody')">📂 选择</button>
                            <button onclick="injectBg('neuBody', false)">注入</button>
                        </div>
                        <button class="success" style="width:100%" onclick="oneKeyBg()">✨ 一键设置背景</button>
                    </div>
                </div>
                
                <!-- 产品名称注入 -->
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">🏷️</span> 产品名称注入
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <input type="text" id="prodName" placeholder="产品名称" style="width:100%; margin-bottom:8px">
                        <div class="form-row">
                            <button onclick="injectName(20)">注入左上角 (20号)</button>
                            <button onclick="injectName(40)">注入封面中心 (40号)</button>
                        </div>
                        <button class="danger" style="width:100%" onclick="changeShell()">✨ 当前文档：一键换壳</button>
                        <button class="danger" style="width:100%; margin-top:4px" onclick="batchChangeShell()">📑 批量处理：所有打开文档换壳</button>
                    </div>
                </div>
                
                <!-- 表格处理 -->
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">📊</span> 表格处理
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <div class="form-row">
                            <button onclick="transposeTable()">🔄 表格转置</button>
                            <button onclick="deleteEnRows()">🗑️ 删除英文行</button>
                            <button onclick="deleteCnRows()">🗑️ 删除中文行</button>
                        </div>
                        <div class="form-row">
                            <button onclick="setA4Paper()">📄 设置A4竖版</button>
                            <button onclick="addDefaultTitles()">📌 添加标准标题</button>
                        </div>
                        <div class="form-row">
                            <span>主题色：</span>
                            <button onclick="applyThemeColor(41,128,185)" style="background:#2980b9">蓝</button>
                            <button onclick="applyThemeColor(230,126,34)" style="background:#e67e22">橙</button>
                            <button onclick="applyThemeColor(231,76,60)" style="background:#e74c3c">红</button>
                            <button onclick="applyThemeColor(39,174,96)" style="background:#27ae60">绿</button>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- MD编辑器内容 -->
            <div class="content-area" id="tab-md" style="display:none">
                <label>📋 在此粘贴外部复制的文案（支持Markdown）</label>
                <textarea id="mdEditor" class="md-editor" placeholder="# 标题&#10;支持 Markdown 语法..."></textarea>
                <div class="form-row" style="margin-top:8px">
                    <button class="success" style="flex:1" onclick="renderMarkdown()">✨ 预览</button>
                    <button style="flex:1" onclick="insertMarkdownToWPS()">📝 写入WPS</button>
                    <button class="danger" onclick="clearMdEditor()">🗑️ 清空</button>
                </div>
                <div id="mdPreview" class="preview-area" style="display:none">
                    <h3>📄 预览</h3>
                    <div id="mdPreviewContent" class="preview-content"></div>
                </div>
            </div>
            
            <!-- AI文案内容 -->
            <div class="content-area" id="tab-ai" style="display:none">
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">📂</span> 资料导入
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <button style="width:100%" onclick="importDocument()">📂 导入说明文档 (PDF/DOCX)</button>
                        <div class="progress-bar" id="aiProgress" style="display:none">
                            <div class="progress-fill"></div>
                        </div>
                        <div id="docInfo" class="log-area" style="height:80px; margin-top:8px">等待资料输入...</div>
                    </div>
                </div>
                
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">📌</span> 章节选择
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <div class="form-row">
                            <label><input type="checkbox" id="secDesc" checked> 产品描述</label>
                            <label><input type="checkbox" id="secFeat" checked> 产品特点</label>
                            <label><input type="checkbox" id="secSpec" checked> 产品指标</label>
                        </div>
                        <div class="form-row">
                            <label><input type="checkbox" id="secApp" checked> 应用场景</label>
                            <label><input type="checkbox" id="secInstall" checked> 安装方式</label>
                            <label><input type="checkbox" id="secUsage" checked> 使用方法</label>
                        </div>
                        <div class="form-row">
                            <label><input type="checkbox" id="secPack"> 产品包装</label>
                            <label><input type="checkbox" id="secPic"> 产品图片</label>
                        </div>
                        <div class="form-row">
                            <input type="text" id="customSections" placeholder="自定义章节(逗号分隔)" style="flex:1">
                        </div>
                        <label><input type="checkbox" id="bilingualTable"> 生成中文时表格采用中英双语对照</label>
                    </div>
                </div>
                
                <div class="collapsible">
                    <div class="collapsible-header">
                        <span class="icon">🌐</span> 网页爬虫
                        <span class="toggle">▼</span>
                    </div>
                    <div class="collapsible-content open">
                        <input type="text" id="webUrl" placeholder="输入网址..." style="width:100%">
                        <textarea id="customPrompt" rows="3" placeholder="自定义指令..." style="width:100%; margin-top:8px"></textarea>
                        <label><input type="checkbox" id="enableCrawler"> 启用网页爬虫与自定义指令</label>
                    </div>
                </div>
                
                <div class="form-row" style="margin-top:12px">
                    <button class="success" style="flex:1" onclick="generateAI('cn')">⚡ 智能生成中文</button>
                    <button style="flex:1" onclick="generateAI('en')">⚡ 智能生成英文</button>
                </div>
                
                <div id="aiLog" class="log-area">💬 AI生成日志就绪...</div>
            </div>
            
            <!-- 翻译内容 -->
            <div class="content-area" id="tab-trans" style="display:none">
                <div class="form-row">
                    <span>翻译引擎:</span>
                    <label><input type="radio" name="transEngine" value="baidu" checked> 百度翻译</label>
                    <label><input type="radio" name="transEngine" value="ai"> AI翻译</label>
                </div>
                <div class="form-row">
                    <span>方向:</span>
                    <select id="transDir">
                        <option value="en→zh">英→中</option>
                        <option value="zh→en">中→英</option>
                        <option value="auto→zh">自动→中</option>
                        <option value="auto→en">自动→英</option>
                    </select>
                    <span>范围:</span>
                    <label><input type="radio" name="transScope" value="selection" checked> 选中段落</label>
                    <label><input type="radio" name="transScope" value="all"> 全篇</label>
                </div>
                <button class="success" style="width:100%" onclick="startTranslation()">🚀 极速翻译并排版</button>
                <button style="width:100%; margin-top:8px" onclick="undoTranslation()">↩ 撤销翻译</button>
                <div id="transLog" class="log-area" style="height:200px">💬 极速翻译引擎就绪...</div>
            </div>
            
            <div class="status-bar" id="statusBar">
                ✅ 就绪
            </div>
        </div>
        
        <!-- 右侧面板 - 图片控制台 -->
        <div class="right-panel">
            <div class="title-bar">
                <h2>🖼️ 图片控制台</h2>
                <div class="title-bar-actions">
                    <button class="title-btn" onclick="swapWindows()">⇄</button>
                </div>
            </div>
            
            <div class="tabs">
                <button class="tab active" data-tab="gallery">🔴 本地图库</button>
                <button class="tab" data-tab="web">🟢 网络搜索</button>
                <button class="tab" data-tab="doc">🔵 文档提取</button>
            </div>
            
            <!-- 本地图库 -->
            <div class="content-area" id="tab-gallery">
                <div class="form-row">
                    <button onclick="uploadImages()">＋ 上传图片</button>
                    <button class="danger" onclick="deleteSelected()">🗑 删除</button>
                    <button onclick="toggleAll()">☐ 全选</button>
                </div>
                <div class="form-row">
                    <span>排版插入:</span>
                    <button onclick="insertPhotos(1, true)">浮动1框</button>
                    <button onclick="insertPhotos(1)">1框</button>
                    <button onclick="insertPhotos(2)">2框</button>
                    <button onclick="insertPhotos(3)">3框</button>
                    <button onclick="insertPhotos(4)">4框</button>
                    <button onclick="insertPhotos(6)">6框</button>
                </div>
                <div class="form-row">
                    <label><input type="checkbox" id="showBorder" checked> 显示边框</label>
                    <span>间距(磅):</span>
                    <input type="number" id="hPad" value="5" style="width:50px"> 水平
                    <input type="number" id="vPad" value="5" style="width:50px"> 垂直
                </div>
                <div id="galleryGrid" class="gallery-grid">
                    <!-- 图库会动态渲染 -->
                </div>
            </div>
            
            <!-- 网络搜索 -->
            <div class="content-area" id="tab-web" style="display:none">
                <div class="form-row">
                    <input type="text" id="searchKeyword" placeholder="输入关键词搜索图片..." style="flex:1">
                    <button onclick="webSearch(true)">🔍 搜索</button>
                    <button onclick="webSearch(false)">▶ 下一批</button>
                </div>
                <div id="webStatus" class="status-bar" style="margin-top:8px">输入关键词搜索全网图片...</div>
                <div id="webGrid" class="gallery-grid" style="margin-top:8px">
                    <!-- 搜索结果会动态渲染 -->
                </div>
            </div>
            
            <!-- 文档提取 -->
            <div class="content-area" id="tab-doc" style="display:none">
                <div class="form-row">
                    <button class="danger" onclick="clearDocExtract()">🗑 清空提取结果</button>
                </div>
                <div id="docStatus" class="status-bar" style="margin-top:8px">未选择文件</div>
                <div id="docGrid" class="gallery-grid" style="margin-top:8px">
                    <!-- 提取的图片会动态渲染 -->
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // ========================================================
        // 全局变量
        // ========================================================
        let galleryData = [];
        let selectedGallery = new Set();
        let webResults = [];
        let docImages = [];
        let currentTab = 'gallery';
        
        // ========================================================
        // 初始化
        // ========================================================
        document.addEventListener('DOMContentLoaded', () => {
            // 标签页切换
            document.querySelectorAll('.tab').forEach(tab => {
                tab.addEventListener('click', () => {
                    const tabName = tab.dataset.tab;
                    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                    tab.classList.add('active');
                    
                    document.querySelectorAll('[id^="tab-"]').forEach(content => {
                        content.style.display = 'none';
                    });
                    document.getElementById(`tab-${tabName}`).style.display = 'block';
                    currentTab = tabName;
                });
            });
            
            // 可折叠面板
            document.querySelectorAll('.collapsible-header').forEach(header => {
                header.addEventListener('click', () => {
                    const content = header.nextElementSibling;
                    content.classList.toggle('open');
                    const toggle = header.querySelector('.toggle');
                    if (toggle) toggle.textContent = content.classList.contains('open') ? '▼' : '▶';
                });
            });
            
            // 加载配置
            loadConfig();
            loadGallery();
            
            // 监听Ctrl+V粘贴图片
            document.addEventListener('paste', (e) => {
                const items = e.clipboardData.items;
                for (const item of items) {
                    if (item.type.indexOf('image') !== -1) {
                        const file = item.getAsFile();
                        if (file) {
                            uploadPastedImage(file);
                        }
                        break;
                    }
                }
            });
        });
        
        // ========================================================
        // API 调用
        // ========================================================
        async function apiCall(endpoint, method = 'GET', data = null) {
            const options = {
                method,
                headers: {
                    'Content-Type': 'application/json'
                }
            };
            if (data) options.body = JSON.stringify(data);
            
            const response = await fetch(endpoint, options);
            if (!response.ok) {
                const error = await response.text();
                throw new Error(error);
            }
            return response.json();
        }
        
        // 加载配置
        async function loadConfig() {
            try {
                const config = await apiCall('/api/config');
                document.getElementById('stdCover').value = config.std_cover || '';
                document.getElementById('stdBody').value = config.std_body || '';
                document.getElementById('neuCover').value = config.neu_cover || '';
                document.getElementById('neuBody').value = config.neu_body || '';
                document.getElementById('prodName').value = config.prod_name || '';
            } catch (e) {
                console.error('加载配置失败', e);
            }
        }
        
        // 保存配置
        async function saveConfig() {
            const config = {
                std_cover: document.getElementById('stdCover').value,
                std_body: document.getElementById('stdBody').value,
                neu_cover: document.getElementById('neuCover').value,
                neu_body: document.getElementById('neuBody').value,
                prod_name: document.getElementById('prodName').value
            };
            await apiCall('/api/config', 'POST', config);
            updateStatus('配置已保存', 'success');
        }
        
        // 加载图库
        async function loadGallery() {
            try {
                galleryData = await apiCall('/api/gallery');
                renderGallery();
            } catch (e) {
                console.error('加载图库失败', e);
            }
        }
        
        // 渲染图库
        function renderGallery() {
            const container = document.getElementById('galleryGrid');
            if (!container) return;
            
            container.innerHTML = '';
            galleryData.forEach((item, idx) => {
                const div = document.createElement('div');
                div.className = 'gallery-item';
                if (selectedGallery.has(idx)) div.classList.add('selected');
                div.onclick = () => {
                    if (selectedGallery.has(idx)) selectedGallery.delete(idx);
                    else selectedGallery.add(idx);
                    renderGallery();
                };
                div.innerHTML = `
                    <img src="data:image/jpeg;base64,${item.thumbnail}" alt="${item.filename}">
                    <div class="desc">${item.filename.substring(0, 20)}</div>
                `;
                container.appendChild(div);
            });
            
            // 更新按钮样式
            const selectedCount = selectedGallery.size;
            document.querySelectorAll('[onclick*="insertPhotos"]').forEach(btn => {
                const match = btn.textContent.match(/(\\d+)/);
                if (match && parseInt(match[1]) === selectedCount && selectedCount > 0) {
                    btn.style.background = '#e74c3c';
                } else {
                    btn.style.background = '';
                }
            });
        }
        
        // 上传图片
        async function uploadImages() {
            const input = document.createElement('input');
            input.type = 'file';
            input.multiple = true;
            input.accept = 'image/*';
            input.onchange = async (e) => {
                const files = Array.from(e.target.files);
                for (const file of files) {
                    await uploadFile(file);
                }
                loadGallery();
            };
            input.click();
        }
        
        async function uploadPastedImage(file) {
            await uploadFile(file);
            loadGallery();
        }
        
        async function uploadFile(file) {
            const formData = new FormData();
            formData.append('file', file);
            
            const response = await fetch('/api/gallery/upload', {
                method: 'POST',
                body: formData
            });
            if (!response.ok) throw new Error('上传失败');
        }
        
        // 删除选中
        async function deleteSelected() {
            const ids = Array.from(selectedGallery).map(i => galleryData[i].id);
            if (ids.length === 0) return;
            
            for (const id of ids) {
                await apiCall(`/api/gallery/${id}`, 'DELETE');
            }
            selectedGallery.clear();
            loadGallery();
        }
        
        function toggleAll() {
            if (selectedGallery.size === galleryData.length) {
                selectedGallery.clear();
            } else {
                selectedGallery = new Set(galleryData.map((_, i) => i));
            }
            renderGallery();
        }
        
        // 插入图片到WPS
        async function insertPhotos(cols, floating = false) {
            const selectedIds = Array.from(selectedGallery).map(i => galleryData[i].id);
            if (selectedIds.length === 0) {
                alert('请先选中图片');
                return;
            }
            
            const showBorder = document.getElementById('showBorder').checked;
            const hPad = parseInt(document.getElementById('hPad').value) || 5;
            const vPad = parseInt(document.getElementById('vPad').value) || 5;
            
            updateStatus('正在生成图片排版...', 'info');
            try {
                const result = await apiCall('/api/wps/insert-photos', 'POST', {
                    image_ids: selectedIds.slice(0, cols),
                    cols: cols,
                    floating: floating,
                    show_border: showBorder,
                    h_pad: hPad,
                    v_pad: vPad
                });
                updateStatus(result.message || '插入成功', 'success');
            } catch (e) {
                updateStatus('插入失败: ' + e.message, 'error');
            }
        }
        
        // 网络搜索
        async function webSearch(newSearch) {
            const keyword = document.getElementById('searchKeyword').value.trim();
            if (!keyword) {
                alert('请输入关键词');
                return;
            }
            
            if (newSearch) {
                window.searchPage = 0;
                window.searchKeyword = keyword;
            } else {
                window.searchPage = (window.searchPage || 0) + 1;
            }
            
            document.getElementById('webStatus').innerHTML = '🔍 全网检索中...';
            
            try {
                const result = await apiCall('/api/web-search', 'POST', {
                    keyword: window.searchKeyword,
                    page: window.searchPage
                });
                
                webResults = result.results;
                renderWebResults();
                document.getElementById('webStatus').innerHTML = `✅ 找到 ${result.count} 张图片，点击保存到本地图库`;
            } catch (e) {
                document.getElementById('webStatus').innerHTML = '⚠️ 搜索失败: ' + e.message;
            }
        }
        
        function renderWebResults() {
            const container = document.getElementById('webGrid');
            container.innerHTML = '';
            
            webResults.forEach((item, idx) => {
                const div = document.createElement('div');
                div.className = 'gallery-item';
                div.onclick = () => saveWebImage(item.url);
                div.innerHTML = `
                    <img src="${item.url}" onerror="this.src='data:image/svg+xml,%3Csvg xmlns=\'http://www.w3.org/2000/svg\' width=\'100\' height=\'100\' viewBox=\'0 0 100 100\'%3E%3Crect fill=\'%23333\' width=\'100\' height=\'100\'/%3E%3Ctext fill=\'%23666\' x=\'50\' y=\'50\' text-anchor=\'middle\' dy=\'.3em\'%3E加载失败%3C/text%3E%3C/svg%3E'">
                    <div class="desc">${item.desc || '图片'}</div>
                `;
                container.appendChild(div);
            });
        }
        
        async function saveWebImage(url) {
            updateStatus('正在下载图片...', 'info');
            try {
                const response = await fetch('/api/web-save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: url })
                });
                if (response.ok) {
                    updateStatus('图片已保存到本地图库', 'success');
                    loadGallery();
                } else {
                    throw new Error('保存失败');
                }
            } catch (e) {
                updateStatus('保存失败: ' + e.message, 'error');
            }
        }
        
        // WPS操作
        async function injectBg(elementId, isCover) {
            const url = document.getElementById(elementId).value;
            if (!url) {
                alert('请先选择图片');
                return;
            }
            updateStatus(`正在注入${isCover ? '封面' : '正文'}背景...`, 'info');
            try {
                const result = await apiCall('/api/wps/inject-bg', 'POST', {
                    image_url: url,
                    is_cover: isCover
                });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('注入失败: ' + e.message, 'error');
            }
        }
        
        async function oneKeyBg() {
            await saveConfig();
            updateStatus('一键设置背景...', 'info');
            try {
                const result = await apiCall('/api/wps/onekey-bg', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('设置失败: ' + e.message, 'error');
            }
        }
        
        async function injectName(fontSize) {
            const name = document.getElementById('prodName').value;
            if (!name) {
                alert('请输入产品名称');
                return;
            }
            updateStatus('正在注入产品名称...', 'info');
            try {
                const result = await apiCall('/api/wps/inject-name', 'POST', {
                    name: name,
                    font_size: fontSize
                });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('注入失败: ' + e.message, 'error');
            }
        }
        
        async function changeShell() {
            await saveConfig();
            updateStatus('正在执行换壳...', 'info');
            try {
                const result = await apiCall('/api/wps/change-shell', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('换壳失败: ' + e.message, 'error');
            }
        }
        
        async function batchChangeShell() {
            await saveConfig();
            updateStatus('批量换壳中...', 'info');
            try {
                const result = await apiCall('/api/wps/batch-change-shell', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('批量换壳失败: ' + e.message, 'error');
            }
        }
        
        async function transposeTable() {
            updateStatus('正在转置表格...', 'info');
            try {
                const result = await apiCall('/api/wps/transpose-table', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('转置失败: ' + e.message, 'error');
            }
        }
        
        async function deleteEnRows() {
            updateStatus('正在删除英文行...', 'info');
            try {
                const result = await apiCall('/api/wps/delete-rows', 'POST', { target: 'en' });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('删除失败: ' + e.message, 'error');
            }
        }
        
        async function deleteCnRows() {
            updateStatus('正在删除中文行...', 'info');
            try {
                const result = await apiCall('/api/wps/delete-rows', 'POST', { target: 'cn' });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('删除失败: ' + e.message, 'error');
            }
        }
        
        async function setA4Paper() {
            updateStatus('设置纸张...', 'info');
            try {
                const result = await apiCall('/api/wps/set-a4', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('设置失败: ' + e.message, 'error');
            }
        }
        
        async function addDefaultTitles() {
            updateStatus('添加标题...', 'info');
            try {
                const result = await apiCall('/api/wps/add-titles', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('添加失败: ' + e.message, 'error');
            }
        }
        
        async function applyThemeColor(r, g, b) {
            updateStatus('应用主题色...', 'info');
            try {
                const result = await apiCall('/api/wps/theme-color', 'POST', { r, g, b });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('应用失败: ' + e.message, 'error');
            }
        }
        
        async function removeDescriptions() {
            updateStatus('处理中...', 'info');
            try {
                const result = await apiCall('/api/wps/remove-descriptions', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('处理失败: ' + e.message, 'error');
            }
        }
        
        async function replaceNumbers() {
            updateStatus('处理中...', 'info');
            try {
                const result = await apiCall('/api/wps/replace-numbers', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('处理失败: ' + e.message, 'error');
            }
        }
        
        async function removeBlankLines() {
            updateStatus('删除空行...', 'info');
            try {
                const result = await apiCall('/api/wps/remove-blank-lines', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('删除失败: ' + e.message, 'error');
            }
        }
        
        async function boldPrefix() {
            updateStatus('处理中...', 'info');
            try {
                const result = await apiCall('/api/wps/bold-prefix', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('处理失败: ' + e.message, 'error');
            }
        }
        
        // MD编辑器
        function renderMarkdown() {
            const text = document.getElementById('mdEditor').value;
            const previewDiv = document.getElementById('mdPreview');
            const previewContent = document.getElementById('mdPreviewContent');
            
            // 简单的Markdown渲染
            let html = text;
            // 标题
            html = html.replace(/^### (.*$)/gm, '<h3>$1</h3>');
            html = html.replace(/^## (.*$)/gm, '<h2>$1</h2>');
            html = html.replace(/^# (.*$)/gm, '<h1>$1</h1>');
            // 加粗
            html = html.replace(/\\*\\*(.*?)\\*\\*/g, '<strong>$1</strong>');
            // 列表
            html = html.replace(/^- (.*$)/gm, '<li>$1</li>');
            html = html.replace(/^• (.*$)/gm, '<li>$1</li>');
            html = html.replace(/(<li>.*<\\/li>)/s, '<ul>$1</ul>');
            // 表格
            html = html.replace(/\\|(.+)\\|/g, (match) => {
                const cells = match.split('|').filter(c => c.trim());
                return '<tr>' + cells.map(c => `<td>${c.trim()}</td>`).join('') + '</tr>';
            });
            html = html.replace(/(<tr>.*<\\/tr>)/s, '<table>$1</table>');
            // 换行
            html = html.replace(/\\n/g, '<br>');
            
            previewContent.innerHTML = html;
            previewDiv.style.display = 'block';
        }
        
        function clearMdEditor() {
            document.getElementById('mdEditor').value = '';
            document.getElementById('mdPreview').style.display = 'none';
        }
        
        async function insertMarkdownToWPS() {
            const text = document.getElementById('mdEditor').value;
            if (!text.trim()) {
                alert('请输入内容');
                return;
            }
            updateStatus('正在写入WPS...', 'info');
            try {
                const result = await apiCall('/api/wps/insert-markdown', 'POST', { text: text });
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('写入失败: ' + e.message, 'error');
            }
        }
        
        // AI文案生成
        async function importDocument() {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.pdf,.docx';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) return;
                
                const formData = new FormData();
                formData.append('file', file);
                
                document.getElementById('docInfo').innerHTML = '📄 正在解析文档...';
                updateStatus('解析文档中...', 'info');
                
                try {
                    const response = await fetch('/api/document/upload', {
                        method: 'POST',
                        body: formData
                    });
                    const result = await response.json();
                    document.getElementById('docInfo').innerHTML = `✅ 文档解析成功 (${result.length}字)<br>📷 已自动提取 ${result.images || 0} 张图片`;
                    updateStatus('文档解析成功', 'success');
                    loadGallery(); // 刷新图库
                } catch (e) {
                    document.getElementById('docInfo').innerHTML = '❌ 解析失败: ' + e.message;
                    updateStatus('解析失败', 'error');
                }
            };
            input.click();
        }
        
        async function generateAI(lang) {
            const sections = {
                desc: document.getElementById('secDesc').checked,
                feat: document.getElementById('secFeat').checked,
                spec: document.getElementById('secSpec').checked,
                app: document.getElementById('secApp').checked,
                install: document.getElementById('secInstall').checked,
                usage: document.getElementById('secUsage').checked,
                pack: document.getElementById('secPack').checked,
                pic: document.getElementById('secPic').checked
            };
            
            const data = {
                lang: lang,
                sections: sections,
                custom_sections: document.getElementById('customSections').value,
                bilingual_table: document.getElementById('bilingualTable').checked,
                custom_prompt: document.getElementById('enableCrawler').checked,
                url: document.getElementById('webUrl').value,
                user_prompt: document.getElementById('customPrompt').value
            };
            
            const logDiv = document.getElementById('aiLog');
            const progressBar = document.getElementById('aiProgress');
            
            logDiv.innerHTML = '🤖 AI正在生成文案...\\n';
            progressBar.style.display = 'block';
            updateStatus('AI生成中...', 'info');
            
            try {
                const response = await fetch('/api/ai/generate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });
                
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                
                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;
                    const chunk = decoder.decode(value);
                    logDiv.innerHTML += chunk;
                    logDiv.scrollTop = logDiv.scrollHeight;
                }
                
                progressBar.style.display = 'none';
                updateStatus('生成完成', 'success');
            } catch (e) {
                logDiv.innerHTML += '\\n❌ 生成失败: ' + e.message;
                progressBar.style.display = 'none';
                updateStatus('生成失败', 'error');
            }
        }
        
        // 翻译
        async function startTranslation() {
            const engine = document.querySelector('input[name="transEngine"]:checked').value;
            const dir = document.getElementById('transDir').value;
            const scope = document.querySelector('input[name="transScope"]:checked').value;
            
            const logDiv = document.getElementById('transLog');
            logDiv.innerHTML = '🚀 翻译中...\\n';
            updateStatus('翻译中...', 'info');
            
            try {
                const response = await fetch('/api/translate/document', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ engine, direction: dir, scope })
                });
                
                const reader = response.body.getReader();
                const decoder = new TextDecoder();
                
                while (true) {
                    const { done, value } = await reader.read();
                    if (done) break;
                    const chunk = decoder.decode(value);
                    logDiv.innerHTML += chunk;
                    logDiv.scrollTop = logDiv.scrollHeight;
                }
                
                updateStatus('翻译完成', 'success');
            } catch (e) {
                logDiv.innerHTML += '\\n❌ 翻译失败: ' + e.message;
                updateStatus('翻译失败', 'error');
            }
        }
        
        async function undoTranslation() {
            updateStatus('撤销中...', 'info');
            try {
                const result = await apiCall('/api/wps/undo', 'POST');
                updateStatus(result.message, 'success');
            } catch (e) {
                updateStatus('撤销失败: ' + e.message, 'error');
            }
        }
        
        // 文档提取相关
        function clearDocExtract() {
            docImages = [];
            document.getElementById('docGrid').innerHTML = '';
            document.getElementById('docStatus').innerHTML = '未选择文件';
        }
        
        // 通用功能
        function swapWindows() {
            // 网页版中窗口互换效果
            updateStatus('窗口已互换', 'info');
        }
        
        function openSettings() {
            // 打开设置对话框
            const settingsHtml = `
                <div style="position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); background:#1e2a3a; padding:20px; border-radius:12px; z-index:1000; width:400px">
                    <h3>⚙️ 软件设置</h3>
                    <div class="form-group">
                        <label>智谱 API KEY:</label>
                        <input type="password" id="settingZhipuKey" style="width:100%">
                    </div>
                    <div class="form-group">
                        <label>百度翻译 APP ID:</label>
                        <input type="text" id="settingBaiduId" style="width:100%">
                    </div>
                    <div class="form-group">
                        <label>百度翻译 KEY:</label>
                        <input type="password" id="settingBaiduKey" style="width:100%">
                    </div>
                    <div class="form-row">
                        <button onclick="saveSettings()">保存</button>
                        <button onclick="closeSettings()">取消</button>
                    </div>
                </div>
                <div style="position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.5); z-index:999" onclick="closeSettings()"></div>
            `;
            const overlay = document.createElement('div');
            overlay.id = 'settingsOverlay';
            overlay.innerHTML = settingsHtml;
            document.body.appendChild(overlay);
        }
        
        async function saveSettings() {
            const config = {
                zhipu_key: document.getElementById('settingZhipuKey').value,
                baidu_id: document.getElementById('settingBaiduId').value,
                baidu_key: document.getElementById('settingBaiduKey').value
            };
            await apiCall('/api/config/keys', 'POST', config);
            closeSettings();
            updateStatus('配置已保存', 'success');
        }
        
        function closeSettings() {
            const overlay = document.getElementById('settingsOverlay');
            if (overlay) overlay.remove();
        }
        
        function selectFile(elementId) {
            // 文件选择回调（实际需要服务器支持文件上传）
            alert('请直接粘贴图片URL或使用上传功能');
        }
        
        function updateStatus(msg, type = 'info') {
            const statusBar = document.getElementById('statusBar');
            statusBar.innerHTML = msg;
            if (type === 'error') statusBar.style.color = '#e74c3c';
            else if (type === 'success') statusBar.style.color = '#27ae60';
            else statusBar.style.color = '#bdc3c7';
            setTimeout(() => {
                statusBar.style.color = '#bdc3c7';
            }, 3000);
        }
    </script>
</body>
</html>
    """
    return HTMLResponse(html_content)

# ========================================================
# API 路由
# ========================================================

@app.get("/api/config")
async def get_config():
    """获取配置"""
    return settings.dict()

@app.post("/api/config")
async def update_config(config: Settings):
    """更新配置"""
    global settings
    settings = config
    save_config()
    return {"success": True}

@app.post("/api/config/keys")
async def update_keys(data: dict):
    """更新密钥配置"""
    if "zhipu_key" in data:
        settings.zhipu_key = data["zhipu_key"]
    if "baidu_id" in data:
        settings.baidu_id = data["baidu_id"]
    if "baidu_key" in data:
        settings.baidu_key = data["baidu_key"]
    save_config()
    return {"success": True}

@app.get("/api/gallery")
async def get_gallery():
    """获取图库列表"""
    return gallery

@app.post("/api/gallery/upload")
async def upload_image(file: UploadFile = File(...)):
    """上传图片到图库"""
    content = await file.read()
    thumbnail = resize_image(content, 200)
    image_data = base64.b64encode(content).decode()
    
    item = {
        "id": str(uuid.uuid4()),
        "filename": file.filename,
        "data": image_data,
        "thumbnail": thumbnail
    }
    gallery.append(item)
    save_gallery()
    return {"success": True, "id": item["id"]}

@app.delete("/api/gallery/{item_id}")
async def delete_gallery_item(item_id: str):
    """删除图库图片"""
    global gallery
    gallery = [item for item in gallery if item["id"] != item_id]
    save_gallery()
    return {"success": True}

@app.post("/api/web-search")
async def web_search(request: WebSearchRequest):
    """网络图片搜索"""
    results = await baidu_image_search(request.keyword, request.page)
    return {"results": results, "count": len(results)}

@app.post("/api/web-save")
async def save_web_image(data: dict):
    """保存网络图片到图库"""
    url = data.get("url")
    if not url:
        raise HTTPException(400, "缺少URL")
    
    content = await download_image(url)
    thumbnail = resize_image(content, 200)
    image_data = base64.b64encode(content).decode()
    
    item = {
        "id": str(uuid.uuid4()),
        "filename":
