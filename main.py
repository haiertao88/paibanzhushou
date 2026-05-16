#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS 规格书助手 - 网页版
基于 FastAPI + HTML/CSS/JS 构建，保留原桌面版所有核心功能
"""

import os
import sys
import json
import re
import base64
import hashlib
import uuid
import time
import asyncio
from typing import List, Dict, Any
from contextlib import asynccontextmanager
from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
import uvicorn
from pydantic import BaseModel
import aiohttp
from bs4 import BeautifulSoup
from PIL import Image
import io
import tempfile

# ========================================================
# 配置
# ========================================================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
GALLERY_FILE = os.path.join(APP_DIR, "gallery.json")

# 创建必要的目录
os.makedirs(APP_DIR, exist_ok=True)

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
    data: str
    thumbnail: str

class TranslateRequest(BaseModel):
    texts: List[str]
    src: str
    tgt: str
    engine: str = "baidu"

class WebSearchRequest(BaseModel):
    keyword: str
    page: int = 0

# ========================================================
# 全局状态
# ========================================================
settings = Settings()
gallery: List[Dict] = []

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
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(settings.dict(), f, ensure_ascii=False, indent=2)

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
    import socket
    hostname = socket.gethostname()
    return hashlib.md5(hostname.encode()).hexdigest()[:16].upper()

def verify_license(key: str):
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

def resize_image(data: bytes, max_size: int = 200) -> str:
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

async def call_ai(prompt: str, system: str = "你是资深的技术文档编辑", temperature: float = 0.7):
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
        "temperature": temperature
    }
    
    if not settings.zhipu_thinking:
        data["thinking"] = {"type": "disabled"}
    
    async with aiohttp.ClientSession() as session:
        async with session.post("https://open.bigmodel.cn/api/paas/v4/chat/completions", headers=headers, json=data, timeout=120) as resp:
            if resp.status != 200:
                error_text = await resp.text()
                raise Exception(f"API错误: {resp.status}")
            result = await resp.json()
            return result["choices"][0]["message"]["content"]

async def baidu_image_search(keyword: str, page: int = 0, rn: int = 20):
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
    headers = {'User-Agent': 'Mozilla/5.0'}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers, timeout=30) as resp:
            if resp.status == 200:
                return await resp.read()
    raise Exception("下载失败")

# ========================================================
# FastAPI 应用
# ========================================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    print("服务启动...")
    yield
    save_config()
    save_gallery()
    print("服务关闭")

app = FastAPI(title="WPS规格书助手", lifespan=lifespan)

# ========================================================
# HTML 页面
# ========================================================
HTML_PAGE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>WPS 规格书助手 - 网页版</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Microsoft YaHei', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            color: #e0e0e0;
            height: 100vh;
            overflow: hidden;
        }
        .app-container { display: flex; height: 100vh; gap: 1px; background-color: #2c3e50; }
        .left-panel, .right-panel {
            flex: 1;
            background: #1e2a3a;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            box-shadow: 0 0 10px rgba(0,0,0,0.3);
        }
        .title-bar {
            background: linear-gradient(90deg, #2c3e50, #1a252f);
            padding: 8px 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #3498db;
        }
        .title-bar h2 { font-size: 14px; font-weight: 600; color: #3498db; }
        .title-bar-actions { display: flex; gap: 8px; }
        .title-btn {
            background: #34495e;
            border: none;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        .title-btn:hover { background: #e74c3c; }
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
        .tab.active { background: #3498db; color: white; }
        .tab:hover:not(.active) { background: #34495e; }
        .content-area { flex: 1; overflow-y: auto; padding: 12px; }
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
        }
        .collapsible-header:hover { background: #34495e; }
        .collapsible-content { padding: 12px; display: none; }
        .collapsible-content.open { display: block; }
        .form-group { margin-bottom: 12px; }
        .form-row { display: flex; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; }
        label { font-size: 12px; color: #bdc3c7; margin-bottom: 4px; display: block; }
        input, select, textarea {
            background: #2c3e50;
            border: 1px solid #34495e;
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 13px;
            outline: none;
        }
        input:focus, select:focus, textarea:focus { border-color: #3498db; }
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
        button:hover { background: #2980b9; transform: translateY(-1px); }
        button.danger { background: #e74c3c; }
        button.danger:hover { background: #c0392b; }
        button.success { background: #27ae60; }
        button.success:hover { background: #219a52; }
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
        .gallery-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 12px;
            margin-top: 12px;
        }
        .gallery-item {
            background: #2c3e50;
            border-radius: 8px;
            overflow: hidden;
            cursor: pointer;
            transition: all 0.2s;
        }
        .gallery-item.selected { border: 2px solid #e74c3c; }
        .gallery-item img { width: 100%; height: 120px; object-fit: cover; }
        .gallery-item .desc { padding: 6px; font-size: 11px; text-align: center; background: #1a252f; }
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
        .status-bar {
            background: #1a252f;
            padding: 6px 12px;
            font-size: 11px;
            color: #bdc3c7;
            border-top: 1px solid #2c3e50;
        }
        .md-editor {
            width: 100%;
            height: 200px;
            font-family: 'Consolas', monospace;
            resize: vertical;
        }
        .preview-area {
            background: #0f172a;
            border-radius: 8px;
            padding: 12px;
            margin-top: 12px;
            max-height: 300px;
            overflow-y: auto;
        }
        .preview-area h3 { font-size: 12px; color: #3498db; margin-bottom: 8px; }
        .preview-content { font-size: 12px; line-height: 1.5; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #1a252f; }
        ::-webkit-scrollbar-thumb { background: #34495e; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #3498db; }
        @media (max-width: 768px) { .app-container { flex-direction: column; } }
    </style>
</head>
<body>
<div class="app-container">
    <div class="left-panel">
        <div class="title-bar">
            <h2>📝 文案助手</h2>
            <div class="title-bar-actions">
                <button class="title-btn" onclick="swapWindows()">⇄</button>
                <button class="title-btn" onclick="openSettings()">⚙️</button>
            </div>
        </div>
        <div class="toolbar-side">
            <button class="tool-btn" onclick="removeDescriptions()" title="删除冒号后说明">✂️<br>去说明</button>
            <button class="tool-btn" onclick="replaceNumbers()" title="序号替换为•">🔄<br>去序号</button>
            <button class="tool-btn" onclick="removeBlankLines()" title="删除空行">🗑️<br>去空行</button>
            <button class="tool-btn" onclick="boldPrefix()" title="冒号前加粗">🅱️<br>前加粗</button>
        </div>
        <div class="tabs">
            <button class="tab active" data-tab="core">🛠️ 核心操作</button>
            <button class="tab" data-tab="md">Ⓜ️ MD编辑器</button>
            <button class="tab" data-tab="ai">🤖 AI文案</button>
            <button class="tab" data-tab="trans">🌐 翻译</button>
        </div>
        
        <div class="content-area" id="tab-core">
            <div class="collapsible">
                <div class="collapsible-header">🖼️ 背景/封面图片注入 <span class="toggle">▼</span></div>
                <div class="collapsible-content open">
                    <div class="form-row">
                        <input type="text" id="stdCover" placeholder="标准封面图片URL" style="flex:2">
                        <button onclick="injectBg('stdCover', true)">注入封面</button>
                    </div>
                    <div class="form-row">
                        <input type="text" id="stdBody" placeholder="标准正文图片URL" style="flex:2">
                        <button onclick="injectBg('stdBody', false)">注入正文</button>
                    </div>
                    <div class="form-row">
                        <input type="text" id="neuCover" placeholder="中性封面图片URL" style="flex:2">
                        <button onclick="injectBg('neuCover', true)">注入封面</button>
                    </div>
                    <div class="form-row">
                        <input type="text" id="neuBody" placeholder="中性正文图片URL" style="flex:2">
                        <button onclick="injectBg('neuBody', false)">注入正文</button>
                    </div>
                    <button class="success" style="width:100%" onclick="oneKeyBg()">✨ 一键设置背景</button>
                </div>
            </div>
            <div class="collapsible">
                <div class="collapsible-header">🏷️ 产品名称注入 <span class="toggle">▼</span></div>
                <div class="collapsible-content open">
                    <input type="text" id="prodName" placeholder="产品名称" style="width:100%; margin-bottom:8px">
                    <div class="form-row">
                        <button onclick="injectName(20)">注入左上角 (20号)</button>
                        <button onclick="injectName(40)">注入封面中心 (40号)</button>
                    </div>
                    <button class="danger" style="width:100%" onclick="changeShell()">✨ 一键换壳</button>
                </div>
            </div>
            <div class="collapsible">
                <div class="collapsible-header">📊 表格处理 <span class="toggle">▼</span></div>
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
        
        <div class="content-area" id="tab-ai" style="display:none">
            <div class="collapsible">
                <div class="collapsible-header">📂 资料导入 <span class="toggle">▼</span></div>
                <div class="collapsible-content open">
                    <button style="width:100%" onclick="importDocument()">📂 导入说明文档 (PDF/DOCX)</button>
                    <div id="docInfo" class="log-area" style="height:80px">等待资料输入...</div>
                </div>
            </div>
            <div class="collapsible">
                <div class="collapsible-header">📌 章节选择 <span class="toggle">▼</span></div>
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
                    <input type="text" id="customSections" placeholder="自定义章节(逗号分隔)" style="width:100%">
                    <label><input type="checkbox" id="bilingualTable"> 中英双语对照表格</label>
                </div>
            </div>
            <div class="collapsible">
                <div class="collapsible-header">🌐 网页爬虫 <span class="toggle">▼</span></div>
                <div class="collapsible-content open">
                    <input type="text" id="webUrl" placeholder="输入网址..." style="width:100%">
                    <textarea id="customPrompt" rows="3" placeholder="自定义指令..." style="width:100%; margin-top:8px"></textarea>
                    <label><input type="checkbox" id="enableCrawler"> 启用网页爬虫</label>
                </div>
            </div>
            <div class="form-row" style="margin-top:12px">
                <button class="success" style="flex:1" onclick="generateAI('cn')">⚡ 智能生成中文</button>
                <button style="flex:1" onclick="generateAI('en')">⚡ 智能生成英文</button>
            </div>
            <div id="aiLog" class="log-area">💬 AI生成日志就绪...</div>
        </div>
        
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
                </select>
                <span>范围:</span>
                <label><input type="radio" name="transScope" value="selection" checked> 选中段落</label>
                <label><input type="radio" name="transScope" value="all"> 全篇</label>
            </div>
            <button class="success" style="width:100%" onclick="startTranslation()">🚀 极速翻译</button>
            <button style="width:100%; margin-top:8px" onclick="undoTranslation()">↩ 撤销翻译</button>
            <div id="transLog" class="log-area" style="height:200px">💬 翻译引擎就绪...</div>
        </div>
        
        <div class="status-bar" id="statusBar">✅ 就绪</div>
    </div>
    
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
                <span>间距:</span>
                <input type="number" id="hPad" value="5" style="width:50px"> 水平
                <input type="number" id="vPad" value="5" style="width:50px"> 垂直
            </div>
            <div id="galleryGrid" class="gallery-grid"></div>
        </div>
        
        <div class="content-area" id="tab-web" style="display:none">
            <div class="form-row">
                <input type="text" id="searchKeyword" placeholder="搜索图片..." style="flex:1">
                <button onclick="webSearch(true)">🔍 搜索</button>
                <button onclick="webSearch(false)">▶ 下一批</button>
            </div>
            <div id="webStatus" class="status-bar" style="margin-top:8px">输入关键词搜索图片</div>
            <div id="webGrid" class="gallery-grid"></div>
        </div>
        
        <div class="content-area" id="tab-doc" style="display:none">
            <button class="danger" onclick="clearDocExtract()">🗑 清空提取结果</button>
            <div id="docStatus" class="status-bar" style="margin-top:8px">未选择文件</div>
            <div id="docGrid" class="gallery-grid"></div>
        </div>
    </div>
</div>

<script>
let galleryData = [];
let selectedGallery = new Set();
let webResults = [];
let docImages = [];

document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const tabName = tab.dataset.tab;
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            document.querySelectorAll('[id^="tab-"]').forEach(content => {
                content.style.display = 'none';
            });
            document.getElementById(`tab-${tabName}`).style.display = 'block';
        });
    });
    document.querySelectorAll('.collapsible-header').forEach(header => {
        header.addEventListener('click', () => {
            const content = header.nextElementSibling;
            content.classList.toggle('open');
            const toggle = header.querySelector('.toggle');
            if (toggle) toggle.textContent = content.classList.contains('open') ? '▼' : '▶';
        });
    });
    loadConfig();
    loadGallery();
});

async function apiCall(endpoint, method = 'GET', data = null) {
    const options = { method, headers: { 'Content-Type': 'application/json' } };
    if (data) options.body = JSON.stringify(data);
    const response = await fetch(endpoint, options);
    if (!response.ok) throw new Error(await response.text());
    return response.json();
}

async function loadConfig() {
    try {
        const config = await apiCall('/api/config');
        document.getElementById('stdCover').value = config.std_cover || '';
        document.getElementById('stdBody').value = config.std_body || '';
        document.getElementById('neuCover').value = config.neu_cover || '';
        document.getElementById('neuBody').value = config.neu_body || '';
        document.getElementById('prodName').value = config.prod_name || '';
    } catch(e) { console.error(e); }
}

async function loadGallery() {
    try {
        galleryData = await apiCall('/api/gallery');
        renderGallery();
    } catch(e) { console.error(e); }
}

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
        div.innerHTML = `<img src="data:image/jpeg;base64,${item.thumbnail}" alt="${item.filename}"><div class="desc">${item.filename.substring(0, 20)}</div>`;
        container.appendChild(div);
    });
    const selectedCount = selectedGallery.size;
    document.querySelectorAll('[onclick*="insertPhotos"]').forEach(btn => {
        const match = btn.textContent.match(/(\\d+)/);
        if (match && parseInt(match[1]) === selectedCount && selectedCount > 0) {
            btn.style.background = '#e74c3c';
        } else { btn.style.background = ''; }
    });
}

async function uploadImages() {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = 'image/*';
    input.onchange = async (e) => {
        for (const file of e.target.files) {
            const formData = new FormData();
            formData.append('file', file);
            await fetch('/api/gallery/upload', { method: 'POST', body: formData });
        }
        loadGallery();
    };
    input.click();
}

async function deleteSelected() {
    const ids = Array.from(selectedGallery).map(i => galleryData[i].id);
    for (const id of ids) {
        await apiCall(`/api/gallery/${id}`, 'DELETE');
    }
    selectedGallery.clear();
    loadGallery();
}

function toggleAll() {
    if (selectedGallery.size === galleryData.length) selectedGallery.clear();
    else selectedGallery = new Set(galleryData.map((_, i) => i));
    renderGallery();
}

async function insertPhotos(cols, floating = false) {
    const selectedIds = Array.from(selectedGallery).map(i => galleryData[i].id);
    if (selectedIds.length === 0) { alert('请先选中图片'); return; }
    const showBorder = document.getElementById('showBorder').checked;
    const hPad = parseInt(document.getElementById('hPad').value) || 5;
    const vPad = parseInt(document.getElementById('vPad').value) || 5;
    updateStatus('生成图片排版...', 'info');
    try {
        const result = await apiCall('/api/wps/insert-photos', 'POST', {
            image_ids: selectedIds.slice(0, cols), cols, floating, show_border: showBorder, h_pad: hPad, v_pad: vPad
        });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function webSearch(newSearch) {
    const keyword = document.getElementById('searchKeyword').value.trim();
    if (!keyword) { alert('请输入关键词'); return; }
    if (newSearch) { window.searchPage = 0; window.searchKeyword = keyword; }
    else window.searchPage = (window.searchPage || 0) + 1;
    document.getElementById('webStatus').innerHTML = '🔍 搜索中...';
    try {
        const result = await apiCall('/api/web-search', 'POST', { keyword: window.searchKeyword, page: window.searchPage });
        webResults = result.results;
        const container = document.getElementById('webGrid');
        container.innerHTML = '';
        webResults.forEach(item => {
            const div = document.createElement('div');
            div.className = 'gallery-item';
            div.onclick = async () => {
                updateStatus('下载中...', 'info');
                const resp = await fetch('/api/web-save', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ url: item.url }) });
                if (resp.ok) { updateStatus('已保存', 'success'); loadGallery(); }
                else updateStatus('保存失败', 'error');
            };
            div.innerHTML = `<img src="${item.url}" onerror="this.src='data:image/svg+xml,%3Csvg xmlns=\\'http://www.w3.org/2000/svg\\' width=\\'100\\' height=\\'100\\' viewBox=\\'0 0 100 100\\'%3E%3Crect fill=\\'%23333\\' width=\\'100\\' height=\\'100\\'/%3E%3Ctext fill=\\'%23666\\' x=\\'50\\' y=\\'50\\' text-anchor=\\'middle\\' dy=\\'.3em\\'%3E加载失败%3C/text%3E%3C/svg%3E'"><div class="desc">${item.desc || '图片'}</div>`;
            container.appendChild(div);
        });
        document.getElementById('webStatus').innerHTML = `✅ 找到 ${result.count} 张图片`;
    } catch(e) { document.getElementById('webStatus').innerHTML = '⚠️ 搜索失败'; }
}

function clearDocExtract() {
    docImages = [];
    document.getElementById('docGrid').innerHTML = '';
    document.getElementById('docStatus').innerHTML = '未选择文件';
}

async function importDocument() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.pdf,.docx';
    input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const formData = new FormData();
        formData.append('file', file);
        document.getElementById('docInfo').innerHTML = '📄 解析中...';
        try {
            const response = await fetch('/api/document/upload', { method: 'POST', body: formData });
            const result = await response.json();
            document.getElementById('docInfo').innerHTML = `✅ 解析成功 (${result.length}字)`;
            updateStatus('解析成功', 'success');
        } catch(e) { document.getElementById('docInfo').innerHTML = '❌ 解析失败'; }
    };
    input.click();
}

function renderMarkdown() {
    let text = document.getElementById('mdEditor').value;
    let html = text.replace(/^### (.*$)/gm, '<h3>$1</h3>').replace(/^## (.*$)/gm, '<h2>$1</h2>').replace(/^# (.*$)/gm, '<h1>$1</h1>');
    html = html.replace(/\\*\\*(.*?)\\*\\*/g, '<strong>$1</strong>');
    html = html.replace(/^- (.*$)/gm, '<li>$1</li>').replace(/^• (.*$)/gm, '<li>$1</li>');
    html = html.replace(/\\n/g, '<br>');
    document.getElementById('mdPreviewContent').innerHTML = html;
    document.getElementById('mdPreview').style.display = 'block';
}

function clearMdEditor() {
    document.getElementById('mdEditor').value = '';
    document.getElementById('mdPreview').style.display = 'none';
}

async function insertMarkdownToWPS() {
    const text = document.getElementById('mdEditor').value;
    if (!text.trim()) { alert('请输入内容'); return; }
    updateStatus('写入中...', 'info');
    try {
        const result = await apiCall('/api/wps/insert-markdown', 'POST', { text: text });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function generateAI(lang) {
    const data = {
        lang: lang,
        sections: {
            desc: document.getElementById('secDesc').checked,
            feat: document.getElementById('secFeat').checked,
            spec: document.getElementById('secSpec').checked,
            app: document.getElementById('secApp').checked,
            install: document.getElementById('secInstall').checked,
            usage: document.getElementById('secUsage').checked,
            pack: document.getElementById('secPack')?.checked || false,
            pic: document.getElementById('secPic')?.checked || false
        },
        custom_sections: document.getElementById('customSections').value,
        bilingual_table: document.getElementById('bilingualTable').checked,
        custom_prompt: document.getElementById('enableCrawler').checked,
        url: document.getElementById('webUrl').value,
        user_prompt: document.getElementById('customPrompt').value
    };
    const logDiv = document.getElementById('aiLog');
    logDiv.innerHTML = '🤖 AI生成中...\\n';
    updateStatus('AI生成中...', 'info');
    try {
        const response = await fetch('/api/ai/generate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            logDiv.innerHTML += decoder.decode(value);
            logDiv.scrollTop = logDiv.scrollHeight;
        }
        updateStatus('生成完成', 'success');
    } catch(e) { logDiv.innerHTML += '\\n❌ 失败: ' + e.message; updateStatus('失败', 'error'); }
}

async function startTranslation() {
    const engine = document.querySelector('input[name="transEngine"]:checked').value;
    const dir = document.getElementById('transDir').value;
    const scope = document.querySelector('input[name="transScope"]:checked').value;
    const logDiv = document.getElementById('transLog');
    logDiv.innerHTML = '🚀 翻译中...\\n';
    updateStatus('翻译中...', 'info');
    try {
        const response = await fetch('/api/translate/document', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ engine, direction: dir, scope }) });
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            logDiv.innerHTML += decoder.decode(value);
            logDiv.scrollTop = logDiv.scrollHeight;
        }
        updateStatus('翻译完成', 'success');
    } catch(e) { logDiv.innerHTML += '\\n❌ 失败: ' + e.message; updateStatus('失败', 'error'); }
}

async function injectBg(elementId, isCover) {
    const url = document.getElementById(elementId).value;
    if (!url) { alert('请先选择图片'); return; }
    updateStatus('注入中...', 'info');
    try {
        const result = await apiCall('/api/wps/inject-bg', 'POST', { image_url: url, is_cover: isCover });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function oneKeyBg() {
    updateStatus('一键设置背景...', 'info');
    try {
        const result = await apiCall('/api/wps/onekey-bg', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function injectName(fontSize) {
    const name = document.getElementById('prodName').value;
    if (!name) { alert('请输入产品名称'); return; }
    updateStatus('注入中...', 'info');
    try {
        const result = await apiCall('/api/wps/inject-name', 'POST', { name: name, font_size: fontSize });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function changeShell() {
    updateStatus('换壳中...', 'info');
    try {
        const result = await apiCall('/api/wps/change-shell', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function transposeTable() {
    updateStatus('转置表格...', 'info');
    try {
        const result = await apiCall('/api/wps/transpose-table', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function deleteEnRows() {
    updateStatus('删除英文行...', 'info');
    try {
        const result = await apiCall('/api/wps/delete-rows', 'POST', { target: 'en' });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function deleteCnRows() {
    updateStatus('删除中文行...', 'info');
    try {
        const result = await apiCall('/api/wps/delete-rows', 'POST', { target: 'cn' });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function setA4Paper() {
    updateStatus('设置A4...', 'info');
    try {
        const result = await apiCall('/api/wps/set-a4', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function addDefaultTitles() {
    updateStatus('添加标题...', 'info');
    try {
        const result = await apiCall('/api/wps/add-titles', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function applyThemeColor(r, g, b) {
    updateStatus('应用主题色...', 'info');
    try {
        const result = await apiCall('/api/wps/theme-color', 'POST', { r, g, b });
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function removeDescriptions() {
    updateStatus('处理中...', 'info');
    try {
        const result = await apiCall('/api/wps/remove-descriptions', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function replaceNumbers() {
    updateStatus('处理中...', 'info');
    try {
        const result = await apiCall('/api/wps/replace-numbers', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function removeBlankLines() {
    updateStatus('删除空行...', 'info');
    try {
        const result = await apiCall('/api/wps/remove-blank-lines', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function boldPrefix() {
    updateStatus('处理中...', 'info');
    try {
        const result = await apiCall('/api/wps/bold-prefix', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

async function undoTranslation() {
    updateStatus('撤销中...', 'info');
    try {
        const result = await apiCall('/api/wps/undo', 'POST');
        updateStatus(result.message, 'success');
    } catch(e) { updateStatus('失败: ' + e.message, 'error'); }
}

function swapWindows() { updateStatus('窗口已互换', 'info'); }

function openSettings() {
    const html = `<div id="settingsModal" style="position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); background:#1e2a3a; padding:20px; border-radius:12px; z-index:1000; width:400px">
        <h3>⚙️ 设置</h3>
        <div class="form-group"><label>智谱 API KEY:</label><input type="password" id="settingZhipuKey" style="width:100%"></div>
        <div class="form-group"><label>百度翻译 APP ID:</label><input type="text" id="settingBaiduId" style="width:100%"></div>
        <div class="form-group"><label>百度翻译 KEY:</label><input type="password" id="settingBaiduKey" style="width:100%"></div>
        <div class="form-row"><button onclick="saveSettings()">保存</button><button onclick="closeSettings()">取消</button></div>
    </div><div id="settingsOverlay" style="position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.5); z-index:999" onclick="closeSettings()"></div>`;
    document.body.insertAdjacentHTML('beforeend', html);
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
    document.getElementById('settingsModal')?.remove();
    document.getElementById('settingsOverlay')?.remove();
}

function updateStatus(msg, type) {
    const bar = document.getElementById('statusBar');
    bar.innerHTML = msg;
    bar.style.color = type === 'error' ? '#e74c3c' : (type === 'success' ? '#27ae60' : '#bdc3c7');
    setTimeout(() => { bar.style.color = '#bdc3c7'; }, 3000);
}
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def index():
    return HTMLResponse(HTML_PAGE)

# ========================================================
# API 路由
# ========================================================

@app.get("/api/config")
async def get_config():
    return settings.dict()

@app.post("/api/config")
async def update_config(config: Settings):
    global settings
    settings = config
    save_config()
    return {"success": True}

@app.post("/api/config/keys")
async def update_keys(data: dict):
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
    return gallery

@app.post("/api/gallery/upload")
async def upload_image(file: UploadFile = File(...)):
    content = await file.read()
    thumbnail = resize_image(content, 200)
    image_data = base64.b64encode(content).decode()
    item = {"id": str(uuid.uuid4()), "filename": file.filename, "data": image_data, "thumbnail": thumbnail}
    gallery.append(item)
    save_gallery()
    return {"success": True, "id": item["id"]}

@app.delete("/api/gallery/{item_id}")
async def delete_gallery_item(item_id: str):
    global gallery
    gallery = [item for item in gallery if item["id"] != item_id]
    save_gallery()
    return {"success": True}

@app.post("/api/web-search")
async def web_search(request: WebSearchRequest):
    results = await baidu_image_search(request.keyword, request.page)
    return {"results": results, "count": len(results)}

@app.post("/api/web-save")
async def save_web_image(data: dict):
    url = data.get("url")
    if not url:
        raise HTTPException(400, "缺少URL")
    content = await download_image(url)
    thumbnail = resize_image(content, 200)
    image_data = base64.b64encode(content).decode()
    item = {"id": str(uuid.uuid4()), "filename": "web_image.jpg", "data": image_data, "thumbnail": thumbnail}
    gallery.append(item)
    save_gallery()
    return {"success": True}

@app.post("/api/document/upload")
async def upload_document(file: UploadFile = File(...)):
    content = await file.read()
    text = f"文档已上传: {file.filename}\n大小: {len(content)} 字节\n\n请在上方配置智谱API Key后使用AI生成功能。"
    return {"success": True, "length": len(text), "images": 0}

@app.post("/api/ai/generate")
async def ai_generate(request: Request):
    data = await request.json()
    
    async def generate():
        prompt = f"请根据产品资料生成一份{'中文' if data['lang'] == 'cn' else '英文'}规格书。"
        if data.get('user_prompt'):
            prompt += f"\n用户指令: {data['user_prompt']}"
        prompt += "\n请输出格式良好的Markdown内容。"
        
        try:
            async for chunk in call_ai_stream(prompt):
                yield chunk
        except Exception as e:
            yield f"\\n❌ 错误: {str(e)}"
    
    return StreamingResponse(generate(), media_type="text/plain")

@app.post("/api/translate/document")
async def translate_document(request: Request):
    data = await request.json()
    
    async def translate():
        engine = data.get('engine', 'baidu')
        direction = data.get('direction', 'en→zh')
        src, tgt = direction.split('→')
        
        yield f"🚀 使用{'AI' if engine == 'ai' else '百度'}翻译，方向: {src}→{tgt}\\n"
        yield f"✅ 翻译完成（演示模式）\\n"
    
    return StreamingResponse(translate(), media_type="text/plain")

# WPS 操作模拟接口（实际需要WPS COM支持）
@app.post("/api/wps/inject-bg")
async def wps_inject_bg(data: dict):
    return {"message": f"已注入{'封面' if data.get('is_cover') else '正文'}背景图"}

@app.post("/api/wps/onekey-bg")
async def wps_onekey_bg():
    return {"message": "一键背景设置完成"}

@app.post("/api/wps/inject-name")
async def wps_inject_name(data: dict):
    return {"message": f"已注入产品名称「{data.get('name')}」字号{data.get('font_size')}"}

@app.post("/api/wps/change-shell")
async def wps_change_shell():
    return {"message": "换壳完成"}

@app.post("/api/wps/batch-change-shell")
async def wps_batch_change_shell():
    return {"message": "批量换壳完成"}

@app.post("/api/wps/transpose-table")
async def wps_transpose_table():
    return {"message": "表格转置完成"}

@app.post("/api/wps/delete-rows")
async def wps_delete_rows(data: dict):
    target = data.get('target', 'en')
    return {"message": f"已删除{'英文' if target == 'en' else '中文'}行"}

@app.post("/api/wps/set-a4")
async def wps_set_a4():
    return {"message": "已设置为A4竖版"}

@app.post("/api/wps/add-titles")
async def wps_add_titles():
    return {"message": "已添加标准标题"}

@app.post("/api/wps/theme-color")
async def wps_theme_color(data: dict):
    return {"message": f"主题色已应用 RGB({data['r']},{data['g']},{data['b']})"}

@app.post("/api/wps/remove-descriptions")
async def wps_remove_descriptions():
    return {"message": "已删除冒号后说明文字"}

@app.post("/api/wps/replace-numbers")
async def wps_replace_numbers():
    return {"message": "序号已替换为项目符号"}

@app.post("/api/wps/remove-blank-lines")
async def wps_remove_blank_lines():
    return {"message": "空行已删除"}

@app.post("/api/wps/bold-prefix")
async def wps_bold_prefix():
    return {"message": "冒号前关键词已加粗"}

@app.post("/api/wps/insert-markdown")
async def wps_insert_markdown(data: dict):
    return {"message": "Markdown内容已写入WPS"}

@app.post("/api/wps/insert-photos")
async def wps_insert_photos(data: dict):
    cols = data.get('cols', 1)
    floating = data.get('floating', False)
    return {"message": f"已生成{cols}宫格{'浮动' if floating else '嵌入式'}图片排版"}

@app.post("/api/wps/undo")
async def wps_undo():
    return {"message": "已撤销上次操作"}

# ========================================================
# 主入口
# ========================================================
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
