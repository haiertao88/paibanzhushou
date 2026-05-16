#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS 规格书助手 - 网页版
使用 Flask 构建，精简依赖，便于部署
"""

import os
import json
import base64
import hashlib
import uuid
import time
import re
from flask import Flask, render_template_string, request, jsonify, Response, stream_with_context
from werkzeug.utils import secure_filename
import requests
from bs4 import BeautifulSoup
from PIL import Image
import io

# ========================================================
# 配置
# ========================================================
app = Flask(__name__)
app.secret_key = os.urandom(24)

CONFIG_FILE = "config.json"
GALLERY_FILE = "gallery.json"

# 全局配置
settings = {
    "zhipu_key": "",
    "zhipu_model": "glm-4.7-flash",
    "zhipu_thinking": False,
    "baidu_id": "",
    "baidu_key": "",
    "std_cover": "",
    "std_body": "",
    "neu_cover": "",
    "neu_body": "",
    "prod_name": "产品名称"
}
gallery = []

# ========================================================
# 加载/保存配置
# ========================================================
def load_config():
    global settings
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                settings.update(data)
        except:
            pass

def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

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
def resize_image(data, max_size=200):
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

def call_ai(prompt, system="你是资深的技术文档编辑", temperature=0.7):
    """调用智谱AI"""
    if not settings.get("zhipu_key"):
        raise Exception("未配置智谱 API Key")
    
    headers = {
        "Authorization": f"Bearer {settings['zhipu_key']}",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": settings["zhipu_model"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt}
        ],
        "temperature": temperature
    }
    
    if not settings.get("zhipu_thinking"):
        data["thinking"] = {"type": "disabled"}
    
    response = requests.post(
        "https://open.bigmodel.cn/api/paas/v4/chat/completions",
        headers=headers,
        json=data,
        timeout=120
    )
    
    if response.status_code != 200:
        raise Exception(f"API错误: {response.status_code}")
    
    result = response.json()
    return result["choices"][0]["message"]["content"]

def baidu_image_search(keyword, page=0, rn=20):
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
        response = requests.get('https://image.baidu.com/search/acjson', params=params, headers=headers, timeout=15)
        if response.status_code == 200:
            data = response.json()
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

# ========================================================
# HTML 模板
# ========================================================
HTML_TEMPLATE = """
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
            min-height: 100vh;
        }
        .app-container { display: flex; min-height: 100vh; gap: 1px; background-color: #2c3e50; flex-wrap: wrap; }
        .left-panel, .right-panel {
            flex: 1;
            min-width: 350px;
            background: #1e2a3a;
            display: flex;
            flex-direction: column;
            overflow: hidden;
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
            flex-wrap: wrap;
        }
        .tab {
            padding: 10px 16px;
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
        .content-area { flex: 1; overflow-y: auto; padding: 12px; max-height: calc(100vh - 100px); }
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
        .form-row { display: flex; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; align-items: center; }
        label { font-size: 12px; color: #bdc3c7; display: inline-flex; align-items: center; gap: 4px; }
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
        button.warning { background: #f39c12; }
        .toolbar-side {
            position: fixed;
            left: 10px;
            top: 50%;
            transform: translateY(-50%);
            background: #1a252f;
            border-radius: 8px;
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
            grid-template-columns: repeat(auto-fill, minmax(120px, 1fr));
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
        .gallery-item img { width: 100%; height: 100px; object-fit: cover; }
        .gallery-item .desc { padding: 6px; font-size: 11px; text-align: center; background: #1a252f; word-break: break-all; }
        .log-area {
            background: #0f172a;
            border-radius: 8px;
            padding: 8px;
            font-family: 'Consolas', monospace;
            font-size: 11px;
            height: 150px;
            overflow-y: auto;
            margin-top: 12px;
            white-space: pre-wrap;
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
        @media (max-width: 768px) { .toolbar-side { display: none; } }
        .loading { display: inline-block; width: 16px; height: 16px; border: 2px solid #fff; border-radius: 50%; border-top-color: transparent; animation: spin 0.6s linear infinite; vertical-align: middle; margin-right: 8px; }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
<div class="toolbar-side">
    <button class="tool-btn" onclick="callWPS('remove_descriptions')" title="删除冒号后说明">✂️<br>去说明</button>
    <button class="tool-btn" onclick="callWPS('replace_numbers')" title="序号替换为•">🔄<br>去序号</button>
    <button class="tool-btn" onclick="callWPS('remove_blank_lines')" title="删除空行">🗑️<br>去空行</button>
    <button class="tool-btn" onclick="callWPS('bold_prefix')" title="冒号前加粗">🅱️<br>前加粗</button>
</div>
<div class="app-container">
    <div class="left-panel">
        <div class="title-bar">
            <h2>📝 文案助手</h2>
            <div class="title-bar-actions">
                <button class="title-btn" onclick="openSettings()">⚙️</button>
            </div>
        </div>
        <div class="tabs">
            <button class="tab active" data-tab="core">🛠️ 核心操作</button>
            <button class="tab" data-tab="md">Ⓜ️ MD编辑器</button>
            <button class="tab" data-tab="ai">🤖 AI文案</button>
            <button class="tab" data-tab="trans">🌐 翻译</button>
        </div>
        
        <div class="content-area" id="tab-core">
            <div class="collapsible">
                <div class="collapsible-header">🖼️ 背景/封面图片 <span class="toggle">▼</span></div>
                <div class="collapsible-content open">
                    <div class="form-row">
                        <input type="text" id="stdCover" placeholder="标准封面图片URL" style="flex:2">
                        <button onclick="injectBg('stdCover', true)">注入封面</button>
                    </div>
                    <div class="form-row">
                        <input type="text" id="stdBody" placeholder="标准正文图片URL" style="flex:2">
                        <button onclick="injectBg('stdBody', false)">注入正文</button>
                    </div>
                    <button class="success" style="width:100%" onclick="oneKeyBg()">✨ 一键设置背景</button>
                </div>
            </div>
            <div class="collapsible">
                <div class="collapsible-header">🏷️ 产品名称 <span class="toggle">▼</span></div>
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
                        <button onclick="callWPS('transpose_table')">🔄 表格转置</button>
                        <button onclick="callWPS('delete_en_rows')">🗑️ 删除英文行</button>
                        <button onclick="callWPS('delete_cn_rows')">🗑️ 删除中文行</button>
                    </div>
                    <div class="form-row">
                        <button onclick="callWPS('set_a4')">📄 设置A4竖版</button>
                        <button onclick="callWPS('add_titles')">📌 添加标准标题</button>
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
            <label>📋 Markdown 编辑器</label>
            <textarea id="mdEditor" class="md-editor" placeholder="# 标题&#10;支持 Markdown 语法..."></textarea>
            <div class="form-row" style="margin-top:8px">
                <button class="success" style="flex:1" onclick="renderMarkdown()">✨ 预览</button>
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
                    <button style="width:100%" onclick="alert('请在设置中配置智谱API Key后使用')">📂 导入说明文档</button>
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
                    <label><input type="checkbox" id="bilingualTable"> 中英双语对照表格</label>
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
                    <option value="en_zh">英→中</option>
                    <option value="zh_en">中→英</option>
                </select>
                <span>范围:</span>
                <label><input type="radio" name="transScope" value="selection" checked> 选中段落</label>
                <label><input type="radio" name="transScope" value="all"> 全篇</label>
            </div>
            <button class="success" style="width:100%" onclick="startTranslation()">🚀 极速翻译</button>
            <div id="transLog" class="log-area" style="height:150px">💬 翻译引擎就绪...</div>
        </div>
        
        <div class="status-bar" id="statusBar">✅ 就绪</div>
    </div>
    
    <div class="right-panel">
        <div class="title-bar">
            <h2>🖼️ 图片控制台</h2>
        </div>
        <div class="tabs">
            <button class="tab active" data-tab="gallery">🔴 本地图库</button>
            <button class="tab" data-tab="web">🟢 网络搜索</button>
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
    </div>
</div>

<script>
// 全局变量
let galleryData = [];
let selectedGallery = new Set();
let webResults = [];

// 初始化
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
            const target = document.getElementById(`tab-${tabName}`);
            if (target) target.style.display = 'block';
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
    loadConfig();
    loadGallery();
});

// API 调用
async function apiCall(endpoint, method = 'GET', data = null) {
    const options = { method, headers: { 'Content-Type': 'application/json' } };
    if (data) options.body = JSON.stringify(data);
    const response = await fetch(endpoint, options);
    if (!response.ok) throw new Error(await response.text());
    return response.json();
}

// 加载配置
async function loadConfig() {
    try {
        const config = await apiCall('/api/config');
        document.getElementById('stdCover').value = config.std_cover || '';
        document.getElementById('stdBody').value = config.std_body || '';
        document.getElementById('prodName').value = config.prod_name || '';
    } catch(e) { console.error(e); }
}

// 图库
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
    // 更新按钮样式
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
    updateStatus('生成图片排版...');
    try {
        const result = await apiCall('/api/wps/insert-photos', 'POST', {
            image_ids: selectedIds.slice(0, cols), cols, floating, show_border: showBorder
        });
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

// 网络搜索
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
                updateStatus('下载中...');
                const resp = await fetch('/api/web-save', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ url: item.url }) });
                if (resp.ok) { updateStatus('已保存'); loadGallery(); }
                else updateStatus('保存失败', true);
            };
            div.innerHTML = `<img src="${item.url}" onerror="this.src='data:image/svg+xml,%3Csvg xmlns=\\'http://www.w3.org/2000/svg\\' width=\\'100\\' height=\\'100\\' viewBox=\\'0 0 100 100\\'%3E%3Crect fill=\\'%23333\\' width=\\'100\\' height=\\'100\\'/%3E%3Ctext fill=\\'%23666\\' x=\\'50\\' y=\\'50\\' text-anchor=\\'middle\\' dy=\\'.3em\\'%3E加载失败%3C/text%3E%3C/svg%3E'"><div class="desc">${item.desc || '图片'}</div>`;
            container.appendChild(div);
        });
        document.getElementById('webStatus').innerHTML = `✅ 找到 ${result.count} 张图片`;
    } catch(e) { document.getElementById('webStatus').innerHTML = '⚠️ 搜索失败'; }
}

// Markdown 渲染
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

// AI 生成
async function generateAI(lang) {
    const data = {
        lang: lang,
        sections: {
            desc: document.getElementById('secDesc').checked,
            feat: document.getElementById('secFeat').checked,
            spec: document.getElementById('secSpec').checked,
            app: document.getElementById('secApp').checked,
            install: document.getElementById('secInstall').checked,
            usage: document.getElementById('secUsage').checked
        },
        bilingual_table: document.getElementById('bilingualTable').checked
    };
    const logDiv = document.getElementById('aiLog');
    logDiv.innerHTML = '🤖 AI生成中...\\n';
    updateStatus('AI生成中...');
    try {
        const response = await fetch('/api/ai/generate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
        const result = await response.json();
        logDiv.innerHTML += result.content || '生成完成';
        updateStatus('生成完成');
    } catch(e) { logDiv.innerHTML += '\\n❌ 失败: ' + e.message; updateStatus('失败', true); }
}

// 翻译
async function startTranslation() {
    const engine = document.querySelector('input[name="transEngine"]:checked').value;
    const dir = document.getElementById('transDir').value;
    const scope = document.querySelector('input[name="transScope"]:checked').value;
    const logDiv = document.getElementById('transLog');
    logDiv.innerHTML = '🚀 翻译中...\\n';
    updateStatus('翻译中...');
    try {
        const response = await fetch('/api/translate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ engine, direction: dir, scope }) });
        const result = await response.json();
        logDiv.innerHTML += result.message || '翻译完成';
        updateStatus('翻译完成');
    } catch(e) { logDiv.innerHTML += '\\n❌ 失败: ' + e.message; updateStatus('失败', true); }
}

// WPS 操作
async function injectBg(elementId, isCover) {
    const url = document.getElementById(elementId).value;
    if (!url) { alert('请先输入图片URL'); return; }
    updateStatus('注入中...');
    try {
        const result = await apiCall('/api/wps/inject-bg', 'POST', { image_url: url, is_cover: isCover });
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

async function oneKeyBg() {
    updateStatus('一键设置背景...');
    try {
        const result = await apiCall('/api/wps/onekey-bg', 'POST');
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

async function injectName(fontSize) {
    const name = document.getElementById('prodName').value;
    if (!name) { alert('请输入产品名称'); return; }
    updateStatus('注入中...');
    try {
        const result = await apiCall('/api/wps/inject-name', 'POST', { name: name, font_size: fontSize });
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

async function changeShell() {
    updateStatus('换壳中...');
    try {
        const result = await apiCall('/api/wps/change-shell', 'POST');
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

async function applyThemeColor(r, g, b) {
    updateStatus('应用主题色...');
    try {
        const result = await apiCall('/api/wps/theme-color', 'POST', { r, g, b });
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

async function callWPS(action) {
    updateStatus('处理中...');
    try {
        const result = await apiCall(`/api/wps/${action}`, 'POST');
        updateStatus(result.message);
    } catch(e) { updateStatus('失败: ' + e.message, true); }
}

// 设置
function openSettings() {
    const html = `<div id="settingsModal" style="position:fixed; top:50%; left:50%; transform:translate(-50%,-50%); background:#1e2a3a; padding:20px; border-radius:12px; z-index:1000; width:400px; box-shadow:0 4px 20px rgba(0,0,0,0.5)">
        <h3 style="margin-bottom:16px">⚙️ 设置</h3>
        <div class="form-group"><label>智谱 API KEY:</label><input type="password" id="settingZhipuKey" style="width:100%" placeholder="请输入智谱API Key"></div>
        <div class="form-group"><label>百度翻译 APP ID:</label><input type="text" id="settingBaiduId" style="width:100%" placeholder="请输入百度翻译APP ID"></div>
        <div class="form-group"><label>百度翻译 KEY:</label><input type="password" id="settingBaiduKey" style="width:100%" placeholder="请输入百度翻译KEY"></div>
        <div class="form-row"><button onclick="saveSettings()">保存</button><button onclick="closeSettings()">取消</button></div>
    </div>
    <div id="settingsOverlay" style="position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(0,0,0,0.5); z-index:999" onclick="closeSettings()"></div>`;
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
    updateStatus('配置已保存');
}

function closeSettings() {
    document.getElementById('settingsModal')?.remove();
    document.getElementById('settingsOverlay')?.remove();
}

function updateStatus(msg, isError = false) {
    const bar = document.getElementById('statusBar');
    bar.innerHTML = msg;
    bar.style.color = isError ? '#e74c3c' : '#27ae60';
    setTimeout(() => { bar.style.color = '#bdc3c7'; bar.innerHTML = '✅ 就绪'; }, 3000);
}
</script>
</body>
</html>
"""

# ========================================================
# Flask 路由
# ========================================================
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/api/config', methods=['GET'])
def get_config():
    return jsonify(settings)

@app.route('/api/config', methods=['POST'])
def update_config():
    data = request.json
    settings.update(data)
    save_config()
    return jsonify({"success": True})

@app.route('/api/config/keys', methods=['POST'])
def update_keys():
    data = request.json
    if 'zhipu_key' in data:
        settings['zhipu_key'] = data['zhipu_key']
    if 'baidu_id' in data:
        settings['baidu_id'] = data['baidu_id']
    if 'baidu_key' in data:
        settings['baidu_key'] = data['baidu_key']
    save_config()
    return jsonify({"success": True})

@app.route('/api/gallery', methods=['GET'])
def get_gallery():
    return jsonify(gallery)

@app.route('/api/gallery/upload', methods=['POST'])
def upload_image():
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "No file"}), 400
    
    content = file.read()
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
    return jsonify({"success": True, "id": item["id"]})

@app.route('/api/gallery/<item_id>', methods=['DELETE'])
def delete_gallery_item(item_id):
    global gallery
    gallery = [item for item in gallery if item['id'] != item_id]
    save_gallery()
    return jsonify({"success": True})

@app.route('/api/web-search', methods=['POST'])
def web_search():
    data = request.json
    keyword = data.get('keyword', '')
    page = data.get('page', 0)
    results = baidu_image_search(keyword, page)
    return jsonify({"results": results, "count": len(results)})

@app.route('/api/web-save', methods=['POST'])
def save_web_image():
    data = request.json
    url = data.get('url')
    if not url:
        return jsonify({"error": "No URL"}), 400
    
    try:
        response = requests.get(url, timeout=30, headers={'User-Agent': 'Mozilla/5.0'})
        if response.status_code == 200:
            content = response.content
            thumbnail = resize_image(content, 200)
            image_data = base64.b64encode(content).decode()
            item = {
                "id": str(uuid.uuid4()),
                "filename": "web_image.jpg",
                "data": image_data,
                "thumbnail": thumbnail
            }
            gallery.append(item)
            save_gallery()
            return jsonify({"success": True})
    except Exception as e:
        pass
    return jsonify({"error": "Download failed"}), 500

@app.route('/api/ai/generate', methods=['POST'])
def ai_generate():
    data = request.json
    lang = data.get('lang', 'cn')
    
    sections = data.get('sections', {})
    section_list = []
    if sections.get('desc'): section_list.append("产品描述")
    if sections.get('feat'): section_list.append("产品特点")
    if sections.get('spec'): section_list.append("产品指标")
    if sections.get('app'): section_list.append("应用场景")
    if sections.get('install'): section_list.append("安装方式")
    if sections.get('usage'): section_list.append("使用方法")
    
    prompt = f"请根据通信产品知识，生成一份{'中文' if lang == 'cn' else '英文'}规格书，包含以下章节: {', '.join(section_list)}。使用Markdown格式输出。"
    
    try:
        result = call_ai(prompt)
        return jsonify({"content": result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/translate', methods=['POST'])
def translate():
    data = request.json
    engine = data.get('engine', 'baidu')
    direction = data.get('direction', 'en_zh')
    scope = data.get('scope', 'selection')
    
    # 模拟翻译响应
    return jsonify({"message": f"翻译完成 (引擎: {engine}, 方向: {direction}, 范围: {scope})"})

# WPS 操作接口（模拟）
@app.route('/api/wps/inject-bg', methods=['POST'])
def wps_inject_bg():
    data = request.json
    return jsonify({"message": f"已注入{'封面' if data.get('is_cover') else '正文'}背景图"})

@app.route('/api/wps/onekey-bg', methods=['POST'])
def wps_onekey_bg():
    return jsonify({"message": "一键背景设置完成"})

@app.route('/api/wps/inject-name', methods=['POST'])
def wps_inject_name():
    data = request.json
    return jsonify({"message": f"已注入产品名称「{data.get('name')}」"})

@app.route('/api/wps/change-shell', methods=['POST'])
def wps_change_shell():
    return jsonify({"message": "换壳完成"})

@app.route('/api/wps/transpose_table', methods=['POST'])
def wps_transpose_table():
    return jsonify({"message": "表格转置完成"})

@app.route('/api/wps/delete_en_rows', methods=['POST'])
def wps_delete_en_rows():
    return jsonify({"message": "已删除英文行"})

@app.route('/api/wps/delete_cn_rows', methods=['POST'])
def wps_delete_cn_rows():
    return jsonify({"message": "已删除中文行"})

@app.route('/api/wps/set_a4', methods=['POST'])
def wps_set_a4():
    return jsonify({"message": "已设置为A4竖版"})

@app.route('/api/wps/add_titles', methods=['POST'])
def wps_add_titles():
    return jsonify({"message": "已添加标准标题"})

@app.route('/api/wps/theme-color', methods=['POST'])
def wps_theme_color():
    data = request.json
    return jsonify({"message": f"主题色已应用"})

@app.route('/api/wps/remove_descriptions', methods=['POST'])
def wps_remove_descriptions():
    return jsonify({"message": "已删除冒号后说明文字"})

@app.route('/api/wps/replace_numbers', methods=['POST'])
def wps_replace_numbers():
    return jsonify({"message": "序号已替换为项目符号"})

@app.route('/api/wps/remove_blank_lines', methods=['POST'])
def wps_remove_blank_lines():
    return jsonify({"message": "空行已删除"})

@app.route('/api/wps/bold_prefix', methods=['POST'])
def wps_bold_prefix():
    return jsonify({"message": "冒号前关键词已加粗"})

@app.route('/api/wps/insert-photos', methods=['POST'])
def wps_insert_photos():
    data = request.json
    cols = data.get('cols', 1)
    floating = data.get('floating', False)
    return jsonify({"message": f"已生成{cols}宫格{'浮动' if floating else '嵌入式'}图片排版"})

# ========================================================
# 主入口
# ========================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(host='0.0.0.0', port=port, debug=False)
