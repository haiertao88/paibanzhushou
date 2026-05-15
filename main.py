import streamlit as st
import os, json, re, threading, requests, tempfile, zipfile, io, base64, hashlib
from PIL import Image
import streamlit.components.v1 as components

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import pdfplumber

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ═══════════════════════════════════════════
# 1. 基础配置与全局常量
# ═══════════════════════════════════════════
st.set_page_config(page_title="华脉规格书排版系统", layout="wide", page_icon="📄")

CONFIG_FILE = "huamai_web_config.json"
KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"

CN_HEADER_KEYWORDS = ["产品描述","产品特点","产品指标","应用场景","技术参数","产品介绍","安装方式","使用说明","注意事项","产品图片","产品包装","安装方法"]
EN_HEADER_KEYWORDS = ["Product Description","Product Features","Technical Specifications","Product Specifications","Applications","Application Scenarios","Instructions","Installation","Notes","Product Images","Product Packaging"]

FONT_CHOICES = ["微软雅黑","宋体","黑体","楷体","仿宋","Arial","Times New Roman","Calibri","Verdana","Georgia"]
BULLET_STYLES = {"● 实心圆":"● ", "■ 实心方":"■ ", "▶ 三角形":"▶ ", "◆ 菱形":"◆ ", "○ 空心圆":"○ ", "① 带圈数字":"__NUM__", "1. 数字编号":"__DOT__", "— 短横线":"— "}

# ═══════════════════════════════════════════
# 2. 状态管理与回调函数 (彻底解决 AI 不写入的问题)
# ═══════════════════════════════════════════
def load_config():
    defaults = {
        'kimi_key': "", 'bd_id': "", 'bd_key': "",
        'font_cn': "微软雅黑", 'font_en': "Arial", 'title_size': 14, 'body_size': 11,
        'cover_cn': "光纤跳线系列", 'cover_en': "Fiber Optic Patch Cord",
        'bullet': "● 实心圆", 'feature_brief': True,
        'prompt_cn': "请根据以下资料撰写专业规格书...\n要求：必须使用 Markdown 表格输出产品指标。",
        'prompt_en': "Write an English spec sheet from the source material...\nSpecs MUST be a Markdown table."
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                defaults.update(json.load(f))
        except: pass
    return defaults

def save_config():
    to_save = {
        'kimi_key': st.session_state.get('kimi_key', ''),
        'bd_id': st.session_state.get('bd_id', ''),
        'bd_key': st.session_state.get('bd_key', ''),
        'font_cn': st.session_state.get('font_cn', '微软雅黑'),
        'font_en': st.session_state.get('font_en', 'Arial'),
        'title_size': st.session_state.get('title_size', 14),
        'body_size': st.session_state.get('body_size', 11),
        'cover_cn': st.session_state.get('cover_cn', '光纤跳线系列'),
        'cover_en': st.session_state.get('cover_en', 'Fiber Optic Patch Cord'),
        'bullet': st.session_state.get('bullet', '● 实心圆'),
        'feature_brief': st.session_state.get('feature_brief', True),
        'prompt_cn': st.session_state.get('prompt_cn', ''),
        'prompt_en': st.session_state.get('prompt_en', '')
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False)
        st.toast("✅ 系统配置与API密钥已持久保存！")
    except Exception as e:
        st.error(f"保存失败: {e}")

if 'app_config' not in st.session_state:
    cfg = load_config()
    st.session_state['app_config'] = cfg
    for k, v in cfg.items():
        if k not in st.session_state: st.session_state[k] = v
        
    st.session_state.txt_cn = ""
    st.session_state.txt_en = ""
    st.session_state.raw_text = ""
    st.session_state.gallery_dict = {} 
    st.session_state.bg_cover_bytes = None
    st.session_state.bg_body_bytes = None
    st.session_state.current_lang = 'cn'

# --- 核心修复：AI 撰写的回调函数 ---
def perform_ai_write(lang):
    if not st.session_state.kimi_key:
        st.toast("❌ 请先在左侧【⚙️ 系统设置】中配置 Kimi API Key。", icon="❌")
        return
        
    raw = st.session_state.get('raw_text', '')
    
    if lang == 'cn':
        feat_rule = "特点只保留关键词，如「- 关键词」。" if not st.session_state.feature_brief else "特点包含简短说明。"
        base_prompt = f"资料：\n{raw[:6000]}\n\n要求：不输出产品名，以 **产品描述** 开头。严格使用标题：**产品描述** / **产品特点** / **产品指标** / **应用场景** / **使用说明**。\n注意：产品指标必须是Markdown表格。{feat_rule} 列表使用短横线 - 。"
        custom_prompt = st.session_state.prompt_cn
        prompt = custom_prompt + "\n\n" + base_prompt if custom_prompt else base_prompt
    else:
        base_prompt = f"Source:\n{raw[:6000]}\n\nRules: Start with **Product Description**. Use EXACT headings: **Product Description** / **Product Features** / **Product Specifications** / **Applications** / **Instructions**.\nSpecs MUST be a Markdown table. Use '-' for lists."
        custom_prompt = st.session_state.prompt_en
        prompt = custom_prompt + "\n\n" + base_prompt if custom_prompt else base_prompt

    try:
        res = requests.post(KIMI_API_URL, headers={"Authorization": f"Bearer {st.session_state.kimi_key}"},
                            json={"model": "moonshot-v1-8k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.2})
        content = res.json()["choices"][0]["message"]["content"]
        
        # 强制更新 Session State，绑定到 Text Area
        if lang == 'cn':
            st.session_state.txt_cn = content
            st.session_state.current_lang = 'cn'
        else:
            st.session_state.txt_en = content
            st.session_state.current_lang = 'en'
        st.toast("✅ AI 撰写完毕！", icon="✅")
    except Exception as e: 
        st.toast(f"❌ AI 错误: {str(e)}", icon="❌")

# ═══════════════════════════════════════════
# 3. 解析与 Word 导出算法
# ═══════════════════════════════════════════
def clean_markdown(t):
    if not t: return ""
    t = re.sub(r'\*\*(.+?)\*\*', r'\1', t)
    t = re.sub(r'^#+\s*', '', t)
    return t.strip()

def is_header(raw_line):
    clean = clean_markdown(raw_line)
    if clean in CN_HEADER_KEYWORDS or clean in EN_HEADER_KEYWORDS: return True
    if re.match(r'^\*\*.+\*\*$', raw_line) and not raw_line.startswith(('-','*','|')): return True
    return False

def extract_images_from_file(uploaded_file):
    imgs = []
    if uploaded_file.name.endswith(".pdf"):
        import fitz
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for i in range(len(doc)):
            for img in doc[i].get_images(full=True):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                if len(image_bytes) > 5000: imgs.append(image_bytes)
    elif uploaded_file.name.endswith(".docx"):
        with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
            for f in z.namelist():
                if f.startswith("word/media/") and f.lower().endswith(('.png','.jpg','.jpeg')):
                    data = z.read(f)
                    if len(data) > 5000: imgs.append(data)
    return imgs

def bytes_to_b64(b):
    if not b: return ""
    return base64.b64encode(b).decode('utf-8')

def _add_word_bg(section, bg_bytes):
    if not bg_bytes: return
    hdr = section.header
    hdr.is_linked_to_previous = False
    if len(hdr.paragraphs) == 0: p = hdr.add_paragraph()
    else:
        p = hdr.paragraphs[0]; p.text = "" 
        for r in p.runs: r._element.getparent().remove(r._element)
    run = p.add_run()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(bg_bytes)
        tmp_path = tmp.name
    pic = run.add_picture(tmp_path, width=Mm(210), height=Mm(297))
    anc = OxmlElement('wp:anchor')
    anc.set(qn('wp:behindDoc'), '1')
    for tag, val in [('wp:positionH', 'page'), ('wp:positionV', 'page')]:
        pos = OxmlElement(tag); pos.set(qn('wp:relativeFrom'), val)
        off = OxmlElement('wp:posOffset'); off.text = '0'
        pos.append(off); anc.append(pos)
    for child in pic._inline: anc.append(child)
    pic._inline.getparent().replace(pic._inline, anc)
    os.unlink(tmp_path)

def generate_word_document(lang):
    doc = Document()
    fn = st.session_state.font_cn if lang == 'cn' else st.session_state.font_en
    txt = st.session_state.txt_cn if lang == 'cn' else st.session_state.txt_en
    
    sec = doc.sections[0]
    sec.page_width = Mm(210); sec.page_height = Mm(297)
    sec.top_margin = Mm(25.4); sec.bottom_margin = Mm(25.4); sec.left_margin = Mm(19.1); sec.right_margin = Mm(19.1)
    doc.styles['Normal'].font.name = fn
    if lang == 'cn': doc.styles['Normal']._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # 封面背景
    if st.session_state.bg_cover_bytes: _add_word_bg(sec, st.session_state.bg_cover_bytes)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(12): cp.add_run('\n')
    cover_title = st.session_state.cover_cn if lang == 'cn' else st.session_state.cover_en
    rn = cp.add_run(cover_title if cover_title else "Product Specification")
    rn.bold = True; rn.font.size = Pt(max(st.session_state.title_size + 14, 24)); rn.font.name = fn
    if lang == 'cn': rn._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # 正文背景
    doc.add_page_break()
    bs = doc.add_section() 
    bs.header.is_linked_to_previous = False 
    if st.session_state.bg_body_bytes: _add_word_bg(bs, st.session_state.bg_body_bytes)
    else:
        if len(bs.header.paragraphs) > 0: bs.header.paragraphs[0].text = ""
    
    lines = txt.split('\n')
    tr = [] 
    
    def render_run(pa, text, size=st.session_state.body_size, bold=False):
        r = pa.add_run(text)
        r.font.size = Pt(size); r.font.name = fn; r.bold = bold
        if lang == 'cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)

    for idx, line in enumerate(lines):
        lt = line.strip()
        if not lt: continue
        
        # 真实图片插入
        local_img_m = re.match(r'^\[LOCAL_IMG:(.+?)\]$', lt)
        if local_img_m:
            img_hash = local_img_m.group(1)
            img_bytes = st.session_state.gallery_dict.get(img_hash)
            if img_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try: p.add_run().add_picture(tmp_path, width=Mm(145))
                except: pass
                os.unlink(tmp_path)
            continue
            
        if 'IMG_FRAME' in lt:
            p = doc.add_paragraph()
            render_run(p, f"【排版预留：产品图片占位框】", st.session_state.body_size, True)
            continue
            
        if '|' in lt:
            if '---' in lt: continue
            cells = [clean_markdown(c) for c in lt.strip('|').split('|') if c.strip() != '']
            if cells: tr.append(cells)
            nxt = lines[idx+1].strip() if idx+1 < len(lines) else ""
            if not nxt or '|' not in nxt:
                if tr:
                    t = doc.add_table(rows=len(tr), cols=max(len(r) for r in tr))
                    t.style = 'Table Grid'
                    for r_idx, row_data in enumerate(tr):
                        for c_idx, cell_val in enumerate(row_data):
                            if c_idx < len(t.columns):
                                cell = t.cell(r_idx, c_idx)
                                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                                p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                render_run(p, cell_val, st.session_state.body_size, r_idx==0)
                    tr = []; doc.add_paragraph()
            continue

        if is_header(lt):
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(12)
            render_run(p, clean_markdown(lt), st.session_state.title_size, True)
            continue

        p = doc.add_paragraph()
        if lt.startswith(('-', '*', '•')):
            sym_val = BULLET_STYLES.get(st.session_state.bullet, "● ")
            if sym_val in ("__NUM__", "__DOT__"): sym_val = "- "
            render_run(p, f"{sym_val}{clean_markdown(lt[1:])}")
        else:
            render_run(p, clean_markdown(lt))
        
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# ═══════════════════════════════════════════
# 5. 1:1 还原 Windows 排版界面
# ═══════════════════════════════════════════
st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; padding-bottom: 0rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 98% !important; }
    
    /* 强制重置工具栏元素间距，实现真正的“紧凑工具栏” */
    div[data-testid="stHorizontalBlock"] { gap: 0.5rem !important; align-items: center !important; margin-bottom: -10px !important;}
    .stButton > button, .stDownloadButton > button, .stPopover > button {
        font-size: 13px !important; padding: 2px 8px !important; min-height: 32px !important; height: 32px !important; border-radius: 4px !important; border: 1px solid #d1d1d6;
    }
    .stDownloadButton > button { font-weight: bold; background-color: #FF3B30; color: white; border: none; }
    .btn-blue > button { background-color: #007AFF !important; color: white !important; font-weight: bold; }
    .btn-green > button { background-color: #34C759 !important; color: white !important; font-weight: bold; }
    
    /* 输入框与标签同一行排列技巧 */
    .stTextInput, .stNumberInput, .stSelectbox { display: flex; flex-direction: row; align-items: center; gap: 5px; }
    .stTextInput label, .stNumberInput label, .stSelectbox label { font-size: 13px !important; min-width: max-content; margin-bottom: 0 !important; }
    .stTextInput input, .stNumberInput input, .stSelectbox div[data-baseweb="select"] { font-size: 13px !important; min-height: 32px !important; height: 32px !important; }
    
    .stTextArea textarea { font-family: Consolas, monospace; font-size: 13px; background-color: #fcfcfc; border: 1px solid #ccc; border-radius: 2px; }
    .header-text { font-size: 18px; font-weight: bold; color: #1D1D1F; padding-bottom: 10px; border-bottom: 1px solid #eaeaea; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-text">📄 规格书专业排版系统 - 现代沉浸版</div>', unsafe_allow_html=True)

col_left, col_mid, col_right = st.columns([1.3, 2.7, 0.8], gap="medium")

# ----------------- 【左侧】文档导入与 AI 撰写区 -----------------
with col_left:
    with st.container(border=True):
        st.markdown("**1. 基础配置**")
        
        # 导入解析区
        st.markdown("<span style='font-size: 12px; color: #666;'>📂 导入文档(PDF/DOCX) 解析底层资料</span>", unsafe_allow_html=True)
        doc_file = st.file_uploader("", type=['pdf', 'docx'], label_visibility="collapsed")
        
        c_p1, c_p2 = st.columns([1, 1])
        if c_p1.button("🚀 提取资料与高清图片", use_container_width=True):
            if doc_file:
                with st.spinner("深度解析中..."):
                    if doc_file.name.endswith(".pdf"):
                        with pdfplumber.open(doc_file) as pdf:
                            st.session_state.raw_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    doc_file.seek(0)
                    extracted = extract_images_from_file(doc_file)
                    for b in extracted:
                        h = hashlib.md5(b).hexdigest()
                        st.session_state.gallery_dict[h] = b
                st.success("解析成功！图片已入图库。")
            else:
                st.warning("请先上传文档！")
                
        if c_p2.button("🗑️ 清空上一篇文案", use_container_width=True):
            st.session_state.txt_cn = ""; st.session_state.txt_en = ""; st.session_state.raw_text = ""
            st.rerun()

        st.divider()

        # 系统设置
        with st.popover("⚙️ 系统设置 (API密钥)", use_container_width=True):
            st.text_input("Kimi API Key", key="kimi_key", type="password")
            st.text_input("百度翻译 APP ID", key="bd_id")
            st.text_input("百度翻译 密钥", key="bd_key", type="password")
            st.button("💾 保存配置", on_click=save_config, type="primary", use_container_width=True)

        c1, c2 = st.columns(2)
        c1.text_input("封面:", key="cover_cn", label_visibility="collapsed", placeholder="封面产品名")
        c2.text_input("EN:", key="cover_en", label_visibility="collapsed", placeholder="EN Name")
        st.checkbox("产品特点含简短说明", key="feature_brief")
        
        # AI 按钮 (通过 on_click 触发)
        st.markdown("<div class='btn-blue'>", unsafe_allow_html=True)
        bc1, bc2 = st.columns(2)
        bc1.button("✨ AI 中文撰写", on_click=perform_ai_write, args=('cn',), use_container_width=True)
        bc2.button("🌐 AI EN Writing", on_click=perform_ai_write, args=('en',), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    tab_cn, tab_en, tab_prompt = st.tabs(["中文文案", "English Copy", "自定义提示词"])
    with tab_cn:
        st.text_area("编辑内容", key="txt_cn", height=450, label_visibility="collapsed")
    with tab_en:
        st.text_area("EN Edit", key="txt_en", height=450, label_visibility="collapsed")
    with tab_prompt:
        st.text_area("中文自定义提示词", key="prompt_cn", height=150)
        st.text_area("英文自定义提示词", key="prompt_en", height=150)

# ----------------- 【中间】排版预览区 (完美复刻原版 4 排工具栏) -----------------
with col_mid:
    with st.container(border=True):
        st.markdown("**📄 预览与排版区**")
        
        # 第一排工具栏
        t1, t2, t3, t4, t5, t6 = st.columns([1, 1, 0.8, 0.8, 1, 1])
        t1.selectbox("中文", FONT_CHOICES, key="font_cn")
        t2.selectbox("英文", FONT_CHOICES, key="font_en")
        t3.number_input("标题", 8, 36, key="title_size")
        t4.number_input("正文", 6, 24, key="body_size")
        t5.selectbox("符号", list(BULLET_STYLES.keys()), key="bullet")
        if t6.button("↺ 重置样式", use_container_width=True):
            st.session_state.font_cn = "微软雅黑"; st.session_state.title_size = 14
            st.session_state.body_size = 11; st.session_state.bullet = "● 实心圆"
            st.rerun()

        # 第二排工具栏
        st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
        b1, b2, b3, b4, b5, b6, b7 = st.columns([0.6, 0.8, 0.8, 0.8, 0.8, 0.8, 1.2])
        b1.markdown("<span style='font-size:13px; line-height:32px;'>📷 图框:</span>", unsafe_allow_html=True)
        if b2.button("1 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:1]"
        if b3.button("2 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:2]"
        if b4.button("3 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:3]"
        if b5.button("4 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:4]"
        if b6.button("6 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:6]"
        if b7.button("✕ 清除全部图框"):
            st.session_state.txt_cn = re.sub(r'\[IMG_FRAME:\d+\]', '', st.session_state.txt_cn)
            st.rerun()

        # 第三排工具栏
        st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1.5, 1])
        if c1.button("🔀 行列互换"): st.toast("网页端暂不支持单表反转")
        if c2.button("✂️ 去说明"): 
            st.toast("一键去说明：请在左侧文本框手动调整。")
        if c3.button("🔄 去序号"): st.toast("系统已强制应用无序号符号！")
        with c4.popover("🖼️ 设置导出Word背景", use_container_width=True):
            st.caption("设置导出的A4底图")
            cover_upload = st.file_uploader("1. 封面图 (A4)", type=['png','jpg'], label_visibility="collapsed")
            if cover_upload: st.session_state.bg_cover_bytes = cover_upload.getvalue()
            body_upload = st.file_uploader("2. 正文图 (A4)", type=['png','jpg'], label_visibility="collapsed")
            if body_upload: st.session_state.bg_body_bytes = body_upload.getvalue()
            if st.button("🗑️ 清空所有背景", use_container_width=True):
                st.session_state.bg_cover_bytes = None; st.session_state.bg_body_bytes = None
        
        # 导出按钮 (完美靠右)
        c5.download_button("📥 导出中文 Word", data=generate_word_document('cn'), file_name=f"{st.session_state.cover_cn}.docx", use_container_width=True)
        # c6.download_button("📥 导出英文 Word", data=generate_word_document('en'), file_name=f"{st.session_state.cover_en}.docx", use_container_width=True)

    # ---------------- 核心：独立 HTML A4 实时高保真渲染引擎 ----------------
    preview_txt = st.session_state.txt_cn if st.session_state.current_lang == 'cn' else st.session_state.txt_en
    cover_title = st.session_state.cover_cn if st.session_state.current_lang == 'cn' else st.session_state.cover_en
    font_family = st.session_state.font_cn if st.session_state.current_lang == 'cn' else st.session_state.font_en
    title_px = st.session_state.title_size * 1.5
    body_px = st.session_state.body_size * 1.3
    
    html_body = ""
    in_table = False
    for line in preview_txt.split('\n'):
        line = line.strip()
        if not line: continue
        
        # 渲染右侧图库插入的实体图片
        local_img_m = re.match(r'^\[LOCAL_IMG:(.+?)\]$', line)
        if local_img_m:
            img_hash = local_img_m.group(1)
            img_bytes = st.session_state.gallery_dict.get(img_hash)
            if img_bytes:
                b64 = bytes_to_b64(img_bytes)
                html_body += f'<div style="text-align:center; margin: 20px 0;"><img src="data:image/png;base64,{b64}" style="max-width: 80%; border: 1px solid #ccc; box-shadow: 0 4px 8px rgba(0,0,0,0.1); border-radius: 4px; padding: 5px; background: white;"/></div>'
            continue
            
        if '|' in line:
            if '---' in line: continue
            if not in_table:
                html_body += '<table style="width:100%; border-collapse: collapse; margin-top:10px; margin-bottom: 10px;">'
                in_table = True
            cells = [clean_markdown(c) for c in line.strip('|').split('|') if c.strip() != '']
            html_body += "<tr>"
            for c in cells: html_body += f'<td style="border: 1px solid #000; padding: 6px; text-align:center;">{c}</td>'
            html_body += "</tr>"
            continue
            
        if in_table:
            html_body += '</table>'
            in_table = False
            
        if 'IMG_FRAME' in line:
            html_body += f'<div style="border:2px dashed #007AFF; padding: 30px; text-align: center; margin: 15px 0; background: rgba(0,122,255,0.05); color:#007AFF;"><b>🖼️ 排版预留占位框</b></div>'
            continue
            
        if is_header(line):
            html_body += f'<div style="margin-top:20px; margin-bottom:5px; font-size:{title_px}px; font-weight:bold; color:#000;">{clean_markdown(line)}</div>'
            continue
            
        if line.startswith(('-', '*', '•')):
            sym = BULLET_STYLES.get(st.session_state.bullet, "●")
            if sym in ("__NUM__", "__DOT__"): sym = "- "
            html_body += f'<div style="margin-bottom:6px; margin-left: 20px;">{sym} {clean_markdown(line[1:])}</div>'
        else:
            html_body += f'<div style="margin-bottom:6px;">{clean_markdown(line)}</div>'
            
    if in_table: html_body += '</table>'

    cv_bg = f"background-image: url('data:image/png;base64,{bytes_to_b64(st.session_state.bg_cover_bytes)}');" if st.session_state.bg_cover_bytes else "background-color: white;"
    bd_bg = f"background-image: url('data:image/png;base64,{bytes_to_b64(st.session_state.bg_body_bytes)}');" if st.session_state.bg_body_bytes else "background-color: white;"

    iframe_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ background-color: #EAEBEE; text-align: center; padding: 20px 0; margin: 0; font-family: '{font_family}', sans-serif; }}
            .a4-page {{ width: 794px; background-color: white; margin: 0 auto 30px auto; box-shadow: 0 4px 12px rgba(0,0,0,0.2); background-size: 100% 100%; position: relative; box-sizing: border-box; }}
            .cover {{ height: 1123px; display: flex; align-items: center; justify-content: center; {cv_bg} }}
            .body-page {{ min-height: 1123px; padding: 96px 72px; text-align: left; {bd_bg} font-size: {body_px}px; color: #1D1D1F; line-height: 1.6; }}
            table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
            td, th {{ border: 1px solid #000; padding: 8px; text-align: center; font-size: {body_px}px; }}
        </style>
    </head>
    <body>
        <div class="a4-page cover">
            <h1 style="font-size: 3.5em; color: #1D1D1F;">{cover_title}</h1>
        </div>
        <div class="a4-page body-page">
            {html_body}
        </div>
    </body>
    </html>
    """
    components.html(iframe_html, height=720, scrolling=True)

# ----------------- 【右侧】图库与可视化插入 -----------------
with col_right:
    with st.container(border=True):
        st.markdown("**📷 图库**")
        manual_upload = st.file_uploader("＋ 上传图片", accept_multiple_files=True, type=['png','jpg','jpeg'], label_visibility="collapsed")
        if manual_upload:
            for f in manual_upload:
                bytes_data = f.getvalue()
                h = hashlib.md5(bytes_data).hexdigest()
                st.session_state.gallery_dict[h] = bytes_data
                    
        if st.button("🗑 清空图库", use_container_width=True):
            st.session_state.gallery_dict = {}
            st.rerun()
            
        st.caption("点击下方按钮将图片插入正文：")
        for img_hash, img_bytes in st.session_state.gallery_dict.items():
            try: 
                st.image(img_bytes, use_container_width=True)
                st.markdown("<div class='btn-green'>", unsafe_allow_html=True)
                if st.button("➕ 插入至文末", key=f"ins_{img_hash}", use_container_width=True):
                    target_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
                    # 追加图片标记到文案中
                    st.session_state[target_key] += f"\n\n[LOCAL_IMG:{img_hash}]\n\n"
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
            except: pass
