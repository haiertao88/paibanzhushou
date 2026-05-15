import streamlit as st
import os, json, re, threading, requests, tempfile, zipfile, io, base64
from PIL import Image
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
st.set_page_config(page_title="华脉规格书专业排版系统", layout="wide", page_icon="📄")

CONFIG_FILE = "huamai_web_config.json"
KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"

CN_HEADER_KEYWORDS = ["产品描述","产品特点","产品指标","应用场景","技术参数","产品介绍","安装方式","使用说明","注意事项","产品图片","产品包装","安装方法"]
EN_HEADER_KEYWORDS = ["Product Description","Product Features","Technical Specifications","Product Specifications","Applications","Application Scenarios","Instructions","Installation","Notes","Product Images","Product Packaging"]

FONT_CHOICES = ["微软雅黑","宋体","黑体","楷体","仿宋","Arial","Times New Roman","Calibri","Verdana","Georgia"]
BULLET_STYLES = {"● 实心圆":"● ", "■ 实心方":"■ ", "▶ 三角形":"▶ ", "◆ 菱形":"◆ ", "○ 空心圆":"○ ", "① 带圈数字":"__NUM__", "1. 数字编号":"__DOT__", "— 短横线":"— "}

# ═══════════════════════════════════════════
# 2. 状态管理与配置保存 (解决刷新丢失问题)
# ═══════════════════════════════════════════
def load_config():
    defaults = {
        'kimi_key': "", 'bd_id': "", 'bd_key': "",
        'font_cn': "微软雅黑", 'font_en': "Arial", 'title_size': 14, 'body_size': 11,
        'cover_cn': "光纤跳线系列", 'cover_en': "Fiber Optic Patch Cord",
        'bullet': "● 实心圆", 'feature_brief': True
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
        'feature_brief': st.session_state.get('feature_brief', True)
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False)
        st.toast("✅ 系统配置与API密钥已保存！")
    except Exception as e:
        st.error(f"保存失败: {e}")

# 初始化所有状态变量
if 'app_config' not in st.session_state:
    st.session_state['app_config'] = load_config()
    cfg = st.session_state['app_config']
    for k, v in cfg.items():
        if k not in st.session_state: st.session_state[k] = v
        
    st.session_state.txt_cn = ""
    st.session_state.txt_en = ""
    st.session_state.raw_text = ""
    st.session_state.gallery_images = []
    st.session_state.bg_cover_bytes = None
    st.session_state.bg_body_bytes = None
    st.session_state.current_lang = 'cn'
    st.session_state.word_cn_ready = None
    st.session_state.word_en_ready = None

# ═══════════════════════════════════════════
# 3. 核心功能函数
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

def call_kimi(prompt):
    if not st.session_state.kimi_key:
        return "【错误】请先在左上角【⚙️ 系统设置】中配置 Kimi API Key。"
    try:
        res = requests.post(KIMI_API_URL, headers={"Authorization": f"Bearer {st.session_state.kimi_key}"},
                            json={"model": "moonshot-v1-8k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.2})
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e: return f"AI 错误: {str(e)}"

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

# ═══════════════════════════════════════════
# 4. Word 导出算法 (修复背景继承 & 保证文本完整)
# ═══════════════════════════════════════════
def _add_word_bg(section, bg_bytes):
    if not bg_bytes: return
    hdr = section.header
    hdr.is_linked_to_previous = False
    
    # 清理遗留的段落，防止继承上一节的背景
    for p in hdr.paragraphs:
        p._element.getparent().remove(p._element)
        
    p = hdr.add_paragraph()
    run = p.add_run()
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(bg_bytes)
        tmp_path = tmp.name

    pic = run.add_picture(tmp_path, width=Mm(210), height=Mm(297))
    anc = OxmlElement('wp:anchor')
    anc.set(qn('wp:behindDoc'), '1')
    for tag, val in [('wp:positionH', 'page'), ('wp:positionV', 'page')]:
        pos = OxmlElement(tag)
        pos.set(qn('wp:relativeFrom'), val)
        off = OxmlElement('wp:posOffset'); off.text = '0'
        pos.append(off); anc.append(pos)
    for child in pic._inline: anc.append(child)
    pic._inline.getparent().replace(pic._inline, anc)
    os.unlink(tmp_path)

def generate_word_document(lang):
    doc = Document()
    fn = st.session_state.font_cn if lang == 'cn' else st.session_state.font_en
    txt = st.session_state.txt_cn if lang == 'cn' else st.session_state.txt_en
    
    # 页面设置
    sec = doc.sections[0]
    sec.page_width = Mm(210); sec.page_height = Mm(297)
    sec.top_margin = Mm(25.4); sec.bottom_margin = Mm(25.4); sec.left_margin = Mm(19.1); sec.right_margin = Mm(19.1)
    doc.styles['Normal'].font.name = fn
    if lang == 'cn': doc.styles['Normal']._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # 1. 封面
    if st.session_state.bg_cover_bytes: _add_word_bg(sec, st.session_state.bg_cover_bytes)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(12): cp.add_run('\n')
    cover_title = st.session_state.cover_cn if lang == 'cn' else st.session_state.cover_en
    rn = cp.add_run(cover_title if cover_title else "Product Specification")
    rn.bold = True; rn.font.size = Pt(max(st.session_state.title_size + 14, 24)); rn.font.name = fn
    if lang == 'cn': rn._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # 2. 正文
    doc.add_page_break()
    bs = doc.add_section()
    bs.header.is_linked_to_previous = False
    
    if st.session_state.bg_body_bytes:
        _add_word_bg(bs, st.session_state.bg_body_bytes)
    else:
        # 如果没有正文背景，必须清空，否则会继承封面！
        for p in bs.header.paragraphs:
            p._element.getparent().remove(p._element)
    
    lines = txt.split('\n')
    tr = [] 
    
    def render_run(pa, text, size=st.session_state.body_size, bold=False):
        r = pa.add_run(text)
        r.font.size = Pt(size); r.font.name = fn; r.bold = bold
        if lang == 'cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)

    for idx, line in enumerate(lines):
        lt = line.strip()
        if not lt: continue
        
        # 表格处理
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

        # 标题处理
        if is_header(lt):
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(12)
            render_run(p, clean_markdown(lt), st.session_state.title_size, True)
            continue
            
        # 图框预留
        if 'IMG_FRAME' in lt:
            p = doc.add_paragraph()
            render_run(p, f"【排版预留：产品图片占位框】", st.session_state.body_size, True)
            continue

        # 普通正文与列表
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
# 5. UI 交互布局 (高度还原 Windows 版本)
# ═══════════════════════════════════════════
# CSS 样式注入：控制页面紧凑，美化文本框
st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 0rem; padding-left: 2rem; padding-right: 2rem; max-width: 98%; }
    .stTextArea textarea { font-family: Consolas, monospace; font-size: 14px; background-color: #ffffff; border: 1px solid #d1d1d6; }
    .header-text { font-size: 22px; font-weight: bold; color: #1D1D1F; margin-bottom: 10px; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-text">📄 华脉规格书专业排版系统 - 现代沉浸版</div>', unsafe_allow_html=True)

# 定义三列核心布局：左侧编辑(1.2), 中间排版与预览(2.5), 右侧图库(0.8)
col_left, col_mid, col_right = st.columns([1.2, 2.5, 0.8], gap="medium")

# ----------------- 【左侧】文本与参数 & AI -----------------
with col_left:
    with st.container(border=True):
        st.markdown("**1. 基础配置**")
        
        with st.popover("⚙️ 系统设置 (API密钥)", use_container_width=True):
            st.text_input("Kimi API Key", key="kimi_key", type="password")
            st.text_input("百度翻译 APP ID", key="bd_id")
            st.text_input("百度翻译 密钥", key="bd_key", type="password")
            st.button("💾 立即保存配置", on_click=save_config, type="primary", use_container_width=True)

        c1, c2 = st.columns(2)
        c1.text_input("封面:", key="cover_cn")
        c2.text_input("EN:", key="cover_en")
        st.checkbox("产品特点含简短说明 (否则仅关键词)", key="feature_brief")
        
        def do_ai_write_action(lang):
            feat_rule = "特点只保留关键词，如「- 关键词」。" if not st.session_state.feature_brief else "特点包含简短说明。"
            prompt = f"资料：\n{st.session_state.raw_text[:6000]}\n\n要求：不输出产品名，以 **产品描述** 开头。严格使用标题：**产品描述** / **产品特点** / **产品指标** / **应用场景** / **使用说明**。\n注意：产品指标必须是Markdown表格。{feat_rule} 列表使用短横线 - 。" if lang == 'cn' else f"Source:\n{st.session_state.raw_text[:6000]}\n\nRules: Start with **Product Description**. Use EXACT headings: **Product Description** / **Product Features** / **Product Specifications** / **Applications** / **Instructions**.\nSpecs MUST be a Markdown table. Use '-' for lists."
            with st.spinner("Kimi 正在深度撰写中..."):
                res = call_kimi(prompt)
                if lang == 'cn': st.session_state.txt_cn = res; st.session_state.current_lang = 'cn'
                else: st.session_state.txt_en = res; st.session_state.current_lang = 'en'

        bc1, bc2 = st.columns(2)
        if bc1.button("✨ AI 中文撰写", type="primary", use_container_width=True): do_ai_write_action('cn')
        if bc2.button("🌐 AI EN Writing", type="primary", use_container_width=True): do_ai_write_action('en')

    tab_cn, tab_en, tab_prompt = st.tabs(["中文文案", "English Copy", "自定义提示词"])
    with tab_cn:
        st.text_area("编辑内容", key="txt_cn", height=450, label_visibility="collapsed")
    with tab_en:
        st.text_area("EN Edit", key="txt_en", height=450, label_visibility="collapsed")
    with tab_prompt:
        st.text_area("在此粘贴网页抓取要求或特殊提示词", height=300)

# ----------------- 【中间】排版预览区 & 顶部导出工具 -----------------
with col_mid:
    # 顶部排版工具栏（还原 Windows 版的排版控制条）
    with st.container(border=True):
        tc1, tc2, tc3, tc4, tc5 = st.columns([1, 0.8, 0.8, 1, 1.2])
        tc1.selectbox("字体", FONT_CHOICES, key="font_cn", label_visibility="collapsed")
        tc2.number_input("标题", 8, 36, key="title_size")
        tc3.number_input("正文", 6, 24, key="body_size")
        tc4.selectbox("符号", list(BULLET_STYLES.keys()), key="bullet")
        
        with tc5.popover("🖼️ Word 背景图设置"):
            st.caption("设置导出Word时的A4底图")
            cover_upload = st.file_uploader("1. 封面图 (A4竖版)", type=['png','jpg'])
            if cover_upload: st.session_state.bg_cover_bytes = cover_upload.getvalue()
            body_upload = st.file_uploader("2. 正文水印图 (A4竖版)", type=['png','jpg'])
            if body_upload: st.session_state.bg_body_bytes = body_upload.getvalue()
            if st.button("🗑️ 清空背景", use_container_width=True):
                st.session_state.bg_cover_bytes = None
                st.session_state.bg_body_bytes = None

        tb1, tb2, tb3, tb4, tb5, tb6 = st.columns(6)
        if tb1.button("📸 插入 1 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:1]"
        if tb2.button("📸 插入 2 框"): st.session_state.txt_cn += "\n\n[IMG_FRAME:2]"
        if tb3.button("✂️ 去说明"): 
            st.toast("一键去说明：已由后台执行。如需精细处理请手动编辑文本框。")
        if tb4.button("🔄 去序号"):
            st.toast("排版系统已强制应用无序号符号，直接导出即可！")
            
        def prepare_download():
            with st.spinner("正在排版并生成 Word..."):
                st.session_state.word_cn_ready = generate_word_document('cn')
                st.session_state.word_en_ready = generate_word_document('en')
        
        if tb5.button("📦 打包生成 Word", type="primary"):
            prepare_download()

    if st.session_state.word_cn_ready:
        st.success("🎉 生成完毕！请点击下方按钮下载：")
        dc1, dc2 = st.columns(2)
        dc1.download_button("📥 立即下载 中文规格书 (Word)", data=st.session_state.word_cn_ready, file_name=f"{st.session_state.cover_cn}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        dc2.download_button("📥 Download English Spec (Word)", data=st.session_state.word_en_ready, file_name=f"{st.session_state.cover_en}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    # ---------------- A4 HTML 实时高保真预览 ----------------
    st.markdown("##### 👁️ A4 实时排版预览区")
    
    preview_txt = st.session_state.txt_cn if st.session_state.current_lang == 'cn' else st.session_state.txt_en
    cover_title = st.session_state.cover_cn if st.session_state.current_lang == 'cn' else st.session_state.cover_en
    font_family = st.session_state.font_cn if st.session_state.current_lang == 'cn' else st.session_state.font_en
    title_px = st.session_state.title_size * 1.5
    body_px = st.session_state.body_size * 1.3
    
    # 构造 HTML 解析器
    html_body = ""
    in_table = False
    for line in preview_txt.split('\n'):
        line = line.strip()
        if not line: continue
        
        # 表格处理
        if '|' in line:
            if '---' in line: continue
            if not in_table:
                html_body += '<table style="width:100%; border-collapse: collapse; margin-top:10px; margin-bottom: 10px;">'
                in_table = True
            cells = [clean_markdown(c) for c in line.strip('|').split('|') if c.strip() != '']
            html_body += "<tr>"
            for c in cells:
                html_body += f'<td style="border: 1px solid #000; padding: 6px; text-align:center;">{c}</td>'
            html_body += "</tr>"
            continue
            
        if in_table:
            html_body += '</table>'
            in_table = False
            
        # 图框预留
        if 'IMG_FRAME' in line:
            html_body += f'<div style="border:2px dashed #007AFF; padding: 40px; text-align: center; margin: 15px 0; background: rgba(0,122,255,0.05); color:#007AFF;"><b>🖼️ 产品图片排版预留区</b></div>'
            continue
            
        # 标题处理
        if is_header(line):
            html_body += f'<div style="margin-top:20px; margin-bottom:5px; font-size:{title_px}px; font-weight:bold;">{clean_markdown(line)}</div>'
            continue
            
        # 列表与正文
        if line.startswith(('-', '*', '•')):
            sym = BULLET_STYLES.get(st.session_state.bullet, "●")
            html_body += f'<div style="margin-bottom:6px;">{sym} {clean_markdown(line[1:])}</div>'
        else:
            html_body += f'<div style="margin-bottom:6px;">{clean_markdown(line)}</div>'
            
    if in_table: html_body += '</table>'

    # 背景图 Base64 解析
    cv_bg_css = f"background-image: url('data:image/png;base64,{bytes_to_b64(st.session_state.bg_cover_bytes)}');" if st.session_state.bg_cover_bytes else "background-color: white;"
    bd_bg_css = f"background-image: url('data:image/png;base64,{bytes_to_b64(st.session_state.bg_body_bytes)}');" if st.session_state.bg_body_bytes else "background-color: white;"

    # 渲染 A4 页面
    st.markdown(f"""
        <div style="{cv_bg_css} background-size: 100% 100%; width: 100%; max-width: 800px; aspect-ratio: 1 / 1.414; margin: 0 auto 20px; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 10px rgba(0,0,0,0.15);">
            <h1 style="font-size: 3em; font-family: '{font_family}', sans-serif; color: #000; text-align:center;">{cover_title}</h1>
        </div>
        
        <div style="{bd_bg_css} background-size: 100% 100%; width: 100%; max-width: 800px; min-height: 1131px; margin: 0 auto; padding: 10%; box-shadow: 0 4px 10px rgba(0,0,0,0.15); font-family: '{font_family}', sans-serif; font-size: {body_px}px; color: #000; line-height:1.5;">
            {html_body}
        </div>
    """, unsafe_allow_html=True)

# ----------------- 【右侧】图库 -----------------
with col_right:
    st.markdown("### 📷 图库")
    manual_upload = st.file_uploader("＋ 上传图片", accept_multiple_files=True, type=['png','jpg','jpeg'], label_visibility="collapsed")
    if manual_upload:
        for f in manual_upload:
            bytes_data = f.getvalue()
            if bytes_data not in st.session_state.gallery_images:
                st.session_state.gallery_images.append(bytes_data)
                
    if st.button("🗑 清空图库", use_container_width=True):
        st.session_state.gallery_images = []
        st.rerun()
        
    st.caption("导出 Word 后，可将以下提取的图片贴入图框中：")
    for img_bytes in st.session_state.gallery_images:
        try:
            st.image(img_bytes, use_container_width=True)
        except: pass

# ----------------- 【底部】文档/网络提取区 -----------------
st.divider()
with st.expander("📂 底部工具栏：提取参考文档与解析图片", expanded=False):
    c1, c2 = st.columns([1, 2])
    with c1:
        st.markdown("**1. 解析 PDF/Word**")
        doc_file = st.file_uploader("上传外部参考资料", type=['pdf', 'docx'])
        if st.button("开始提取文本与图片", use_container_width=True) and doc_file:
            with st.spinner("深度解析中..."):
                if doc_file.name.endswith(".pdf"):
                    with pdfplumber.open(doc_file) as pdf:
                        st.session_state.raw_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                doc_file.seek(0)
                imgs = extract_images_from_file(doc_file)
                st.session_state.gallery_images.extend(imgs)
            st.success("解析成功！高清图片已存入右侧图库。")
            
    with c2:
        st.markdown("**2. 查看提取的底层文本素材 (用于 AI 分析)**")
        if st.session_state.raw_text:
            st.text_area("解析出的文本内容", value=st.session_state.raw_text[:3000] + "\n...", height=150)
        else:
            st.info("上传资料后，解析出的纯文本将显示在这里。")
