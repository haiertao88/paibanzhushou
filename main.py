import streamlit as st
import os, json, re, threading, requests, tempfile, zipfile, io
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
BULLET_STYLES = {"● 实心圆":"● ","■ 实心方":"■ ","▶ 三角形":"▶ ","◆ 菱形":"◆ ","○ 空心圆":"○ ","① 带圈数字":"__NUM__","1. 数字编号":"__DOT__","— 短横线":"— "}

# ═══════════════════════════════════════════
# 2. 配置存取逻辑 (解决刷新丢失配置的问题)
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
                saved = json.load(f)
                defaults.update(saved)
        except: pass
    return defaults

def save_config(current_state):
    to_save = {
        'kimi_key': current_state.get('kimi_key', ''),
        'bd_id': current_state.get('bd_id', ''),
        'bd_key': current_state.get('bd_key', ''),
        'font_cn': current_state.get('font_cn', '微软雅黑'),
        'font_en': current_state.get('font_en', 'Arial'),
        'title_size': current_state.get('title_size', 14),
        'body_size': current_state.get('body_size', 11),
        'cover_cn': current_state.get('cover_cn', '光纤跳线系列'),
        'cover_en': current_state.get('cover_en', 'Fiber Optic Patch Cord'),
        'bullet': current_state.get('bullet', '● 实心圆'),
        'feature_brief': current_state.get('feature_brief', True)
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, ensure_ascii=False)
        return True
    except:
        return False

# 初始化状态
if 'app_config' not in st.session_state:
    st.session_state['app_config'] = load_config()
    st.session_state['txt_cn'] = ""
    st.session_state['txt_en'] = ""
    st.session_state['raw_text'] = ""
    st.session_state['extracted_images'] = []
    st.session_state['current_lang'] = 'cn'

# ═══════════════════════════════════════════
# 3. 核心工具函数
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
    api_key = st.session_state['app_config']['kimi_key']
    if not api_key: return "【错误】请先在左侧栏输入 Kimi API Key 并保存配置。"
    try:
        res = requests.post(KIMI_API_URL, headers={"Authorization": f"Bearer {api_key}"},
                            json={"model": "moonshot-v1-8k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.2})
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e: return f"AI 错误: {str(e)}"

def do_ai_write(lang):
    cfg = st.session_state['app_config']
    feat_rule = "特点只保留关键词，如「- 关键词」。" if not cfg['feature_brief'] else "特点包含简短说明。"
    if lang == 'cn':
        prompt = f"资料：\n{st.session_state['raw_text'][:6000]}\n\n要求：不输出产品名，以 **产品描述** 开头。严格使用标题：**产品描述** / **产品特点** / **产品指标** / **应用场景** / **使用说明**。\n注意：产品指标必须是Markdown表格。{feat_rule} 列表使用短横线 - 。"
    else:
        prompt = f"Source:\n{st.session_state['raw_text'][:6000]}\n\nRules: Start with **Product Description**. Use EXACT headings: **Product Description** / **Product Features** / **Product Specifications** / **Applications** / **Instructions**.\nSpecs MUST be a Markdown table. Use '-' for lists."
    
    with st.spinner("AI 智能撰写中，请稍候..."):
        res = call_kimi(prompt)
        if lang == 'cn': st.session_state['txt_cn'] = res
        else: st.session_state['txt_en'] = res

def extract_images_from_file(uploaded_file):
    td = tempfile.mkdtemp()
    extracted = []
    if uploaded_file.name.endswith(".pdf"):
        import fitz
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for i in range(len(doc)):
            for img in doc[i].get_images(full=True):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                if len(image_bytes) > 5000:
                    img_path = os.path.join(td, f"page{i}_{xref}.{base_image['ext']}")
                    with open(img_path, "wb") as f: f.write(image_bytes)
                    extracted.append(img_path)
    elif uploaded_file.name.endswith(".docx"):
        with zipfile.ZipFile(io.BytesIO(uploaded_file.read())) as z:
            for f in z.namelist():
                if f.startswith("word/media/") and f.lower().endswith(('.png','.jpg','.jpeg')):
                    img_path = os.path.join(td, os.path.basename(f))
                    with open(img_path, "wb") as out: out.write(z.read(f))
                    extracted.append(img_path)
    return extracted

# ═══════════════════════════════════════════
# 4. 完整的 Word 生成算法 (带背景图注入)
# ═══════════════════════════════════════════
def _add_word_bg(section, uploaded_bg_file):
    hdr = section.header
    hdr.is_linked_to_previous = False
    run = hdr.paragraphs[0].add_run()
    
    # 将上传的文件保存为临时文件供 docx 读取
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(uploaded_bg_file.getvalue())
        tmp_path = tmp.name

    pic = run.add_picture(tmp_path, width=Mm(210), height=Mm(297))
    anc = OxmlElement('wp:anchor')
    anc.set(qn('wp:behindDoc'), '1')
    for tag, val in [('wp:positionH', 'page'), ('wp:positionV', 'page')]:
        pos = OxmlElement(tag)
        pos.set(qn('wp:relativeFrom'), val)
        off = OxmlElement('wp:posOffset')
        off.text = '0'
        pos.append(off)
        anc.append(pos)
    for child in pic._inline:
        anc.append(child)
    pic._inline.getparent().replace(pic._inline, anc)
    os.unlink(tmp_path)

def generate_word_document(lang, cover_bg_file=None, body_bg_file=None):
    cfg = st.session_state['app_config']
    doc = Document()
    fn = cfg['font_cn'] if lang == 'cn' else cfg['font_en']
    txt = st.session_state['txt_cn'] if lang == 'cn' else st.session_state['txt_en']
    
    # 页面设置 A4
    sec = doc.sections[0]
    sec.page_width = Mm(210); sec.page_height = Mm(297)
    sec.top_margin = Mm(25.4); sec.bottom_margin = Mm(25.4); sec.left_margin = Mm(19.1); sec.right_margin = Mm(19.1)
    doc.styles['Normal'].font.name = fn
    if lang == 'cn': doc.styles['Normal']._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # --- 1. 生成封面 ---
    if cover_bg_file: _add_word_bg(sec, cover_bg_file)
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(12): cp.add_run('\n')
    cover_title = cfg['cover_cn'] if lang == 'cn' else cfg['cover_en']
    rn = cp.add_run(cover_title if cover_title else "Product Specification")
    rn.bold = True; rn.font.size = Pt(max(cfg['title_size'] + 14, 24)); rn.font.name = fn
    if lang == 'cn': rn._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # --- 2. 生成正文 ---
    doc.add_page_break()
    bs = doc.add_section()
    if body_bg_file: _add_word_bg(bs, body_bg_file)
    
    lines = txt.split('\n')
    tr = [] # 表格缓冲
    
    def render_run(pa, text, size=cfg['body_size'], bold=False):
        r = pa.add_run(text)
        r.font.size = Pt(size); r.font.name = fn; r.bold = bold
        if lang == 'cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)

    for idx, line in enumerate(lines):
        lt = line.strip()
        if not lt: continue
        
        # 处理表格
        if '|' in lt:
            if '---' in lt: continue
            cells = [clean_markdown(c) for c in lt.strip('|').split('|') if c.strip() != '']
            if cells: tr.append(cells)
            
            # 判断表格是否结束
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
                                r = p.add_run(cell_val); r.font.name = fn; r.font.size = Pt(cfg['body_size'])
                                if lang == 'cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
                                if r_idx == 0: r.bold = True
                    tr = []; doc.add_paragraph()
            continue

        # 处理标题
        if is_header(lt):
            p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(12)
            render_run(p, clean_markdown(lt), cfg['title_size'], True)
            continue
            
        # 处理列表
        if lt.startswith(('-', '*', '•')):
            p = doc.add_paragraph()
            # 这里统一用黑圆点占位，保证网页导出格式干净
            render_run(p, f"● {clean_markdown(lt[1:])}")
            continue
            
        # 处理图框预留
        if 'IMG_FRAME' in lt:
            p = doc.add_paragraph()
            render_run(p, f"【请在此处插入产品图片】", cfg['body_size'], True)
            continue

        # 普通文本
        p = doc.add_paragraph()
        render_run(p, clean_markdown(lt))
        
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# ═══════════════════════════════════════════
# 5. UI 交互布局
# ═══════════════════════════════════════════
def main():
    cfg = st.session_state['app_config']
    
    # ---------------- 侧边栏：系统与参数设置 ----------------
    with st.sidebar:
        st.title("⚙️ 华脉系统设置")
        
        with st.expander("🔑 接口与密钥", expanded=True):
            cfg['kimi_key'] = st.text_input("Kimi API Key", value=cfg['kimi_key'], type="password")
            cfg['bd_id'] = st.text_input("百度翻译 ID", value=cfg['bd_id'])
            cfg['bd_key'] = st.text_input("百度翻译 Key", value=cfg['bd_key'], type="password")
        
        with st.expander("📐 排版与字体", expanded=False):
            cfg['font_cn'] = st.selectbox("中文字体", FONT_CHOICES, index=FONT_CHOICES.index(cfg['font_cn']) if cfg['font_cn'] in FONT_CHOICES else 0)
            cfg['title_size'] = st.number_input("标题字号", 8, 36, cfg['title_size'])
            cfg['body_size'] = st.number_input("正文字号", 6, 24, cfg['body_size'])
            cfg['cover_cn'] = st.text_input("中文封面名称", cfg['cover_cn'])
            cfg['cover_en'] = st.text_input("英文封面名称", cfg['cover_en'])
            cfg['feature_brief'] = st.checkbox("产品特点含简短说明", cfg['feature_brief'])
            
        with st.expander("🖼️ Word 背景图设置", expanded=True):
            st.info("导出 Word 时，以下图片将自动贴入底层作为背景")
            cover_bg = st.file_uploader("1. 封面背景图 (A4竖版)", type=['png', 'jpg', 'jpeg'])
            body_bg = st.file_uploader("2. 正文水印图 (A4竖版)", type=['png', 'jpg', 'jpeg'])

        if st.button("💾 保存所有配置", use_container_width=True):
            st.session_state['app_config'] = cfg
            if save_config(cfg): st.success("配置已持久保存！")
            else: st.error("保存失败，请检查读写权限。")

    # ---------------- 主界面区 ----------------
    st.markdown('<div style="font-size: 26px; font-weight: bold; color: #007AFF;">📄 华脉规格书专业排版系统 - 网页增强版</div>', unsafe_allow_html=True)
    
    # 顶部工具栏
    t1, t2, t3, t4 = st.columns([1, 1, 1.5, 1.5])
    with t1:
        if st.button("✂️ 一键去冒号说明", use_container_width=True):
            st.toast("提示：请在下方文本框中手动编辑或删除多余说明。")
    with t2:
        if st.button("🔀 一键去序号", use_container_width=True):
            st.toast("提示：自动排版已禁用序号，均使用符号替代。")
    with t3:
        st.write("📥 **中文导出**")
        word_cn = generate_word_document('cn', cover_bg, body_bg)
        st.download_button("📄 下载 中文版规格书 (Word)", data=word_cn, file_name=f"{cfg['cover_cn']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with t4:
        st.write("📥 **英文导出**")
        word_en = generate_word_document('en', cover_bg, body_bg)
        st.download_button("📄 下载 English Spec (Word)", data=word_en, file_name=f"{cfg['cover_en']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    st.divider()

    # ---------------- 功能页签 ----------------
    tab1, tab2, tab3 = st.tabs(["📝 文案创作与编辑", "📂 素材解析与图库", "👁️ 简易预览"])
    
    with tab1:
        col_cn, col_en = st.columns(2)
        with col_cn:
            st.subheader("🇨🇳 中文排版区")
            if st.button("✨ 启动 Kimi 撰写 (中文)", use_container_width=True): do_ai_write('cn')
            st.session_state['txt_cn'] = st.text_area("中文 Markdown 代码", value=st.session_state['txt_cn'], height=500, label_visibility="collapsed")
            if st.session_state['txt_cn']: st.session_state['current_lang'] = 'cn'
            
        with col_en:
            st.subheader("🇬🇧 English Copy")
            if st.button("✨ Start AI Writing (EN)", use_container_width=True): do_ai_write('en')
            st.session_state['txt_en'] = st.text_area("English Markdown Code", value=st.session_state['txt_en'], height=500, label_visibility="collapsed")
            if st.session_state['txt_en'] and not st.session_state['txt_cn']: st.session_state['current_lang'] = 'en'

    with tab2:
        col_doc, col_gallery = st.columns([1, 2])
        with col_doc:
            st.subheader("1. 导入资料并提取文本/图片")
            doc_file = st.file_uploader("上传 PDF 或 DOCX 文件", type=['pdf', 'docx'])
            if st.button("解析文本与图片", use_container_width=True) and doc_file:
                with st.spinner("深度解析中..."):
                    # 提文本
                    if doc_file.name.endswith(".pdf"):
                        with pdfplumber.open(doc_file) as pdf:
                            st.session_state['raw_text'] = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    # 提图片
                    doc_file.seek(0)
                    imgs = extract_images_from_file(doc_file)
                    st.session_state['extracted_images'].extend(imgs)
                st.success("解析成功！已更新参考素材及图库。")
                
            with st.expander("查看当前识别的文本素材"):
                st.text(st.session_state['raw_text'][:2000] + "\n...")

        with col_gallery:
            st.subheader("2. 提取图片画廊")
            st.info("从文档中提取的高清图片列表（导出 Word 后，可将它们贴入图框位置）")
            if st.session_state['extracted_images']:
                cols = st.columns(3)
                for i, img_path in enumerate(st.session_state['extracted_images']):
                    try:
                        with cols[i % 3]:
                            st.image(img_path, use_container_width=True, caption=os.path.basename(img_path))
                    except: pass
            else:
                st.write("暂无提取的图片。")

    with tab3:
        st.subheader("网页实时内容预览")
        st.caption("提示：此处仅预览文本内容结构。准确的A4排版、字体、以及您上传的封面背景/水印，请点击上方『下载 Word』查看。")
        preview_txt = st.session_state['txt_cn'] if st.session_state['current_lang'] == 'cn' else st.session_state['txt_en']
        
        with st.container(border=True):
            cover = cfg['cover_cn'] if st.session_state['current_lang'] == 'cn' else cfg['cover_en']
            st.markdown(f"# {cover}")
            st.divider()
            
            # 简化网页端预览展示
            html_txt = preview_txt.replace("[IMG_FRAME:1]", "🖼️ **[此处预留了 1 张产品大图的位置]**")
            html_txt = html_txt.replace("[IMG_FRAME:2]", "🖼️ **[此处预留了 2 张产品小图的位置]**")
            st.markdown(html_txt)

if __name__ == "__main__":
    main()
