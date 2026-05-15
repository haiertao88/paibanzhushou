import streamlit as st
import os, json, re, threading, requests, tempfile, copy, time, random, hashlib, zipfile, io
from urllib.parse import urlparse, quote
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import pdfplumber

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ═══════════════════════════════════════════
# 常量与配置
# ═══════════════════════════════════════════
st.set_page_config(page_title="华脉专业规格书智能排版系统", layout="wide", page_icon="📄")

KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"
CN_HEADER_KEYWORDS = ["产品描述","产品特点","产品指标","应用场景","技术参数","产品介绍","安装方式","使用说明","注意事项","产品图片","产品包装","安装方法"]
EN_HEADER_KEYWORDS = ["Product Description","Product Features","Technical Specifications","Product Specifications","Applications","Application Scenarios","Instructions","Installation","Notes","Product Images","Product Packaging"]

IMG_FRAME_PATTERN = re.compile(r'^\[IMG_FRAME:(\d+):(.*?)(?:\|S:(\d+))?\]$')
FLOAT_IMG_PATTERN = re.compile(r'^\[FLOAT_IMG:(left|right):(\d+):(.*)\]$')
URL_PATTERN = re.compile(r'https?://[^\s<>"\'，。、；：）\)\]\}]+')

FONT_CHOICES = ["微软雅黑","宋体","黑体","楷体","仿宋","Arial","Times New Roman","Calibri","Verdana","Georgia"]
BULLET_STYLES = {"●  实心圆":"● ","■  实心方":"■ ","▶  三角形":"▶ ","◆  菱形":"◆ ","○  空心圆":"○ ","①  带圈数字":"__NUM__","1.  数字编号":"__DOT__","—  短横线":"— "}
LANGUAGES = {"自动检测":"auto", "中文(简体)":"zh", "英语":"en", "日语":"jp", "韩语":"kor", "法语":"fra", "德语":"de", "俄语":"ru", "西班牙语":"spa", "葡萄牙语":"pt"}

# ═══════════════════════════════════════════
# 核心引擎类 (保留原版所有算法)
# ═══════════════════════════════════════════
class BaiduTranslator:
    API_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"
    MAX_BYTES = 4500

    @staticmethod
    def _baidu_api_call(q: str, from_lang: str, to_lang: str, appid: str, appkey: str) -> list:
        salt = random.randint(32768, 65536)
        sign = hashlib.md5((appid + q + str(salt) + appkey).encode("utf-8")).hexdigest()
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        payload = {"appid": appid, "q": q, "from": from_lang, "to": to_lang, "salt": salt, "sign": sign}
        resp = requests.post(BaiduTranslator.API_URL, data=payload, headers=headers, timeout=30)
        data = resp.json()
        if "error_code" in data: raise RuntimeError(f"百度翻译错误: {data.get('error_msg', '')}")
        return data.get("trans_result", [])

    @staticmethod
    def translate_paragraphs(paragraphs: list, from_lang: str, to_lang: str, appid: str, appkey: str) -> list:
        total = len(paragraphs); translated = [""] * total
        batch_idx = []; batch_lines = []; batch_bytes = 0

        def flush_batch():
            if not batch_idx: return
            q = "\n".join(batch_lines)
            results = BaiduTranslator._baidu_api_call(q, from_lang, to_lang, appid, appkey)
            for i, orig_idx in enumerate(batch_idx):
                translated[orig_idx] = results[i]["dst"] if i < len(results) else paragraphs[orig_idx]
            time.sleep(1.1)

        for idx, para in enumerate(paragraphs):
            if not para.strip(): translated[idx] = para; continue
            para_bytes = len(para.encode("utf-8"))
            if batch_bytes + para_bytes + 1 > BaiduTranslator.MAX_BYTES:
                flush_batch(); batch_idx.clear(); batch_lines.clear(); batch_bytes = 0
            batch_idx.append(idx); batch_lines.append(para); batch_bytes += para_bytes + 1
        if batch_idx: flush_batch()
        return translated

class WebFetcher:
    UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'
    @staticmethod
    def fetch(url, max_chars=8000):
        try:
            resp = requests.get(url, headers={'User-Agent': WebFetcher.UA}, timeout=30, verify=False)
            resp.encoding = resp.apparent_encoding or 'utf-8'
            html = resp.text
            body, table = "", ""
            if HAS_BS4:
                soup = BeautifulSoup(html, 'html.parser')
                tparts = []
                for tbl in soup.find_all('table'):
                    md = []
                    for row in tbl.find_all('tr'):
                        cells = [c.get_text(strip=True).replace('|','/').replace('\n',' ') for c in row.find_all(['td','th'])]
                        if any(cells): md.append("| "+" | ".join(cells)+" |")
                    if len(md) >= 2: tparts.append(md[0]+"\n|"+" | ".join(["---"]*(md[0].count('|')-1))+" |\n"+"\n".join(md[1:]))
                if tparts: table = "\n[规格表格]\n"+"\n\n".join(tparts)
                for tag in soup(['script','style','nav','footer']): tag.decompose()
                body = soup.get_text(separator='\n', strip=True)
            else:
                body = re.sub(r'<[^>]+>','\n', html)
            return body[:max_chars], table[:3000], ""
        except Exception as e: return "", "", str(e)

def baidu_image_search(keyword, pn=0, rn=20):
    results = []
    try:
        url = f'https://image.baidu.com/search/index?tn=baiduimage&word={quote(keyword)}&pn={pn}'
        resp = requests.get(url, headers={'User-Agent': WebFetcher.UA}, timeout=15, verify=False)
        img_urls = re.findall(r'"thumbURL"\s*:\s*"(https?://[^"]+)"', resp.text)
        for u in img_urls[:rn]: results.append({'thumb': u, 'desc': keyword})
    except: pass
    return results

# ═══════════════════════════════════════════
# Session State 初始化
# ═══════════════════════════════════════════
def init_state():
    defaults = {
        'txt_cn': "", 'txt_en': "", 'raw_text': "", 'current_lang': 'cn',
        'kimi_key': "", 'bd_id': "", 'bd_key': "",
        'gallery': [], 'extracted_images': [],
        'font_cn': "微软雅黑", 'font_en': "Arial", 'title_size': 14, 'body_size': 11,
        'line_gap': 8, 'title_gap': 14, 'bullet': "●  实心圆",
        'cover_cn': "光纤跳线系列", 'cover_en': "Fiber Optic Patch Cord",
        'opt_img': False, 'opt_pkg': False, 'feature_brief': True
    }
    for k, v in defaults.items():
        if k not in st.session_state: st.session_state[k] = v

# ═══════════════════════════════════════════
# 文本处理工具函数
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

def tool_remove_descriptions():
    txt_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
    lines = st.session_state[txt_key].split('\n')
    for i, l in enumerate(lines):
        if is_header(l) or '|' in l or 'IMG_FRAME' in l: continue
        m = re.search(r'^(.*?)(：|:)', l)
        if m and len(re.sub(r'^[-*•\s]+', '', m.group(1))) <= 30:
            lines[i] = l[:m.end()]
    st.session_state[txt_key] = '\n'.join(lines)
    st.toast("一键去说明完成！")

def tool_replace_numbers():
    txt_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
    lines = st.session_state[txt_key].split('\n')
    sym_val = BULLET_STYLES.get(st.session_state.bullet, "● ")
    if sym_val in ("__NUM__", "__DOT__"): sym_val = "- "
    for i, l in enumerate(lines):
        if is_header(l) or '|' in l or 'IMG_FRAME' in l: continue
        new_l = re.sub(r'^(\d+[\.、]|[①-⑳])\s*', sym_val, l.lstrip())
        if new_l != l.lstrip(): lines[i] = l[:len(l)-len(l.lstrip())] + new_l
    st.session_state[txt_key] = '\n'.join(lines)
    st.toast("一键去序号完成！")

def transpose_markdown_table(text):
    lines = text.split('\n')
    tables = []; current_table = []; start_idx = -1
    for i, l in enumerate(lines):
        if '|' in l:
            if not current_table: start_idx = i
            current_table.append(l)
        else:
            if current_table: tables.append((start_idx, current_table)); current_table = []
    if current_table: tables.append((start_idx, current_table))

    if not tables: return text
    
    # 默认只翻转第一个表格（网页版简化操作）
    start_idx, table_lines = tables[0]
    data = []
    for l in table_lines:
        if '---' in l: continue
        data.append([c.strip() for c in l.strip().strip('|').split('|')])
    
    if not data: return text
    mc = max(len(r) for r in data)
    for r in data: r.extend([""] * (mc - len(r)))
    
    tr = [[data[r][c] for r in range(len(data))] for c in range(mc)]
    new_table = []
    for ri, row in enumerate(tr):
        new_table.append("| " + " | ".join(row) + " |")
        if ri == 0: new_table.append("| " + " | ".join(["---"] * len(row)) + " |")
    
    lines[start_idx:start_idx+len(table_lines)] = new_table
    st.toast("已对检测到的表格进行行列互换！")
    return '\n'.join(lines)

# ═══════════════════════════════════════════
# AI 与 翻译动作
# ═══════════════════════════════════════════
def call_kimi(prompt):
    if not st.session_state.kimi_key:
        st.error("请先在左侧栏配置 Kimi API Key")
        return None
    try:
        res = requests.post(KIMI_API_URL, headers={"Authorization": f"Bearer {st.session_state.kimi_key}"},
                            json={"model": "moonshot-v1-8k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.2})
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e:
        st.error(f"AI 请求失败: {str(e)}")
        return None

def do_ai_write(lang):
    feat_rule = "特点只保留关键词，如「- 关键词」。" if not st.session_state.feature_brief else "特点包含简短说明。"
    if lang == 'cn':
        prompt = f"资料：\n{st.session_state.raw_text[:6000]}\n\n要求：不输出产品名，以 **产品描述** 开头。严格使用标题：**产品描述** / **产品特点** / **产品指标** / **应用场景** / **使用说明**。\n注意：产品指标必须是Markdown表格。{feat_rule} 列表使用短横线 - 。"
    else:
        prompt = f"Source:\n{st.session_state.raw_text[:6000]}\n\nRules: Start with **Product Description**. Use EXACT headings: **Product Description** / **Product Features** / **Product Specifications** / **Applications** / **Instructions**.\nSpecs MUST be a Markdown table. Use '-' for lists, NO numbers."
    
    with st.spinner("AI 撰写中..."):
        res = call_kimi(prompt)
        if res:
            if lang == 'cn': st.session_state.txt_cn = res
            else: st.session_state.txt_en = res

def do_translate(from_lang, to_lang):
    if not st.session_state.bd_id or not st.session_state.bd_key:
        st.error("请在侧边栏配置百度翻译 API Key")
        return
    txt_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
    content = st.session_state[txt_key]
    if not content.strip(): return
    
    lines = content.split('\n')
    to_trans = [l for l in lines if l.strip() and not '|' in l and 'IMG_FRAME' not in l]
    
    with st.spinner("调用百度翻译引擎..."):
        res = BaiduTranslator.translate_paragraphs(to_trans, LANGUAGES[from_lang], LANGUAGES[to_lang], st.session_state.bd_id, st.session_state.bd_key)
        
        # 简单替换逻辑 (保持原始格式)
        new_lines = []
        trans_idx = 0
        for l in lines:
            if l.strip() and not '|' in l and 'IMG_FRAME' not in l:
                # 尝试保持 markdown 标记
                prefix = re.match(r'^([-*•#]+\s*|\d+\.\s*)', l)
                p_str = prefix.group(1) if prefix else ""
                clean_res = re.sub(r'^[-*•#]+\s*', '', res[trans_idx])
                if l.strip().startswith('**') and l.strip().endswith('**'):
                    new_lines.append(f"{p_str}**{clean_res.replace('**','')}**")
                else:
                    new_lines.append(f"{p_str}{clean_res}")
                trans_idx += 1
            else:
                new_lines.append(l)
        
        target_key = 'txt_en' if to_lang == '英语' else 'txt_cn'
        st.session_state[target_key] = '\n'.join(new_lines)
        st.session_state.current_lang = 'en' if to_lang == '英语' else 'cn'
        st.toast("翻译完成！")

# ═══════════════════════════════════════════
# 文档提取与图片库
# ═══════════════════════════════════════════
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
# Word 导出生成逻辑 (完全复用 Windows 算法)
# ═══════════════════════════════════════════
def _get_bullet_prefix(index, style):
    sym = BULLET_STYLES.get(style, "● ")
    if sym == "__NUM__": return "①②③④⑤⑥⑦⑧⑨⑩"[index % 10] + " "
    elif sym == "__DOT__": return f"{index+1}. "
    return sym

def generate_word_document(lang):
    doc = Document()
    fn = st.session_state.font_cn if lang == 'cn' else st.session_state.font_en
    txt = st.session_state.txt_cn if lang == 'cn' else st.session_state.txt_en
    
    # 页面设置 A4
    sec = doc.sections[0]
    sec.page_width = Mm(210); sec.page_height = Mm(297)
    sec.top_margin = Mm(25.4); sec.bottom_margin = Mm(25.4); sec.left_margin = Mm(19.1); sec.right_margin = Mm(19.1)
    
    doc.styles['Normal'].font.name = fn
    if lang == 'cn': doc.styles['Normal']._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    # 封面
    cp = doc.add_paragraph(); cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(12): cp.add_run('\n')
    cover_title = st.session_state.cover_cn if lang == 'cn' else st.session_state.cover_en
    rn = cp.add_run(cover_title if cover_title else "Product Specification")
    rn.bold = True; rn.font.size = Pt(28); rn.font.name = fn
    if lang == 'cn': rn._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
    
    doc.add_page_break()
    bs = doc.add_section()
    
    # 正文解析算法
    lines = txt.split('\n')
    tr = []; bc = 0
    
    def arf(pa, text, size=st.session_state.body_size, bold=False):
        r = pa.add_run(text); r.font.size = Pt(size); r.font.name = fn; r.bold = bold
        if lang == 'cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), fn)
        return r

    for idx, line in enumerate(lines):
        if not line.strip(): bc = 0; continue
        lt = line.strip()
        
        # 图框解析
        fi = IMG_FRAME_PATTERN.match(lt)
        if fi:
            cols = int(fi.group(1)); paths = fi.group(2).split('|')
            scale = int(fi.group(3)) if fi.group(3) else 100
            # 网页版暂不处理极其复杂的表格嵌入图片，用占位符替代
            pa = doc.add_paragraph()
            arf(pa, f"[图框组: {cols}列, 包含图片: {len(paths)}张, 缩放: {scale}%]", bold=True)
            continue
            
        # 表格解析
        if '|' in lt:
            if '---' in lt: continue
            cells = [clean_markdown(c.strip()) for c in lt.strip().strip('|').split('|')]
            if not cells: continue
            if not tr: tr.append(cells)
            else:
                tl = len(tr[0])
                tr.append((cells + [""] * tl)[:tl])
            
            nxt = lines[idx+1].strip() if idx+1 < len(lines) else ""
            if not nxt or '|' not in nxt:
                t = doc.add_table(rows=len(tr), cols=len(tr[0])); t.style = 'Table Grid'
                for ri, rd in enumerate(tr):
                    for ci, val in enumerate(rd):
                        cl = t.cell(ri, ci); cl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        pc = cl.paragraphs[0]; pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        rc = pc.add_run(val); rc.font.size = Pt(st.session_state.body_size); rc.font.name = fn
                        if ri == 0: rc.bold = True
                tr = []; doc.add_paragraph()
            continue

        # 标题解析
        if is_header(lt):
            bc = 0
            pa = doc.add_paragraph()
            pa.paragraph_format.space_before = Pt(12)
            arf(pa, clean_markdown(lt), st.session_state.title_size, True)
            continue
            
        # 列表解析
        if lt.startswith(('-', '*', '•')):
            rt = clean_markdown(lt[1:])
            bt = _get_bullet_prefix(bc, st.session_state.bullet) + rt; bc += 1
            pa = doc.add_paragraph(); arf(pa, bt)
            continue
            
        # 普通正文
        pa = doc.add_paragraph(); arf(pa, clean_markdown(lt))
        
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# ═══════════════════════════════════════════
# 网页 UI 渲染
# ═══════════════════════════════════════════
def main():
    init_state()
    
    # --- 左侧配置栏 ---
    with st.sidebar:
        st.title("⚙️ 华脉系统设置")
        with st.expander("🔑 API 密钥 (必填)", expanded=True):
            st.session_state.kimi_key = st.text_input("Kimi API Key", value=st.session_state.kimi_key, type="password")
            st.session_state.bd_id = st.text_input("百度翻译 APP ID", value=st.session_state.bd_id)
            st.session_state.bd_key = st.text_input("百度翻译 APP Key", value=st.session_state.bd_key, type="password")
        
        with st.expander("📐 排版与字体", expanded=False):
            st.session_state.font_cn = st.selectbox("中文字体", FONT_CHOICES, index=0)
            st.session_state.font_en = st.selectbox("英文字体", FONT_CHOICES, index=5)
            st.session_state.title_size = st.number_input("标题字号", 8, 36, st.session_state.title_size)
            st.session_state.body_size = st.number_input("正文字号", 6, 24, st.session_state.body_size)
            st.session_state.bullet = st.selectbox("列表符号", list(BULLET_STYLES.keys()))
            
        with st.expander("附加与封面", expanded=False):
            st.session_state.cover_cn = st.text_input("中文封面", st.session_state.cover_cn)
            st.session_state.cover_en = st.text_input("英文封面", st.session_state.cover_en)
            st.session_state.feature_brief = st.checkbox("产品特点含简短说明", st.session_state.feature_brief)
    
    # --- 顶部工具栏 ---
    st.markdown("""
        <style>
        .stButton>button { width: 100%; border-radius: 4px; }
        .main-header { font-size: 24px; font-weight: bold; color: #007AFF; margin-bottom: 0px; }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="main-header">📄 规格书专业排版系统 - 网页增强版</div>', unsafe_allow_html=True)
    
    t_col1, t_col2, t_col3, t_col4 = st.columns([1, 1, 1.5, 1])
    with t_col1:
        if st.button("✂️ 一键去说明"): tool_remove_descriptions()
        if st.button("🔄 一键去序号"): tool_replace_numbers()
    with t_col2:
        if st.button("🔀 行列互换"): 
            txt_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
            st.session_state[txt_key] = transpose_markdown_table(st.session_state[txt_key])
            st.rerun()
        if st.button("📷 插入[图框]标记"):
            txt_key = 'txt_cn' if st.session_state.current_lang == 'cn' else 'txt_en'
            st.session_state[txt_key] += "\n\n[IMG_FRAME:2]"
            st.toast("已在文末插入两列图框标记")
            st.rerun()
    with t_col3:
        st.write("🌐 翻译")
        c1, c2, c3 = st.columns([3, 3, 4])
        from_l = c1.selectbox("从", list(LANGUAGES.keys()), index=0, label_visibility="collapsed")
        to_l = c2.selectbox("到", list(LANGUAGES.keys()), index=2, label_visibility="collapsed")
        if c3.button("执行翻译"): do_translate(from_l, to_l)
    with t_col4:
        st.write("📥 导出")
        word_cn = generate_word_document('cn')
        st.download_button("📄 导出中文 Word", data=word_cn, file_name="中文规格书.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        word_en = generate_word_document('en')
        st.download_button("📄 导出英文 Word", data=word_en, file_name="English_Spec.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    st.divider()

    # --- 主操作区 ---
    tab1, tab2, tab3 = st.tabs(["📝 文案编辑 (CN/EN)", "📂 素材提取与图库", "👁️ 实时预览"])
    
    with tab1:
        l_col, r_col = st.columns(2)
        with l_col:
            st.subheader("🇨🇳 中文文案")
            if st.button("✨ 启动 Kimi AI 撰写 (中文)"): do_ai_write('cn')
            st.session_state.txt_cn = st.text_area("Markdown 编辑器", value=st.session_state.txt_cn, height=500, label_visibility="collapsed")
            if st.session_state.txt_cn: st.session_state.current_lang = 'cn'
            
        with r_col:
            st.subheader("🇬🇧 English Copy")
            if st.button("✨ Start AI Writing (EN)"): do_ai_write('en')
            st.session_state.txt_en = st.text_area("Markdown Editor", value=st.session_state.txt_en, height=500, label_visibility="collapsed")
            if st.session_state.txt_en and not st.session_state.txt_cn: st.session_state.current_lang = 'en'

    with tab2:
        st.subheader("提取参考资料与图片")
        c1, c2 = st.columns(2)
        with c1:
            st.write("**1. 文档解析 (PDF/Word)**")
            doc_file = st.file_uploader("上传产品资料", type=['pdf', 'docx'])
            if st.button("提取文本与图片") and doc_file:
                with st.spinner("解析中..."):
                    if doc_file.name.endswith(".pdf"):
                        with pdfplumber.open(doc_file) as pdf:
                            st.session_state.raw_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    st.session_state.extracted_images = extract_images_from_file(doc_file)
                st.success("资料解析成功！可在左侧编辑区呼叫 AI。")
        with c2:
            st.write("**2. 网页抓取**")
            url = st.text_input("产品网页链接")
            if st.button("抓取网页内容") and url:
                with st.spinner("抓取中..."):
                    b, t, e = WebFetcher.fetch(url)
                    st.session_state.raw_text += f"\n\n{b}\n{t}"
                st.success("抓取成功！已追加到参考素材。")
        
        if st.session_state.extracted_images:
            st.write("**已提取的图片 (这些图片在导出Word时可替换占位符)**")
            st.write(st.session_state.extracted_images)

    with tab3:
        st.subheader("网页简易预览版")
        st.info("注：由于浏览器限制，网页版无法实现完全所见即所得的A4编辑。请在此确认文本结构，最终排版效果请点击上方【导出 Word】查看。")
        preview_txt = st.session_state.txt_cn if st.session_state.current_lang == 'cn' else st.session_state.txt_en
        
        # 简单替换标记以便在网页端显示
        preview_html = preview_txt.replace("[IMG_FRAME:", "**[📸 预留图片框]** 参数: ")
        with st.container(border=True):
            st.markdown(f"# {st.session_state.cover_cn if st.session_state.current_lang == 'cn' else st.session_state.cover_en}")
            st.divider()
            st.markdown(preview_html)

if __name__ == "__main__":
    main()
