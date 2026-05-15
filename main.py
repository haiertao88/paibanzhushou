import streamlit as st
import os, re, json, requests, tempfile, copy, time, random, hashlib, io
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
import pdfplumber
from bs4 import BeautifulSoup

# --- 1. 基础配置与常量 ---
KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"
BAIDU_FANYI_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"

st.set_page_config(page_title="华脉规格书智能排版软件-网页版", layout="wide")

# --- 2. 核心处理引擎 ---
class Processor:
    @staticmethod
    def call_kimi_ai(prompt, api_key):
        if not api_key: return "错误：请在左侧侧边栏配置 Kimi API Key"
        headers = {"Authorization": f"Bearer {api_key}"}
        payload = {
            "model": "moonshot-v1-8k",
            "messages": [
                {"role": "system", "content": "你是一个专业的光纤通信工艺工程师，擅长编写产品规格书。"},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3
        }
        try:
            res = requests.post(KIMI_API_URL, headers=headers, json=payload, timeout=60)
            return res.json()["choices"][0]["message"]["content"]
        except Exception as e:
            return f"AI 撰写出错：{str(e)}"

    @staticmethod
    def fetch_web_content(url):
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            resp = requests.get(url, headers=headers, timeout=15, verify=False)
            resp.encoding = resp.apparent_encoding or 'utf-8'
            soup = BeautifulSoup(resp.text, 'html.parser')
            for tag in soup(['script', 'style', 'nav', 'footer']): tag.decompose()
            return soup.get_text(separator='\n', strip=True)[:8000]
        except Exception as e:
            return f"抓取失败: {str(e)}"

# --- 3. Word 导出逻辑 (完全保留原版排版算法) ---
def generate_word_file(content, title_size, body_size, cn_font, cover_name):
    doc = Document()
    # 页面设置
    section = doc.sections[0]
    section.page_height = Mm(297)
    section.page_width = Mm(210)
    section.top_margin = Mm(25.4)
    section.bottom_margin = Mm(25.4)
    section.left_margin = Mm(19.1)
    section.right_margin = Mm(19.1)

    # 绘制封面
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(12): cp.add_run('\n')
    run = cp.add_run(cover_name if cover_name else "产品规格书")
    run.bold = True
    run.font.size = Pt(28)
    run.font.name = cn_font
    run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), cn_font)
    
    doc.add_page_break()

    # 写入正文
    lines = content.split('\n')
    for line in lines:
        if not line.strip(): continue
        
        # 模拟原版的标题识别逻辑
        is_header = False
        clean_line = line.replace('*', '').strip()
        if line.strip().startswith('**') and line.strip().endswith('**'): is_header = True
        
        p = doc.add_paragraph()
        if is_header:
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run(clean_line)
            run.bold = True
            run.font.size = Pt(title_size)
        else:
            # 处理列表符
            display_text = line
            if line.strip().startswith(('-', '•', '*')):
                display_text = "● " + line.strip()[1:].strip()
            run = p.add_run(display_text)
            run.font.size = Pt(body_size)
            
        run.font.name = cn_font
        run._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'), cn_font)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream

# --- 4. 网页 UI 布局 ---
def main():
    st.sidebar.title("🛠️ 华脉系统设置")
    
    # API 配置
    with st.sidebar.expander("🔑 API 密钥配置", expanded=True):
        api_key = st.text_input("Kimi API Key", type="password")
        st.caption("用于 AI 智能撰写文案")

    # 排版参数
    with st.sidebar.expander("📐 排版与字体", expanded=True):
        cn_font = st.selectbox("中文字体", ["微软雅黑", "宋体", "黑体", "楷体"])
        title_size = st.slider("标题字号", 10, 24, 14)
        body_size = st.slider("正文字号", 8, 16, 11)
        cover_name = st.text_input("封面产品名称", "光纤跳线系列")

    st.sidebar.divider()
    st.sidebar.info("作为工艺工程师，您可以通过该网页版快速生成 A4 规格书。")

    # 主界面分栏
    st.title("🚀 华脉专业规格书智能排版软件")
    
    tab1, tab2 = st.tabs(["📝 文案创作", "📄 导出下载"])

    with tab1:
        col_in, col_out = st.columns([1, 1])
        
        with col_in:
            st.subheader("1. 素材录入")
            upload = st.file_uploader("上传参考资料 (PDF/DOCX)", type=["pdf", "docx"])
            web_url = st.text_input("或粘贴产品网页链接", placeholder="https://...")
            
            raw_material = ""
            if upload:
                if upload.name.endswith(".pdf"):
                    with pdfplumber.open(upload) as pdf:
                        raw_material = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                st.success("资料解析成功")

            if web_url:
                with st.spinner("抓取网页内容中..."):
                    raw_material += "\n" + Processor.fetch_web_content(web_url)

            st.subheader("2. AI 智能撰写")
            ai_req = st.text_area("额外要求", "请撰写专业规格书，包含：产品描述、特点、技术指标（表格格式）、应用场景。")
            
            if st.button("✨ 启动 AI 撰写", use_container_width=True):
                if not raw_material and not web_url:
                    st.error("请先提供参考素材或链接")
                else:
                    prompt = f"参考资料：\n{raw_material}\n要求：{ai_req}\n请输出完整的 Markdown 规格书文案。"
                    with st.spinner("Kimi 正在思考编写中..."):
                        res = Processor.call_kimi_ai(prompt, api_key)
                        st.session_state['web_content'] = res

        with col_out:
            st.subheader("3. 文案编辑预览")
            final_text = st.text_area("在此调整最终文案 (Markdown 格式)", 
                                     value=st.session_state.get('web_content', ""), 
                                     height=600)
            st.session_state['web_content'] = final_text

    with tab2:
        st.subheader("预览与导出成果")
        if st.session_state.get('web_content'):
            with st.container(border=True):
                st.markdown(st.session_state['web_content'])
            
            st.divider()
            
            # 生成 Word 下载按钮
            word_io = generate_word_file(
                st.session_state['web_content'], 
                title_size, body_size, cn_font, cover_name
            )
            
            st.download_button(
                label="📥 下载生成的 Word 规格书",
                data=word_io,
                file_name=f"{cover_name}_规格书.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        else:
            st.warning("暂无文案，请在『文案创作』页签点击生成。")

if __name__ == "__main__":
    main()
