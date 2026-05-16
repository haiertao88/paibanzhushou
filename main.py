import streamlit as st
import requests
from docx import Document
from docx.shared import Inches
import re
import os

# ========================
# 配置
# ========================
st.set_page_config(layout="wide", page_title="海风米奇奇迹版 Web 完全体")

API_KEY = "你的key"
API_URL = "https://api.moonshot.cn/v1/chat/completions"

# ========================
# AI生成（支持双语）
# ========================
def generate_ai(text, bilingual=False):

    if bilingual:
        extra = "并提供中英文双语（中文+English）"
    else:
        extra = "仅中文"

    prompt = f"""
生成专业产品规格书，要求：

结构：
【产品名称】
【产品描述】
【产品特点】
- xxx
【技术参数】
参数: 值

要求：
{extra}
可包含图片URL（如果适合）

输入：
{text}
"""

    res = requests.post(
        API_URL,
        headers={"Authorization": f"Bearer {API_KEY}"},
        json={
            "model": "kimi",
            "messages": [{"role": "user", "content": prompt}]
        }
    )

    return res.json()


# ========================
# 图片识别
# ========================
def extract_images(text):
    urls = re.findall(r'(https?://\S+\.(jpg|png|jpeg))', text)
    return [u[0] for u in urls]


# ========================
# ⭐ 排版引擎（强化版）
# ========================
def format_content(ai_data):

    text = ai_data["choices"][0]["message"]["content"]

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    content = {
        "title": "",
        "description": "",
        "features": [],
        "params": {},
        "images": [],
        "bilingual": []
    }

    current = None

    for line in lines:

        # 标题识别
        if "产品名称" in line:
            current = "title"
            continue
        if "产品描述" in line:
            current = "desc"
            continue
        if "产品特点" in line:
            current = "features"
            continue
        if "技术参数" in line:
            current = "params"
            continue

        # 内容解析
        if current == "title":
            content["title"] = line

        elif current == "desc":
            content["description"] += line + "\n"

        elif current == "features":
            if line.startswith("-"):
                content["features"].append(line[1:].strip())
            else:
                content["features"].append(line)

        elif current == "params":
            if ":" in line or "：" in line:
                k, v = line.replace("：", ":").split(":", 1)
                content["params"][k.strip()] = v.strip()

        else:
            content["description"] += line + "\n"

    # 图片提取
    content["images"] = extract_images(text)

    # 双语检测（简单规则）
    for f in content["features"]:
        if any(c.isalpha() for c in f):
            content["bilingual"].append(f)

    return content


# ========================
# ⭐ DOCX导出（完全体）
# ========================
def export_docx(content):

    doc = Document()

    # 标题
    doc.add_heading(content["title"], 0)

    # 描述
    doc.add_heading("产品描述", 1)
    doc.add_paragraph(content["description"])

    # 图片（自动插入）
    if content["images"]:
        doc.add_heading("产品图片", 1)
        for img in content["images"]:
            try:
                import requests
                r = requests.get(img)
                with open("temp.jpg", "wb") as f:
                    f.write(r.content)
                doc.add_picture("temp.jpg", width=Inches(4))
                os.remove("temp.jpg")
            except:
                pass

    # 特点
    doc.add_heading("产品特点", 1)
    for f in content["features"]:
        doc.add_paragraph(f"• {f}")

    # 参数表（增强）
    doc.add_heading("技术参数", 1)

    if content["params"]:
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"

        hdr = table.rows[0].cells
        hdr[0].text = "参数"
        hdr[1].text = "说明"

        for k, v in content["params"].items():
            row = table.add_row().cells
            row[0].text = k
            row[1].text = v

    file = "完整版规格书.docx"
    doc.save(file)

    return file


# ========================
# UI
# ========================
st.title("🌊 海风米奇奇迹版（完全体Web）")

col1, col2 = st.columns([1, 2])

# 左侧
with col1:

    st.subheader("🧩 控制面板")

    user_input = st.text_area("输入产品信息", height=220)

    bilingual = st.checkbox("🌍 中英双语")

    generate_btn = st.button("🚀 生成")
    export_btn = st.button("📄 导出Word")

# 右侧
with col2:

    st.subheader("📊 预览")

    if "data" not in st.session_state:
        st.session_state.data = None

    if generate_btn:

        if not user_input:
            st.warning("请输入内容")
        else:
            with st.spinner("AI生成中..."):
                ai = generate_ai(user_input, bilingual)
                parsed = format_content(ai)
                st.session_state.data = parsed

    data = st.session_state.data

    if data:

        st.markdown(f"# {data['title']}")

        st.markdown("### 📖 产品描述")
        st.write(data["description"])

        # 图片展示
        if data["images"]:
            st.markdown("### 🖼 产品图片")
            for img in data["images"]:
                st.image(img, width=300)

        # 特点
        st.markdown("### ✨ 产品特点")
        for f in data["features"]:
            st.write(f"• {f}")

        # 参数
        st.markdown("### 📌 技术参数")
        for k, v in data["params"].items():
            st.write(f"{k}：{v}")

    else:
        st.info("等待生成...")


# ========================
# 导出
# ========================
if export_btn:

    if st.session_state.data:
        file = export_docx(st.session_state.data)

        with open(file, "rb") as f:
            st.download_button("⬇️ 下载Word", f, file_name=file)
    else:
        st.warning("请先生成")
