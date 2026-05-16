from flask import Flask, render_template_string, request
import markdown
import re

app = Flask(__name__)

# =========================
# 核心渲染逻辑（复刻桌面版）
# =========================
def render_advanced(text):
    # 双语块处理 [cn]中文[/cn][en]English[/en]
    def replace_bilingual(match):
        cn = match.group(1)
        en = match.group(2)
        return f"""
        <div class="bilingual">
            <div class="cn">{cn}</div>
            <div class="en">{en}</div>
        </div>
        """
    text = re.sub(r'\[cn\](.*?)\[/cn\]\s*\[en\](.*?)\[/en\]', replace_bilingual, text, flags=re.S)

    # 图片块 [img]url[/img]
    text = re.sub(r'\[img\](.*?)\[/img\]', r'<div class="img-box"><img src="\1"></div>', text)

    # 表格增强（简单处理）
    html = markdown.markdown(text, extensions=['tables'])

    return html


# =========================
# 商业UI模板（核心）
# =========================
HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>海风米奇 · Web商业版</title>

<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">

<style>
body {
    background: #0f172a;
    color: #e5e7eb;
    font-family: "Segoe UI", sans-serif;
}

.topbar {
    background: #020617;
    padding: 12px 20px;
    font-size: 18px;
    font-weight: bold;
    border-bottom: 1px solid #1e293b;
}

.container-main {
    display: flex;
    height: calc(100vh - 60px);
}

/* 左侧输入 */
.left {
    width: 50%;
    padding: 15px;
    border-right: 1px solid #1e293b;
}

textarea {
    width: 100%;
    height: 100%;
    background: #020617;
    color: #e2e8f0;
    border: none;
    padding: 15px;
    font-size: 14px;
    border-radius: 10px;
}

/* 右侧预览 */
.right {
    width: 50%;
    padding: 20px;
    overflow-y: auto;
}

.preview {
    background: #020617;
    padding: 20px;
    border-radius: 12px;
    box-shadow: 0 0 20px rgba(0,0,0,0.4);
}

/* 双语 */
.bilingual {
    display: flex;
    gap: 20px;
    margin: 20px 0;
}
.bilingual .cn {
    width: 50%;
    font-weight: bold;
    color: #38bdf8;
}
.bilingual .en {
    width: 50%;
    color: #cbd5f5;
}

/* 图片 */
.img-box {
    text-align: center;
    margin: 20px 0;
}
.img-box img {
    max-width: 90%;
    border-radius: 10px;
    box-shadow: 0 0 10px #000;
}

/* 表格 */
table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
}
th, td {
    border: 1px solid #334155;
    padding: 8px;
}
th {
    background: #1e293b;
}

/* 工具栏 */
.toolbar {
    margin-bottom: 10px;
}
button {
    margin-right: 5px;
}
</style>
</head>

<body>

<div class="topbar">🌊 海风米奇奇迹版 · Web商业版</div>

<div class="container-main">

    <div class="left">
        <div class="toolbar">
            <button class="btn btn-sm btn-primary" onclick="insertText('[cn]中文[/cn][en]English[/en]')">双语</button>
            <button class="btn btn-sm btn-success" onclick="insertText('[img]图片URL[/img]')">图片</button>
            <button class="btn btn-sm btn-warning" onclick="insertText('| 列1 | 列2 |\\n|---|---|\\n| 内容 | 内容 |')">表格</button>
        </div>

        <form method="post">
            <textarea name="text">{{text}}</textarea>
            <br><br>
            <button class="btn btn-success">实时渲染</button>
        </form>
    </div>

    <div class="right">
        <div class="preview">
            {{content|safe}}
        </div>
    </div>

</div>

<script>
function insertText(txt){
    let textarea = document.querySelector("textarea");
    textarea.value += "\\n" + txt;
}
</script>

</body>
</html>
"""


# =========================
# 路由
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    text = ""
    content = ""

    if request.method == "POST":
        text = request.form.get("text")
        content = render_advanced(text)

    return render_template_string(HTML, text=text, content=content)


# =========================
# 启动
# =========================
if __name__ == "__main__":
    app.run(debug=True)
