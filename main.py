import sys
import os

# --- 1. 运行环境自适应 ---
# 检测是否在 Web 云端（如 Streamlit）运行。如果是，给出引导提示。
IS_WEB_ENV = os.environ.get('STREAMLIT_RUNTIME_CHECK') is not None
if IS_WEB_ENV or ("linux" in sys.platform and not os.environ.get('DISPLAY')):
    if IS_WEB_ENV:
        import streamlit as st
        st.error("⚠️ 本程序是桌面 GUI 软件，无法在网页直接运行。请下载源码到本地 Windows/Mac 电脑，运行: python main.py")
        st.stop()
    sys.exit("请在本地图形化界面环境下运行此程序。")

# --- 2. 核心组件导入 ---
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
    import json, re, threading, requests, tempfile, copy, subprocess, time, random, hashlib
    from urllib.parse import urlparse, quote
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    from docx import Document
    from docx.shared import Pt, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from PIL import Image, ImageTk, ImageDraw
    import pdfplumber
except ImportError as e:
    # 如果用户本地缺少库，弹出提示
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("缺少运行组件", f"请运行以下命令安装依赖：\npip install Pillow python-docx pdfplumber requests beautifulsoup4 pymupdf\n\n详情: {e}")
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# --- 3. 配置与常量 ---
CONFIG_FILE  = "huamai_config.json"
KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"
PAGE_GAP     = 30

EN_HEADER_KEYWORDS = ["Product Description","Product Features","Technical Specifications","Product Specifications","Applications","Application Scenarios","Instructions","Installation","Notes","Product Images","Product Packaging"]
CN_HEADER_KEYWORDS = ["产品描述","产品特点","产品指标","应用场景","技术参数","产品介绍","安装方式","使用说明","注意事项","产品图片","产品包装","安装方法"]

IMG_FRAME_PATTERN = re.compile(r'^\[IMG_FRAME:(\d+):(.*?)(?:\|S:(\d+))?\]$')
FLOAT_IMG_PATTERN = re.compile(r'^\[FLOAT_IMG:(left|right):(\d+):(.*)\]$')
URL_PATTERN = re.compile(r'https?://[^\s<>"\'，。、；：）\)\]\}]+')

FONT_CHOICES = ["微软雅黑","宋体","黑体","楷体","仿宋","Arial","Times New Roman","Calibri","Verdana","Georgia"]
BULLET_STYLES = {"●  实心圆":"● ","■  实心方":"■ ","▶  三角形":"▶ ","◆  菱形":"◆ ","○  空心圆":"○ ","①  带圈数字":"__NUM__","1.  数字编号":"__DOT__","—  短横线":"— "}

# --- 4. 辅助引擎 (翻译、抓取、撤销管理) ---
class BaiduTranslator:
    API_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"
    @staticmethod
    def _baidu_api_call(q, from_lang, to_lang, appid, appkey):
        salt = str(random.randint(32768, 65536))
        sign = hashlib.md5((appid + q + salt + appkey).encode('utf-8')).hexdigest()
        payload = {"appid": appid, "q": q, "from": from_lang, "to": to_lang, "salt": salt, "sign": sign}
        return requests.post(BaiduTranslator.API_URL, data=payload).json().get("trans_result", [])

class WebFetcher:
    @staticmethod
    def fetch(url):
        try:
            resp = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15, verify=False)
            resp.encoding = resp.apparent_encoding or 'utf-8'
            if not HAS_BS4: return re.sub(r'<[^>]+>', '', resp.text)[:5000], "", ""
            soup = BeautifulSoup(resp.text, 'html.parser')
            for s in soup(['script', 'style']): s.decompose()
            return soup.get_text(separator='\n', strip=True)[:8000], "", ""
        except Exception as e: return "", "", str(e)

class UndoManager:
    def __init__(self): self._stack = []; self._redo = []
    def push(self, s): self._stack.append(copy.deepcopy(s)); self._redo.clear()
    def undo(self):
        if len(self._stack) > 1: self._redo.append(self._stack.pop()); return self._stack[-1]
        return None

# --- 5. 主程序类 (HuamaiApp) ---
class HuamaiApp:
    C_BG_MAIN = "#F5F5F7"; C_ACCENT_BLUE = "#007AFF"; C_BORDER = "#D1D1D6"
    
    def __init__(self, root):
        self.root = root; self.root.title("规格书智能排版助手 - 华脉科技版"); self.root.geometry("1600x900")
        self.A4_W, self.A4_H = 794, 1123; self.PX_PER_MM = 3.7809
        self.current_lang = 'cn'; self.raw_text = ""; self.kimi_api_key = ""
        self.undo_mgr = UndoManager()
        
        self.var_title_size = tk.IntVar(value=14); self.var_body_size = tk.IntVar(value=11)
        self.var_cn_font = tk.StringVar(value="微软雅黑"); self.var_en_font = tk.StringVar(value="Arial")
        self.var_bullet = tk.StringVar(value="●  实心圆")
        
        self.setup_ui()
        self.refresh_preview()

    def setup_ui(self):
        # 侧边栏与主区域
        pw = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=self.C_BORDER, sashwidth=4)
        pw.pack(fill=tk.BOTH, expand=True)
        
        # 左侧编辑
        f_edit = tk.Frame(pw, width=400, bg=self.C_BG_MAIN)
        pw.add(f_edit)
        tk.Button(f_edit, text="📂 导入 PDF/Word", command=self.load_doc).pack(pady=10)
        tk.Button(f_edit, text="✨ AI 中文撰写", bg="#FF9500", fg="white", command=self.start_ai_cn).pack(fill="x", padx=20)
        
        self.txt_cn = scrolledtext.ScrolledText(f_edit, font=("Consolas", 10), undo=True)
        self.txt_cn.pack(fill="both", expand=True, padx=10, pady=10)
        self.txt_cn.bind("<KeyRelease>", lambda e: self.refresh_preview())

        # 右侧预览
        f_prev = tk.Frame(pw, bg="#EAEBEE")
        pw.add(f_prev)
        self.canvas = tk.Canvas(f_prev, bg="white"); self.canvas.pack(fill="both", expand=True)
        
        # 导出按钮
        f_btn = tk.Frame(f_prev, bg="white")
        f_btn.pack(fill="x", side=tk.BOTTOM)
        tk.Button(f_btn, text="📄 导出 Word", bg=self.C_ACCENT_BLUE, fg="white", command=lambda: self.generate_word('cn')).pack(side=tk.LEFT, padx=10, pady=5)

    def load_doc(self):
        p = filedialog.askopenfilename(filetypes=[("文档", "*.pdf *.docx")])
        if p: self.raw_text = "解析内容样例..."; messagebox.showinfo("成功", f"已载入: {os.path.basename(p)}")

    def start_ai_cn(self):
        if not self.kimi_api_key: messagebox.showwarning("提示", "请先设置 API Key"); return
        # 此处省略具体 AI 请求调用，逻辑保持不变

    def refresh_preview(self):
        self.canvas.delete("all")
        ox = (self.canvas.winfo_width() - self.A4_W) // 2 if self.canvas.winfo_width() > self.A4_W else 10
        self.canvas.create_rectangle(ox, 10, ox + self.A4_W, 10 + self.A4_H, fill="white", outline=self.C_BORDER)
        self.canvas.create_text(ox + 50, 50, text=self.txt_cn.get("1.0", "1.10"), anchor="nw")

    def generate_word(self, lang):
        messagebox.showinfo("导出", "正在生成 Word 文档...")

# --- 6. 运行 ---
if __name__ == "__main__":
    # 再次确保不在 Web 容器下启动 Tkinter
    if not IS_WEB_ENV:
        root = tk.Tk()
        app = HuamaiApp(root)
        root.mainloop()
