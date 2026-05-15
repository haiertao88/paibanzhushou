import sys
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
    import os, json, re, threading, requests, tempfile, copy, subprocess, time, random, hashlib
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
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("缺少运行组件", f"启动失败，缺少必要的Python库。\n请在命令行(cmd)中运行以下命令安装：\n\npip install Pillow python-docx pdfplumber requests bs4\n\n详细报错：{e}")
    sys.exit()

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

CONFIG_FILE  = "huamai_config.json"
KIMI_API_URL = "https://api.moonshot.cn/v1/chat/completions"
PAGE_GAP     = 30

EN_HEADER_KEYWORDS = [
    "Product Description","Product Features","Technical Specifications",
    "Product Specifications","Applications","Application Scenarios",
    "Instructions","Installation","Notes","Product Images","Product Packaging"
]
CN_HEADER_KEYWORDS = [
    "产品描述","产品特点","产品指标","应用场景","技术参数",
    "产品介绍","安装方式","使用说明","注意事项","产品图片","产品包装","安装方法"
]

IMG_FRAME_PATTERN = re.compile(r'^\[IMG_FRAME:(\d+):(.*?)(?:\|S:(\d+))?\]$')
FLOAT_IMG_PATTERN = re.compile(r'^\[FLOAT_IMG:(left|right):(\d+):(.*)\]$')
URL_PATTERN = re.compile(r'https?://[^\s<>"\'，。、；：）\)\]\}]+')

FONT_CHOICES = ["微软雅黑","宋体","黑体","楷体","仿宋",
                "Arial","Times New Roman","Calibri","Verdana","Georgia"]
BULLET_STYLES = {
    "●  实心圆":"● ","■  实心方":"■ ","▶  三角形":"▶ ",
    "◆  菱形":"◆ ","○  空心圆":"○ ",
    "①  带圈数字":"__NUM__","1.  数字编号":"__DOT__","—  短横线":"— ",
}

# ═══════════════════════════════════════════
# 百度通用翻译引擎配置
# ═══════════════════════════════════════════
LANGUAGES = {
    "自动检测":    "auto", "中文(简体)":  "zh", "中文(繁体)":  "cht", "英语":        "en",
    "日语":        "jp",   "韩语":        "kor","法语":        "fra", "德语":        "de",
    "俄语":        "ru",   "西班牙语":    "spa","葡萄牙语":    "pt",  "意大利语":    "it",
    "阿拉伯语":    "ara",  "越南语":      "vie","泰语":        "th",  "马来语":      "may",
    "印尼语":      "id",   "荷兰语":      "nl", "波兰语":      "pl",  "土耳其语":    "tr",
}

class BaiduTranslator:
    API_URL = "https://fanyi-api.baidu.com/api/trans/vip/translate"
    MAX_BYTES = 4500

    @staticmethod
    def make_md5(s: str, encoding: str = "utf-8") -> str:
        return hashlib.md5(s.encode(encoding)).hexdigest()

    @staticmethod
    def _baidu_api_call(q: str, from_lang: str, to_lang: str, appid: str, appkey: str) -> list:
        salt = random.randint(32768, 65536)
        sign = BaiduTranslator.make_md5(appid + q + str(salt) + appkey)
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        payload = {"appid": appid, "q": q, "from": from_lang, "to": to_lang, "salt": salt, "sign": sign}
        resp = requests.post(BaiduTranslator.API_URL, data=payload, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if "error_code" in data:
            raise RuntimeError(f"百度翻译错误 [{data['error_code']}]: {data.get('error_msg', '')}")
        return data.get("trans_result", [])

    @staticmethod
    def translate_paragraphs(paragraphs: list, from_lang: str, to_lang: str, appid: str, appkey: str, progress_cb=None) -> list:
        total = len(paragraphs)
        translated = [""] * total
        valid_total = sum(1 for p in paragraphs if p.strip())
        if valid_total == 0: return translated
        
        done_count  = 0
        batch_idx   = []; batch_lines = []; batch_bytes = 0

        def flush_batch():
            nonlocal done_count
            if not batch_idx: return
            q = "\n".join(batch_lines)
            results = BaiduTranslator._baidu_api_call(q, from_lang, to_lang, appid, appkey)
            for i, orig_idx in enumerate(batch_idx):
                if i < len(results):
                    translated[orig_idx] = results[i]["dst"]
                else:
                    translated[orig_idx] = paragraphs[orig_idx]
            done_count += len(batch_idx)
            if progress_cb: progress_cb(done_count, valid_total)
            time.sleep(1.1)

        for idx, para in enumerate(paragraphs):
            if not para.strip():
                translated[idx] = para; continue
            para_bytes = len(para.encode("utf-8"))
            if para_bytes > BaiduTranslator.MAX_BYTES:
                if batch_idx:
                    flush_batch()
                    batch_idx.clear(); batch_lines.clear(); batch_bytes = 0
                translated[idx] = BaiduTranslator._translate_long(para, from_lang, to_lang, appid, appkey)
                done_count += 1
                if progress_cb: progress_cb(done_count, valid_total)
                time.sleep(1.1)
                continue

            if batch_bytes + para_bytes + 1 > BaiduTranslator.MAX_BYTES:
                flush_batch()
                batch_idx.clear(); batch_lines.clear(); batch_bytes = 0

            batch_idx.append(idx)
            batch_lines.append(para)
            batch_bytes += para_bytes + 1

        if batch_idx: flush_batch()
        return translated

    @staticmethod
    def _translate_long(text: str, from_lang: str, to_lang: str, appid: str, appkey: str) -> str:
        import re
        sentences = re.split(r'(?<=[。.!?！？])\s*', text)
        result_parts = []; cur = ""
        for s in sentences:
            if not s: continue
            if len((cur + s).encode("utf-8")) > BaiduTranslator.MAX_BYTES:
                if cur.strip():
                    res = BaiduTranslator._baidu_api_call(cur, from_lang, to_lang, appid, appkey)
                    result_parts.append(" ".join(r["dst"] for r in res))
                    time.sleep(1.1)
                cur = s
            else:
                cur += s
        if cur.strip():
            res = BaiduTranslator._baidu_api_call(cur, from_lang, to_lang, appid, appkey)
            result_parts.append(" ".join(r["dst"] for r in res))
        return " ".join(result_parts) if result_parts else text

# ═══════════════════════════════════════════
# 网页抓取引擎
# ═══════════════════════════════════════════
class WebFetcher:
    UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'
    
    @staticmethod
    def fetch(url, max_chars=8000):
        headers = {'User-Agent': WebFetcher.UA, 'Accept': 'text/html,*/*;q=0.8',
                   'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'}
        try:
            resp = requests.get(url, headers=headers, timeout=30, verify=False, allow_redirects=True)
            ct = resp.headers.get('Content-Type', '')
            if 'charset=' in ct:
                resp.encoding = ct.split('charset=')[-1].split(';')[0].strip()
            elif resp.encoding and resp.encoding.lower() in ('iso-8859-1','ascii'):
                m = re.search(rb'charset=["\']?([a-zA-Z0-9_-]+)', resp.content[:3000])
                resp.encoding = m.group(1).decode('ascii','ignore') if m else (resp.apparent_encoding or 'utf-8')
            elif not resp.encoding:
                resp.encoding = resp.apparent_encoding or 'utf-8'
            html = resp.text
            if len(html) < 100: return "", "", f"内容过短({len(html)}字)"
        except Exception as e:
            return "", "", f"请求失败: {str(e)[:80]}"
            
        body, table = "", ""
        if HAS_BS4:
            try:
                soup = BeautifulSoup(html, 'html.parser')
                tparts = []
                for tbl in soup.find_all('table'):
                    md = []
                    for row in tbl.find_all('tr'):
                        cells = [c.get_text(strip=True).replace('|','/').replace('\n',' ') for c in row.find_all(['td','th'])]
                        if any(cells): md.append("| "+" | ".join(cells)+" |")
                    if len(md) >= 2:
                        nc = md[0].count('|')-1
                        tparts.append(md[0]+"\n|"+" | ".join(["---"]*max(nc,1))+" |\n"+"\n".join(md[1:]))
                    elif md: tparts.append("\n".join(md))
                if tparts: table = "\n[规格表格]\n"+"\n\n".join(tparts)
                for tag in soup(['script','style','nav','footer','header','aside','iframe','noscript','form']):
                    tag.decompose()
                main = soup.find('main') or soup.find('article') or soup.find(class_=re.compile(r'product|detail|content',re.I)) or soup.body or soup
                body = main.get_text(separator='\n', strip=True)
            except: 
                body = re.sub(r'<[^>]+>','\n',re.sub(r'<script.*?</script>','',html,flags=re.DOTALL|re.I))
        else:
            body = re.sub(r'<script.*?</script>','',html,flags=re.DOTALL|re.I)
            body = re.sub(r'<style.*?</style>','',body,flags=re.DOTALL|re.I)
            body = re.sub(r'<[^>]+>','\n',body)
            
        lines = [l.strip() for l in body.split('\n') if l.strip() and len(l.strip())>2]
        body = '\n'.join(lines)[:max_chars]
        table = table[:3000]
        if len(body)+len(table) < 30: return "","",f"提取内容过少"
        return body, table, ""

    @staticmethod
    def fetch_multi(urls, max_total=10000):
        ab,at,ae,af = [],[],[],[]
        tl = 0
        for url in urls[:5]:
            b,t,e = WebFetcher.fetch(url, max_chars=max(3000,(max_total-tl)//max(1,5-len(af))))
            if e: ae.append(f"[{url}] {e}")
            if b or t:
                ab.append(f"\n来源:{url}\n{b}")
                if t: at.append(t)
                af.append(url)
                tl += len(b)+len(t)
            if tl >= max_total: break
        return '\n'.join(ab), '\n'.join(at), ae, af

def baidu_image_search(keyword, pn=0, rn=20):
    results = []
    try:
        ts = str(int(time.time()*1000))
        headers = {
            'User-Agent': WebFetcher.UA,
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Referer': 'https://image.baidu.com/search/index?tn=baiduimage&word=' + quote(keyword),
            'X-Requested-With': 'XMLHttpRequest',
        }
        params = {
            'tn': 'resultjson_com', 'logid': ts, 'ipn': 'rj',
            'ct': '201326592', 'is': '', 'fp': 'result', 'fr': '',
            'word': keyword, 'queryWord': keyword,
            'cl': '2', 'lm': '-1', 'ie': 'utf-8', 'oe': 'utf-8',
            'pn': str(pn), 'rn': str(rn), 'gsm': hex(pn)[2:]
        }
        resp = requests.get('https://image.baidu.com/search/acjson', params=params, headers=headers, timeout=15, verify=False)
        if resp.status_code == 200:
            try: data = resp.json()
            except:
                text = re.sub(r"'", '"', resp.text)
                try: data = json.loads(text)
                except: data = {}
            for item in data.get('data', []):
                if not isinstance(item, dict): continue
                thumb = item.get('thumbURL','') or item.get('middleURL','') or item.get('objURL','')
                if thumb and thumb.startswith('http'):
                    results.append({'thumb': thumb, 'desc': (item.get('fromPageTitleEnc','') or item.get('fromPageTitle','') or '')[:25]})
    except Exception as e:
        pass

    if not results:
        try:
            url2 = f'https://image.baidu.com/search/index?tn=baiduimage&word={quote(keyword)}&pn={pn}'
            resp2 = requests.get(url2, headers={'User-Agent': WebFetcher.UA}, timeout=15, verify=False)
            resp2.encoding = 'utf-8'
            img_urls = re.findall(r'"thumbURL"\s*:\s*"(https?://[^"]+)"', resp2.text)
            if not img_urls: img_urls = re.findall(r'"objURL"\s*:\s*"(https?://[^"]+)"', resp2.text)
            for u in img_urls[:rn]: results.append({'thumb': u, 'desc': keyword})
        except: pass
    return results

class UndoManager:
    def __init__(self, max_size=80):
        self._stack=[]; self._redo=[]; self._max=max_size
    def push(self, state):
        if self._stack and self._stack[-1] == state:
            return
        self._stack.append(copy.deepcopy(state))
        if len(self._stack)>self._max: self._stack.pop(0)
        self._redo.clear()
    def undo(self):
        if len(self._stack)>1:
            self._redo.append(self._stack.pop())
            return copy.deepcopy(self._stack[-1])
        return None
    def redo(self):
        if self._redo:
            s=self._redo.pop(); self._stack.append(s)
            return copy.deepcopy(s)
        return None
    def can_undo(self): return len(self._stack)>1
    def can_redo(self): return len(self._redo)>0

class HuamaiApp:
    # --- Apple 风格配色方案 ---
    C_BG_MAIN = "#F5F5F7"      
    C_BG_PANEL = "#FFFFFF"     
    C_BG_HDR = "#E5E5EA"       
    C_TEXT = "#1D1D1F"         
    C_TEXT_MUTED = "#86868B"   
    C_ACCENT_BLUE = "#007AFF"  
    C_ACCENT_GREEN = "#34C759" 
    C_ACCENT_ORANGE = "#FF9500"
    C_ACCENT_RED = "#FF3B30"   
    C_BORDER = "#D1D1D6"       
    C_ACCENT_PURPLE = "#AF52DE" # 新增用于AI功能的配色

    def __init__(self, root):
        self.root = root
        self.root.title("规格书专业排版系统 - 现代沉浸版")
        self.root.geometry("1800x980")
        self.root.configure(bg=self.C_BG_MAIN)

        self.layout_config = {"edit_w": 400, "gallery_w": 140, "bottom_h": 160}

        self.PX_PER_MM=3.7809; self.A4_W,self.A4_H=794,1123
        self.M_TB,self.M_LR=25.4,19.1
        self.raw_text=""; self.bg_cover_path=""; self.bg_body_path=""
        self.tk_cache={}; self.current_lang='cn'

        self.var_title_size=tk.IntVar(value=14); self.var_body_size=tk.IntVar(value=11)
        self.var_line_gap=tk.IntVar(value=8); self.var_title_gap=tk.IntVar(value=14)
        self.var_cn_font=tk.StringVar(value="微软雅黑"); self.var_en_font=tk.StringVar(value="Arial")
        self.var_bullet=tk.StringVar(value="●  实心圆"); self.var_auto_shrink=tk.BooleanVar(value=True)

        self.var_cover_cn = tk.StringVar(value="")
        self.var_cover_en = tk.StringVar(value="")
        self.var_feature_brief = tk.BooleanVar(value=True)
        self.var_custom_prompt = tk.BooleanVar(value=False)
        self._float_insert_mode = None
        self._float_path = None
        
        self._web_searching = False
        self._web_search_page = 0
        self._web_search_keyword = ""
        self._web_search_thumbs = []
        self._last_web_body = ""
        self._last_web_tables = ""
        self._last_web_errors = []

        self.gallery_paths=[]; self.gallery_thumbs=[]; self.gallery_sel=set()
        self.opt_product_image=tk.BooleanVar(value=False)
        self.opt_product_package=tk.BooleanVar(value=False)
        self.custom_sections=[]

        self._doc_extract_files  = []   
        self._doc_extract_imgs   = []   
        self._doc_extract_thumbs = []   
        self._bottom_panel_expanded = True  
        self._doc_extract_running = False 

        self._edit_items=[]; self._active_editor=None
        self._cursor_raw_text = None  
        self._right_click_item = None
        self.undo_mgr=UndoManager(); self._undo_timer=None
        self._fmt_brush_active = False   

        self.opt_sections_order = []  

        self._drag_mode = None
        self._drag_opt = None        
        self._drag_ghost = None      
        self._drag_line_id = None    
        self._drag_start_y = 0
        self._drag_hit = None        
        self._drag_activated = False 
        self._drag_press_cy = 0      
        self._animation_step = 0
        self._translating = False
        self._t_pressed = False

        # API 密钥变量
        self.kimi_api_key = ""
        self.baidu_appid = ""
        self.baidu_appkey = ""

        os.environ.pop("HTTP_PROXY",None); os.environ.pop("HTTPS_PROXY",None)

        try:
            self.load_config() 
            self.setup_ui()    
            self._save_undo_state(); self.refresh_preview()
            self.root.after(200, self._enforce_initial_layout)
            self._pulse_border() 
        except Exception as e:
            import traceback
            messagebox.showerror("启动错误",f"初始化失败：\n{traceback.format_exc()}")

    def _enforce_initial_layout(self):
        try:
            self.apply_layout()
        except Exception:
            pass

    def apply_layout(self):
        win_w = self.main_paned.winfo_width()
        if win_w < 800: win_w = self.root.winfo_width()
        if win_w < 800: win_w = 2000
        
        self.main_paned.sash_place(0, self.layout_config["edit_w"], 0)
        if self.right_panel_visible:
            self.main_paned.sash_place(1, win_w - self.layout_config["gallery_w"], 0)
            
        if self._bottom_panel_expanded:
            self.btm_body.config(height=self.layout_config["bottom_h"])
            self.doc_img_canvas.config(height=self.layout_config["bottom_h"] - 20)
            self.wicv.config(height=self.layout_config["bottom_h"] - 20)

    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("⚙️ 系统与 API 设置")
        win.geometry("400x520")
        win.configure(bg=self.C_BG_MAIN)
        win.resizable(False, False)
        
        # 布局配置
        tk.Label(win, text="[布局尺寸配置]", bg=self.C_BG_MAIN, fg=self.C_ACCENT_BLUE, font=("微软雅黑", 10, "bold")).pack(pady=(15,5))
        f_layout = tk.Frame(win, bg=self.C_BG_MAIN)
        f_layout.pack()
        
        tk.Label(f_layout, text="左侧编辑区宽度 (px):", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).grid(row=0, column=0, pady=2, sticky="e")
        v_edit_w = tk.IntVar(value=self.layout_config["edit_w"])
        tk.Entry(f_layout, textvariable=v_edit_w, relief="solid", bd=1, justify="center", width=12).grid(row=0, column=1, pady=2, padx=5)
        
        tk.Label(f_layout, text="右侧图库宽度 (px):", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).grid(row=1, column=0, pady=2, sticky="e")
        v_gal_w = tk.IntVar(value=self.layout_config["gallery_w"])
        tk.Entry(f_layout, textvariable=v_gal_w, relief="solid", bd=1, justify="center", width=12).grid(row=1, column=1, pady=2, padx=5)
        
        tk.Label(f_layout, text="底部面板高度 (px):", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).grid(row=2, column=0, pady=2, sticky="e")
        v_btm_h = tk.IntVar(value=self.layout_config["bottom_h"])
        tk.Entry(f_layout, textvariable=v_btm_h, relief="solid", bd=1, justify="center", width=12).grid(row=2, column=1, pady=2, padx=5)
        
        ttk.Separator(win, orient='horizontal').pack(fill='x', pady=15, padx=20)
        
        # API 密钥配置
        tk.Label(win, text="[智能 API 密钥配置]", bg=self.C_BG_MAIN, fg=self.C_ACCENT_BLUE, font=("微软雅黑", 10, "bold")).pack(pady=(0,5))
        
        tk.Label(win, text="Kimi API Key:", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).pack()
        v_kimi = tk.StringVar(value=self.kimi_api_key)
        tk.Entry(win, textvariable=v_kimi, relief="solid", bd=1, justify="center", width=42).pack(pady=2)

        tk.Label(win, text="百度翻译 APP ID:", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).pack(pady=(10,0))
        v_baidu_id = tk.StringVar(value=self.baidu_appid)
        tk.Entry(win, textvariable=v_baidu_id, relief="solid", bd=1, justify="center", width=42).pack(pady=2)

        tk.Label(win, text="百度翻译 APP Key (密钥):", bg=self.C_BG_MAIN, font=("微软雅黑", 9)).pack(pady=(10,0))
        v_baidu_key = tk.StringVar(value=self.baidu_appkey)
        tk.Entry(win, textvariable=v_baidu_key, relief="solid", bd=1, justify="center", width=42).pack(pady=2)
        
        def save():
            self.layout_config["edit_w"] = v_edit_w.get()
            self.layout_config["gallery_w"] = v_gal_w.get()
            self.layout_config["bottom_h"] = v_btm_h.get()
            self.kimi_api_key = v_kimi.get().strip()
            self.baidu_appid = v_baidu_id.get().strip()
            self.baidu_appkey = v_baidu_key.get().strip()
            self.save_config()
            self.apply_layout()
            self._render_gallery()
            win.destroy()
            messagebox.showinfo("保存成功", "系统布局与API配置已保存并生效。")
            
        tk.Button(win, text="保存并应用设置", command=save, bg=self.C_ACCENT_BLUE, fg="white", font=("微软雅黑", 10, "bold"), relief="flat", cursor="hand2", padx=25, pady=6).pack(pady=25)

    # ══════════════════════════════════════════
    # 状态栏与系统反馈
    # ══════════════════════════════════════════
    def _update_status_bar(self):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        text_content = txt.get("1.0", "end-1c")
        word_count = len(text_content.replace('\n', '').replace(' ', ''))
        img_count = self._count_img_frames_in_text()
        self.lbl_status_words.config(text=f"📝 字数: {word_count}")
        self.lbl_status_imgs.config(text=f"🖼️ 图框组: {img_count}")
        self.lbl_status_mode.config(text=f"🌐 语言: {'中文' if self.current_lang == 'cn' else 'English'}")

    def _pulse_border(self):
        if self._fmt_brush_active or self._float_insert_mode:
            self.canvas.delete("pulse_border")
            color = self.C_ACCENT_ORANGE if self._animation_step % 2 == 0 else self.C_ACCENT_BLUE
            w = self.canvas.winfo_width()
            h = self.canvas.winfo_height()
            self.canvas.create_rectangle(2, 2, w-2, h-2, outline=color, width=4, tags="pulse_border")
            self._animation_step += 1
        else:
            self.canvas.delete("pulse_border")
            self._animation_step = 0
        self.root.after(500, self._pulse_border)

    # ══════════════════════════════════════════
    # 撤销/重做
    # ══════════════════════════════════════════
    def _get_state(self):
        return {'cn':self.txt_cn.get("1.0",tk.END),'en':self.txt_en.get("1.0",tk.END)}
    def _set_state(self, st):
        self.txt_cn.delete("1.0",tk.END); self.txt_cn.insert("1.0",st['cn'].rstrip('\n'))
        self.txt_en.delete("1.0",tk.END); self.txt_en.insert("1.0",st['en'].rstrip('\n'))
        self._highlight_syntax(self.txt_cn); self._highlight_syntax(self.txt_en)
        self.refresh_preview()
    def _save_undo_state(self):
        self.undo_mgr.push(self._get_state()); self._update_undo_buttons()
    def _schedule_undo_save(self):
        if self._undo_timer: self.root.after_cancel(self._undo_timer)
        self._undo_timer=self.root.after(800,self._save_undo_state)
        
    def do_undo(self,event=None):
        if self._undo_timer:
            self.root.after_cancel(self._undo_timer)
            self._undo_timer = None
        curr = self._get_state()
        if not self.undo_mgr._stack or self.undo_mgr._stack[-1] != curr:
            self.undo_mgr.push(curr)
            
        s=self.undo_mgr.undo()
        if s: self._set_state(s)
        self._update_undo_buttons(); return "break"
        
    def do_redo(self,event=None):
        if self._undo_timer:
            self.root.after_cancel(self._undo_timer)
            self._undo_timer = None
        s=self.undo_mgr.redo()
        if s: self._set_state(s)
        self._update_undo_buttons(); return "break"
        
    def _update_undo_buttons(self):
        self.btn_undo.config(state=tk.NORMAL if self.undo_mgr.can_undo() else tk.DISABLED)
        self.btn_redo.config(state=tk.NORMAL if self.undo_mgr.can_redo() else tk.DISABLED)

    # ══════════════════════════════════════════
    # ★ 联动交互逻辑：图库与编辑区的自适应
    # ══════════════════════════════════════════
    def toggle_right_panel(self):
        if self.right_panel_visible:
            try:
                cur_x = self.main_paned.sash_coord(0)[0]
                new_x = int(cur_x * 1.5)
                self.main_paned.forget(self.gallery_panel)
                self.btn_toggle_right.config(text="◀ 展开图库", bg=self.C_ACCENT_BLUE, fg="#fff")
                self.right_panel_visible = False
                self.main_paned.sash_place(0, new_x, 0)
            except Exception: pass
        else:
            self.main_paned.add(self.gallery_panel, minsize=self.layout_config["gallery_w"]) 
            self.btn_toggle_right.config(text="隐藏图库 ▶", bg=self.C_BORDER, fg=self.C_TEXT)
            self.right_panel_visible = True
            try:
                cur_x = self.main_paned.sash_coord(0)[0]
                new_x = int(cur_x / 1.5)
                self.main_paned.sash_place(0, new_x, 0)
                self.root.update_idletasks()
                win_w = self.main_paned.winfo_width()
                self.main_paned.sash_place(1, win_w - self.layout_config["gallery_w"], 0)
            except Exception: pass

    # ══════════════════════════════════════════
    # UI 主体构建 
    # ══════════════════════════════════════════
    def setup_ui(self):
        self._ctx_menu = tk.Menu(self.root, tearoff=0, font=("微软雅黑", 10), bg=self.C_BG_PANEL, fg=self.C_TEXT, activebackground=self.C_ACCENT_BLUE)

        self.status_bar = tk.Frame(self.root, bg=self.C_BG_HDR, height=26)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_bar.pack_propagate(False)
        self.lbl_status_words = tk.Label(self.status_bar, text="📝 字数: 0", bg=self.C_BG_HDR, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8))
        self.lbl_status_words.pack(side=tk.LEFT, padx=10)
        self.lbl_status_imgs = tk.Label(self.status_bar, text="🖼️ 图框组: 0", bg=self.C_BG_HDR, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8))
        self.lbl_status_imgs.pack(side=tk.LEFT, padx=10)
        self.lbl_status_mode = tk.Label(self.status_bar, text="🌐 语言: 中文", bg=self.C_BG_HDR, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8))
        self.lbl_status_mode.pack(side=tk.RIGHT, padx=10)

        self.main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, sashwidth=4, bg=self.C_BORDER)
        self.main_paned.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.editor_frame = tk.Frame(self.main_paned, width=self.layout_config["edit_w"], bg=self.C_BG_MAIN) 
        self.editor_frame.pack_propagate(False)
        self.main_paned.add(self.editor_frame, minsize=320)

        f_edit_hdr = tk.Frame(self.editor_frame, bg=self.C_BG_HDR, height=38)
        f_edit_hdr.pack(fill="x"); f_edit_hdr.pack_propagate(False)
        tk.Label(f_edit_hdr, text="􀌇 文本与参数", font=("微软雅黑", 10, "bold"), bg=self.C_BG_HDR, fg=self.C_TEXT).pack(side=tk.LEFT, padx=10)
        
        # 尺寸与系统设置按钮
        tk.Button(f_edit_hdr, text="⚙️ 系统设置", command=self.open_settings, bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 9), relief="flat", cursor="hand2").pack(side=tk.RIGHT, padx=10)

        f_ai_top = tk.Frame(self.editor_frame, bg=self.C_BG_MAIN)
        f_ai_top.pack(fill="x", padx=14, pady=10)
        
        f_up_hdr = tk.Frame(f_ai_top, bg=self.C_BG_MAIN)
        f_up_hdr.pack(fill="x", pady=2)
        tk.Label(f_up_hdr, text="1. 基础配置", font=("微软雅黑", 9, "bold"), bg=self.C_BG_MAIN, fg=self.C_TEXT).pack(side=tk.LEFT)
        
        export_mb = tk.Menubutton(f_up_hdr, text="􀈿 导出", bg=self.C_ACCENT_BLUE, fg="white", font=("微软雅黑", 9, "bold"), relief="flat", padx=8, pady=2, cursor="hand2")
        export_menu = tk.Menu(export_mb, tearoff=0, font=("微软雅黑", 9), bg=self.C_BG_PANEL)
        export_menu.add_command(label="📄 导出 中文 Word", command=lambda: self.generate_word('cn'))
        export_menu.add_command(label="📄 导出 English Word", command=lambda: self.generate_word('en'))
        export_menu.add_separator()
        export_menu.add_command(label="📑 导出 PDF", command=self.export_pdf)
        export_mb.config(menu=export_menu)
        export_mb.pack(side=tk.RIGHT)

        f_up = tk.Frame(f_ai_top, bg=self.C_BG_MAIN)
        f_up.pack(fill="x", pady=4)
        tk.Button(f_up, text="导入文档 (PDF/DOCX)", command=self.load_doc, bg=self.C_BG_PANEL, fg=self.C_TEXT, font=("微软雅黑", 9), relief="solid", bd=1, padx=8, cursor="hand2").pack(side=tk.LEFT)
        self.lbl_info = tk.Label(f_up, text="等待资料解析...", fg=self.C_TEXT_MUTED, bg=self.C_BG_MAIN, font=("微软雅黑", 8))
        self.lbl_info.pack(side=tk.LEFT, padx=8)

        f_cover = tk.Frame(f_ai_top, bg=self.C_BG_MAIN)
        f_cover.pack(fill="x", pady=4)
        tk.Label(f_cover, text="封面:", font=("微软雅黑", 9), bg=self.C_BG_MAIN, fg=self.C_TEXT).pack(side=tk.LEFT)
        e_cn = tk.Entry(f_cover, textvariable=self.var_cover_cn, font=("微软雅黑", 9), width=12, relief="solid", bd=1, bg=self.C_BG_PANEL)
        e_cn.pack(side=tk.LEFT, padx=4)
        e_cn.bind("<KeyRelease>", lambda e: self.refresh_preview())
        tk.Label(f_cover, text="EN:", bg=self.C_BG_MAIN, font=("微软雅黑", 9), fg=self.C_TEXT).pack(side=tk.LEFT, padx=(4,0))
        e_en = tk.Entry(f_cover, textvariable=self.var_cover_en, font=("微软雅黑", 9), width=12, relief="solid", bd=1, bg=self.C_BG_PANEL)
        e_en.pack(side=tk.LEFT, padx=4)
        e_en.bind("<KeyRelease>", lambda e: self.refresh_preview())

        tk.Checkbutton(f_ai_top, text="产品特点含简短说明 (否则仅关键词)", variable=self.var_feature_brief, font=("微软雅黑", 9), bg=self.C_BG_MAIN, activebackground=self.C_BG_MAIN).pack(anchor="w", pady=2)

        f_ai_btns = tk.Frame(f_ai_top, bg=self.C_BG_MAIN)
        f_ai_btns.pack(fill="x", pady=(8, 0))
        tk.Button(f_ai_btns, text="✨ AI 中文撰写", command=self.start_ai_cn, bg=self.C_ACCENT_ORANGE, fg="white", font=("微软雅黑", 9, "bold"), relief="flat", cursor="hand2").pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 4))
        tk.Button(f_ai_btns, text="🌐 AI EN Writing", command=self.start_ai_en, bg=self.C_ACCENT_BLUE, fg="white", font=("微软雅黑", 9, "bold"), relief="flat", cursor="hand2").pack(side=tk.LEFT, fill="x", expand=True, padx=(4, 0))

        tk.Frame(self.editor_frame, height=1, bg=self.C_BORDER).pack(fill="x", padx=14, pady=8)

        f_tab=tk.Frame(self.editor_frame,bg=self.C_BG_MAIN); f_tab.pack(fill="x", padx=14, pady=0)
        self.tab_cn_btn=tk.Button(f_tab,text="中文文案",font=("微软雅黑",9,"bold"),bg=self.C_ACCENT_BLUE,fg="white",relief="flat",padx=12,cursor="hand2",command=self.switch_to_cn)
        self.tab_cn_btn.pack(side=tk.LEFT,padx=(0,4))
        self.tab_en_btn=tk.Button(f_tab,text="English Copy",font=("微软雅黑",9),bg=self.C_BG_HDR,fg=self.C_TEXT,relief="flat",padx=12,cursor="hand2",command=self.switch_to_en)
        self.tab_en_btn.pack(side=tk.LEFT,padx=4)
        self.tab_prompt_btn = tk.Button(f_tab, text="􀌇 自定义提示词", font=("微软雅黑", 9), bg=self.C_BG_HDR, fg=self.C_TEXT, relief="flat", padx=12, cursor="hand2", command=self.switch_to_prompt)
        self.tab_prompt_btn.pack(side=tk.LEFT, padx=4)

        self.txt_container=tk.Frame(self.editor_frame,bg=self.C_BG_MAIN); self.txt_container.pack(fill="both", expand=True, padx=14, pady=(8,14))
        self.txt_cn=scrolledtext.ScrolledText(self.txt_container,font=("Consolas",10),wrap=tk.WORD, relief="flat", bd=0, bg=self.C_BG_PANEL, spacing1=4, spacing3=4, spacing2=2)
        self.txt_cn.pack(fill="both",expand=True); self.txt_cn.bind("<KeyRelease>",self._on_text_change)
        
        self.txt_en=scrolledtext.ScrolledText(self.txt_container,font=("Consolas",10),wrap=tk.WORD, relief="flat", bd=0, bg=self.C_BG_PANEL, spacing1=4, spacing3=4, spacing2=2)
        self.txt_en.bind("<KeyRelease>",self._on_text_change)
        
        # 强制接管原生文本框撤销逻辑
        self.txt_cn.bind("<Control-z>", self.do_undo); self.txt_cn.bind("<Control-Z>", self.do_undo)
        self.txt_cn.bind("<Control-y>", self.do_redo); self.txt_cn.bind("<Control-Y>", self.do_redo)
        self.txt_en.bind("<Control-z>", self.do_undo); self.txt_en.bind("<Control-Z>", self.do_undo)
        self.txt_en.bind("<Control-y>", self.do_redo); self.txt_en.bind("<Control-Y>", self.do_redo)

        self.prompt_container = tk.Frame(self.txt_container, bg=self.C_BG_PANEL, relief="flat", bd=0)
        fp = tk.Frame(self.prompt_container, bg=self.C_BG_PANEL, padx=10, pady=10)
        fp.pack(fill="both", expand=True)
        tk.Checkbutton(fp, text="启用自定义提示词（可粘贴网址）", variable=self.var_custom_prompt, font=("微软雅黑", 9, "bold"), bg=self.C_BG_PANEL, activebackground=self.C_BG_PANEL, command=self._on_prompt_toggle).pack(anchor="w", pady=(0, 6))
        
        self.lp1 = tk.Label(fp, text="中文提示词：", bg=self.C_BG_PANEL, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8, "bold"))
        self.lp1.pack(anchor="w")
        self.tp_cn = tk.Text(fp, height=7, font=("Consolas", 9), wrap=tk.WORD, bg=self.C_BG_MAIN, relief="flat", bd=0)
        self.tp_cn.pack(fill="both", expand=True, pady=(2, 6))
        self.tp_cn.insert("1.0", "请根据产品网页撰写中文规格书。\n不输出产品名，从**产品描述**开始。\n章节：产品描述、产品特点、产品指标(表格)、应用场景、使用说明。\n\n⚠️注意：特点和应用场景必须用横线 \"-\" 作为项目符号，严禁使用数字序号(1.2.3.)。\n\n产品网址：\nhttps://")
        
        self.lp2 = tk.Label(fp, text="EN Prompt：", bg=self.C_BG_PANEL, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8, "bold"))
        self.lp2.pack(anchor="w")
        self.tp_en = tk.Text(fp, height=7, font=("Consolas", 9), wrap=tk.WORD, bg=self.C_BG_MAIN, relief="flat", bd=0)
        self.tp_en.pack(fill="both", expand=True, pady=(2, 6))
        self.tp_en.insert("1.0", "Write EN spec from product page.\nNo name. Start with **Product Description**.\nSections MUST exactly be: **Product Description**, **Product Features**, **Product Specifications** (MUST be a Markdown table), **Applications**, **Instructions**.\n\n⚠️CRITICAL: No numbered lists. Use '-' for bullets.\n\nURL:\nhttps://")
        
        fb2 = tk.Frame(fp, bg=self.C_BG_PANEL)
        fb2.pack(fill="x", pady=2)
        tk.Button(fb2, text="测试抓取", command=self._test_fetch, bg=self.C_ACCENT_BLUE, fg="white", font=("微软雅黑", 9), relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(fb2, text="查看内容", command=self._show_fetch, bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 9), relief="flat", cursor="hand2").pack(side=tk.LEFT)
        self.lbl_f = tk.Label(fb2, text="💡 粘贴网址后点击测试抓取", bg=self.C_BG_PANEL, fg=self.C_TEXT_MUTED, font=("微软雅黑", 8))
        self.lbl_f.pack(side=tk.LEFT, padx=8)
        
        self._set_prompt_ui(False)

        self.preview_frame = tk.Frame(self.main_paned, bg="#EAEBEE") 
        self.main_paned.add(self.preview_frame, minsize=500)

        f_prev_hdr = tk.Frame(self.preview_frame, bg=self.C_BG_HDR, height=38)
        f_prev_hdr.pack(fill="x"); f_prev_hdr.pack_propagate(False)
        
        self.right_panel_visible = True
        self.btn_toggle_right = tk.Button(f_prev_hdr, text="隐藏图库 ▶", font=("微软雅黑", 9), bg=self.C_BORDER, fg=self.C_TEXT, relief="flat", cursor="hand2", padx=12, command=self.toggle_right_panel)
        self.btn_toggle_right.pack(side=tk.RIGHT, padx=10, pady=6)
        tk.Label(f_prev_hdr, text="􀏆 预览与排版区", font=("微软雅黑", 10, "bold"), bg=self.C_BG_HDR, fg=self.C_TEXT).pack(side=tk.LEFT, padx=10)

        self._build_toolbar(self.preview_frame)
        self._build_bottom_panel(self.preview_frame)

        scroll_outer=tk.Frame(self.preview_frame,bg="#EAEBEE"); scroll_outer.pack(fill="both",expand=True)
        self.v_scr=tk.Scrollbar(scroll_outer,orient=tk.VERTICAL); self.v_scr.pack(side=tk.RIGHT,fill=tk.Y)
        self.h_scr=tk.Scrollbar(scroll_outer,orient=tk.HORIZONTAL); self.h_scr.pack(side=tk.BOTTOM,fill=tk.X)
        self.canvas=tk.Canvas(scroll_outer,bg="#EAEBEE",highlightthickness=0,yscrollcommand=self.v_scr.set,xscrollcommand=self.h_scr.set)
        self.canvas.pack(side=tk.LEFT,fill="both",expand=True)
        self.v_scr.config(command=self.canvas.yview); self.h_scr.config(command=self.canvas.xview)
        
        self.canvas.bind("<Configure>",lambda e:self.refresh_preview())
        self.canvas.bind("<Button-1>",self._on_drag_start)
        self.canvas.bind("<Double-Button-1>",self._on_canvas_double_click)
        self.canvas.bind("<Enter>",self._bind_canvas_scroll)
        self.canvas.bind("<Leave>",self._unbind_canvas_scroll)
        self.canvas.bind("<Button-3>", self._on_canvas_right_click)
        self.canvas.bind("<B1-Motion>",       self._on_drag_motion)
        self.canvas.bind("<ButtonRelease-1>", self._on_drag_release)

        self._build_gallery_panel()
        self.main_paned.add(self.gallery_panel, minsize=self.layout_config["gallery_w"]) 

        self.root.bind("<Control-z>",self.do_undo); self.root.bind("<Control-Z>",self.do_undo)
        self.root.bind("<Control-y>",self.do_redo); self.root.bind("<Control-Y>",self.do_redo)
        self.root.bind("<Escape>", lambda e: self._cancel_float_mode() if self._float_insert_mode else (self._deactivate_fmt_brush() if self._fmt_brush_active else None))
        
        # 绑定 T 键以控制图片缩放
        self.root.bind("<KeyPress-t>", lambda e: setattr(self, '_t_pressed', True))
        self.root.bind("<KeyRelease-t>", lambda e: setattr(self, '_t_pressed', False))
        self.root.bind("<KeyPress-T>", lambda e: setattr(self, '_t_pressed', True))
        self.root.bind("<KeyRelease-T>", lambda e: setattr(self, '_t_pressed', False))
        self._t_pressed = False

    def _highlight_syntax(self, txt_widget):
        txt_widget.tag_config("header", foreground=self.C_ACCENT_BLUE, font=("Consolas", 10, "bold"))
        txt_widget.tag_remove("header", "1.0", tk.END)
        content = txt_widget.get("1.0", tk.END).split('\n')
        for i, line in enumerate(content):
            if not line.strip(): continue
            is_hdr = False
            clean = self.clean_markdown(line)
            if re.match(r'^\*\*.+\*\*$', line.strip()) or line.strip().startswith('#'): is_hdr = True
            elif clean in CN_HEADER_KEYWORDS or clean in EN_HEADER_KEYWORDS: is_hdr = True
            if is_hdr:
                start = f"{i+1}.0"; end = f"{i+1}.end"
                txt_widget.tag_add("header", start, end)

    def _on_text_change(self,event=None):
        txt_widget = event.widget if event else (self.txt_cn if self.current_lang == 'cn' else self.txt_en)
        self._highlight_syntax(txt_widget)
        self._schedule_undo_save(); self.refresh_preview()
        
    def _on_opt_toggle(self):
        self._sync_opts_to_text()

    def _bind_canvas_scroll(self,e=None): self.root.bind_all("<MouseWheel>",self._canvas_scroll)
    def _unbind_canvas_scroll(self,e=None): self.root.unbind_all("<MouseWheel>")
    def _canvas_scroll(self,e): 
        if getattr(self, '_t_pressed', False):
            cx = self.canvas.canvasx(e.x)
            cy = self.canvas.canvasy(e.y)
            hit = self._find_edit_item(cx, cy)
            if hit and hit[4] == 'img_frame':
                raw_text = hit[5]
                fi = self._parse_img_frame_tag(raw_text)
                if fi and fi[0] == 1:
                    delta = 5 if e.delta > 0 else -5
                    new_scale = max(20, min(100, fi[2] + delta))
                    if new_scale != fi[2]:
                        self._update_img_scale(raw_text, new_scale)
                    return "break"
        self.canvas.yview_scroll(int(-1*(e.delta/120)),"units")
        
    def _update_img_scale(self, raw_text, new_scale):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                fi = self._parse_img_frame_tag(raw_text)
                if fi:
                    lines[i] = self._build_img_frame_tag(fi[0], fi[1], new_scale)
                    txt.delete("1.0", tk.END)
                    txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
                    self.refresh_preview()
                    if getattr(self, '_cursor_raw_text', None) == raw_text:
                        self._cursor_raw_text = lines[i]
                    break

    # ══════════════════════════════════════════
    # 提示词逻辑 
    # ══════════════════════════════════════════
    def _on_prompt_toggle(self):
        self._set_prompt_ui(self.var_custom_prompt.get())
    def _set_prompt_ui(self, en):
        st = tk.NORMAL if en else tk.DISABLED
        bg = self.C_BG_PANEL if en else self.C_BG_MAIN
        self.tp_cn.config(state=st, bg=bg)
        self.tp_en.config(state=st, bg=bg)

    def _test_fetch(self):
        s1, s2 = self.tp_cn.cget("state"), self.tp_en.cget("state")
        self.tp_cn.config(state=tk.NORMAL); self.tp_en.config(state=tk.NORMAL)
        pt = self.tp_cn.get("1.0", tk.END) + "\n" + self.tp_en.get("1.0", tk.END)
        self.tp_cn.config(state=s1); self.tp_en.config(state=s2)
        urls = [u for u in URL_PATTERN.findall(pt) if len(u) > 15 and 'example.com' not in u]
        if not urls: return messagebox.showinfo("提示", "未发现有效网址")
        self.lbl_f.config(text=f"🔍 抓取{len(urls)}个网址...", fg=self.C_ACCENT_ORANGE)
        self.root.update()
        def do():
            b, t, e, f = WebFetcher.fetch_multi(urls)
            self._last_web_body, self._last_web_tables, self._last_web_errors = b, t, e
            def up():
                if f: self.lbl_f.config(text=f"✅ 成功{len(f)}个，共{len(b)+len(t)}字", fg=self.C_ACCENT_GREEN)
                else: self.lbl_f.config(text=f"❌ 失败: {e[0][:60] if e else '未知'}", fg=self.C_ACCENT_RED)
            self.root.after(0, up)
        threading.Thread(target=do, daemon=True).start()

    def _show_fetch(self):
        win = tk.Toplevel(self.root)
        win.title("抓取预览")
        win.geometry("800x600")
        nb = ttk.Notebook(win)
        for title, content in [
            (f"正文({len(self._last_web_body)}字)", self._last_web_body or "无内容"),
            (f"表格({len(self._last_web_tables)}字)", self._last_web_tables or "无表格"),
            (f"日志({len(self._last_web_errors)}条)", '\n'.join(self._last_web_errors) or "无错误日志")
        ]:
            f = tk.Frame(nb)
            nb.add(f, text=title)
            t = scrolledtext.ScrolledText(f, font=("Consolas", 10), wrap=tk.WORD)
            t.pack(fill="both", expand=True)
            t.insert("1.0", content)
        nb.pack(fill="both", expand=True)

    def _build_prompt(self, lang):
        if self.var_custom_prompt.get():
            s1, s2 = self.tp_cn.cget("state"), self.tp_en.cget("state")
            self.tp_cn.config(state=tk.NORMAL); self.tp_en.config(state=tk.NORMAL)
            pt = self.tp_cn.get("1.0", tk.END).strip() if lang == 'cn' else self.tp_en.get("1.0", tk.END).strip()
            self.tp_cn.config(state=s1); self.tp_en.config(state=s2)
            
            urls = [u for u in URL_PATTERN.findall(pt) if len(u)>15 and 'example.com' not in u]
            wb, wt = "", ""
            if urls:
                self.root.after(0, lambda: self.lbl_info.config(text=f"🔍 抓取{len(urls)}个网页...", fg=self.C_ACCENT_ORANGE))
                b, t, e, f = WebFetcher.fetch_multi(urls)
                wb, wt = b, t
                self._last_web_body, self._last_web_tables, self._last_web_errors = b, t, e
                if f: self.root.after(0, lambda: self.lbl_info.config(text=f"🚀 已提取{len(f)}页({len(b)+len(t)}字)AI生成中", fg=self.C_ACCENT_ORANGE))
            clean = pt
            for u in urls: clean = clean.replace(u, '')
            clean = re.sub(r'\n{3,}', '\n\n', clean).strip()
            
            parts = [clean]
            if self.raw_text: parts.append(f"\n\n【上传资料】\n{self.raw_text[:6000]}")
            if wb: parts.append(f"\n\n【网页正文】\n{wb}")
            if wt: parts.append(f"\n\n【网页表格】\n{wt}")
            return '\n'.join(parts)
        else:
            return self._builtin_prompt(lang)

    def _builtin_prompt(self, lang):
        if lang == 'cn':
            feat = ("产品特点：- 关键词：简短说明（10个字以内）" if self.var_feature_brief.get() 
                    else "【核心要求】产品特点只保留关键词，绝对不能有解释说明。格式必须是「- 关键词」。")
            return (
                "你是一位严谨的通信工程规格书专家。请根据以下资料撰写专业的中文产品规格书。\n\n"
                f"【资料】\n{self.raw_text[:8000]}\n\n"
                "【格式要求】\n不输出产品名。必须以 **产品描述** 作为开头。\n"
                "正文章节必须严格原样使用以下标题：**产品描述** / **产品特点** / **产品指标** / **应用场景** / **使用说明** / **安装方法**\n\n"
                "【执行指令】\n①去除品牌名 ②产品描述80-100字连贯段落 ③产品特点≥5条\n"
                "④ ⚠️严格注意：**产品指标** 必须强制使用 Markdown 表格 格式（≥3列），绝对不能使用纯文本列表。\n"
                "⑤应用场景≥4条 ⑥使用说明 ⑦仅允许标题用**加粗** ⑧无URL ⑨无分隔线 ⑩表格前后需留空行\n"
                f"⑪{feat} ⑫特点/场景/安装绝对不要用数字序号(1.2.3.)，必须以短横线 \"- \" 开头作为项目符号。"
            )
        else:
            feat = ("Product Features: '- Keyword: short description (max 10 words)'" if self.var_feature_brief.get() 
                    else "[CRITICAL] Product features must be KEYWORDS ONLY. No descriptions or explanations. Format must be '- Keyword'.")
            return (
                "You are a professional telecom product spec writer. Write an English spec sheet from the source material.\n\n"
                f"[Source]\n{self.raw_text[:8000]}\n\n"
                "【FORMATTING RULES】\n"
                "CRITICAL: Do not output the product name. Start immediately with **Product Description**.\n"
                "Sections MUST strictly use these EXACT headings: **Product Description** / **Product Features** / **Product Specifications** / **Applications** / **Instructions**\n\n"
                "【INSTRUCTIONS】\n"
                "1) Remove all brand names. 2) Description should be a 60-80 words paragraph. 3) Provide 5+ features.\n"
                "4) CRITICAL: The **Product Specifications** section MUST strictly be formatted as a Markdown table with at least 3 columns. DO NOT output plain text for specifications.\n"
                f"5) {feat}\n"
                "6) Provide 4+ applications. 7) Use **bold** for headings only. No URLs.\n"
                "8) CRITICAL: Do NOT use numbers for lists. Features and Applications MUST start with a dash '- '."
            )

    # ══════════════════════════════════════════
    # ★ 智能写入核心逻辑 (智撰、智表、智行)
    # ══════════════════════════════════════════
    def _get_smart_context(self):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        draft = txt.get("1.0", tk.END).strip()
        ctx = "【产品相关核心资料】\n"
        if self.raw_text: ctx += self.raw_text[:5000] + "\n"
        ctx += "\n【当前正在编辑的规格书上下文】\n" + draft[:3000]
        return ctx

    def _call_smart_kimi(self, prompt, callback):
        if not self.kimi_api_key:
            messagebox.showerror("配置错误", "请先在右上角【⚙️ 系统设置】中配置 Kimi API Key 才能使用智能生成功能。")
            return
        
        self.lbl_info.config(text="✨ AI 智能创作中，请稍候...", fg=self.C_ACCENT_PURPLE)
        def run():
            try:
                res = requests.post(KIMI_API_URL, headers={"Authorization": f"Bearer {self.kimi_api_key}"},
                    json={"model": "moonshot-v1-8k", "messages": [{"role": "user", "content": prompt}], "temperature": 0.2})
                rj = res.json()
                if "choices" in rj:
                    nt = rj["choices"][0]["message"]["content"].strip()
                    self.root.after(0, lambda: callback(nt))
                else:
                    self.root.after(0, lambda: messagebox.showerror("AI错误", rj.get("error", {}).get("message", "")))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("网络异常", str(e)))
            finally:
                self.root.after(0, lambda: self.lbl_info.config(text="✅ 创作完成", fg=self.C_ACCENT_GREEN))
                self.root.after(3000, lambda: self.lbl_info.config(text=""))
        threading.Thread(target=run, daemon=True).start()

    def _ai_smart_write(self, raw_text):
        """智撰功能：点击标题后根据标题和全文写出正文内容"""
        ctx = self._get_smart_context()
        header = self.clean_markdown(raw_text)
        lang_str = "中文" if self.current_lang == 'cn' else "English"
        
        prompt = f"基于以下产品资料，请为规格书的【{header}】章节撰写一段专业的{lang_str}正文。\n要求：\n1. 充分结合上下文产品信息进行专业书写。\n2. 不要输出标题，直接输出正文内容。\n3. 如果是多点说明，请使用 '-' 作为列表符号。\n4. 不要换行过多，保持紧凑。\n\n{ctx}"
        
        def cb(result):
            txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
            lines = txt.get("1.0", "end-1c").split('\n')
            target_idx = len(lines)
            for i, l in enumerate(lines):
                if l.strip() == raw_text.strip() and l.strip():
                    target_idx = i + 1; break
            
            res_lines = [rl for rl in result.split('\n') if rl.strip()]
            for i, rl in enumerate(res_lines):
                lines.insert(target_idx + i, rl)
                
            txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()
            
        self._call_smart_kimi(prompt, cb)

    def _ai_smart_table(self, raw_text):
        """智表功能：点击表格单元格后自动填写该行的指标"""
        ctx = self._get_smart_context()
        lang_str = "中文" if self.current_lang == 'cn' else "English"
        
        prompt = f"基于以下产品资料，请智能补全规格书({lang_str})Markdown表格中的这一行数据：\n\n当前行：\n{raw_text}\n\n任务与严格要求：\n1. 请根据行中的已知指标名称（如第一列的文字等），从资料中提取准确对应的参数值，填补空缺的单元格。\n2. 只输出补全后的【完整一行】Markdown表格代码（必须包含|），绝对不要输出任何额外解释或表头。\n\n{ctx}"
        
        def cb(result):
            res_line = raw_text
            for rl in result.split('\n'):
                if '|' in rl:
                    res_line = rl.strip()
                    break
                    
            txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
            lines = txt.get("1.0", "end-1c").split('\n')
            for i, l in enumerate(lines):
                if l.strip() == raw_text.strip() and l.strip():
                    lines[i] = res_line; break
                    
            txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()
            
        self._call_smart_kimi(prompt, cb)

    def _ai_smart_line(self, raw_text):
        """智行功能：结合标题和上一行格式，智能续写同行内容"""
        ctx = self._get_smart_context()
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        
        # 往上找这行属于哪个标题
        header = "未知章节"
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                for j in range(i, -1, -1):
                    if self.is_header(lines[j]):
                        header = self.clean_markdown(lines[j])
                        break
                break
                
        lang_str = "中文" if self.current_lang == 'cn' else "English"
        
        prompt = f"基于以下产品资料，请在规格书({lang_str})的【{header}】章节中，顺着当前行的思路，智能续写出与之并列的【下一行内容】。\n\n当前行内容：\n{raw_text}\n\n任务与严格要求：\n1. 提取资料中其它尚未列出的相关产品特性进行续写。\n2. 格式、语气、缩进和前缀符号（如 - 或 ● 等）必须与“当前行”完全一致。\n3. 只输出续写的那【一行】纯文本内容，绝对不要换行，不要任何解释。\n\n{ctx}"
        
        def cb(result):
            res_line = result.replace('\n', ' ').strip()
            target_idx = len(lines)
            for i, l in enumerate(lines):
                if l.strip() == raw_text.strip() and l.strip():
                    target_idx = i + 1; break
                    
            lines.insert(target_idx, res_line)
            txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()
            
        self._call_smart_kimi(prompt, cb)

    # ══════════════════════════════════════════
    # ★ 右键菜单、精确行删除及格式刷
    # ══════════════════════════════════════════
    def _delete_line_by_exact_text(self, raw_text):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        for i in range(len(lines)):
            if lines[i].strip() == raw_text.strip() and lines[i].strip():
                del lines[i]
                break
        txt.delete("1.0", tk.END)
        txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt)
        self._save_undo_state(); self.refresh_preview()

    def _delete_section_by_exact_text(self, raw_text):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        src_start = -1
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                src_start = i; break
        if src_start == -1: return
        
        src_end = src_start + 1
        while src_end < len(lines):
            ls = lines[src_end].strip()
            if ls and self.is_header(ls): break
            src_end += 1
            
        del lines[src_start:src_end]
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt)
        self._save_undo_state(); self.refresh_preview()

    def _delete_entire_table(self, line_idx):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        all_lines = txt.get("1.0", "end-1c").split('\n')
        tbl_start = line_idx
        while tbl_start > 0 and '|' in all_lines[tbl_start - 1].strip(): tbl_start -= 1
        tbl_end = line_idx
        while tbl_end < len(all_lines) - 1 and '|' in all_lines[tbl_end + 1].strip(): tbl_end += 1
        del all_lines[tbl_start:tbl_end + 1]
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(all_lines).rstrip('\n'))
        self._highlight_syntax(txt)
        self._save_undo_state(); self.refresh_preview()

    def _insert_text_below(self, raw_text):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        target_idx = len(lines)
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                target_idx = i + 1; break
        new_text = "请输入正文..."
        lines.insert(target_idx, new_text)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state()
        self._cursor_raw_text = new_text; self.refresh_preview()
        self.root.after(150, lambda: self._auto_edit_by_raw(new_text))

    def toggle_fmt_brush(self):
        self._fmt_brush_active = not self._fmt_brush_active
        if self._fmt_brush_active:
            self.btn_fmt_brush.config(bg=self.C_ACCENT_ORANGE, text="🖌️ 格式刷（激活）")
            self.canvas.config(cursor="dotbox")
        else:
            self._deactivate_fmt_brush()

    def _deactivate_fmt_brush(self):
        self._fmt_brush_active = False
        self.btn_fmt_brush.config(bg=self.C_BG_HDR, text="🖌️ 格式刷")
        self.canvas.config(cursor="")

    def _apply_fmt_brush(self, line_type, raw_text):
        if line_type in ('header', 'opt_section'):
            messagebox.showinfo("格式刷", "该行已经是标题样式，无需处理。"); self._deactivate_fmt_brush(); return
        if line_type in ('cover', 'img_frame', 'table_cell', 'float_img'):
            messagebox.showinfo("格式刷", "封面标题、图片框、浮动图、表格不支持格式刷。"); return
            
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        
        target_idx = -1
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                target_idx = i; break
        
        if target_idx == -1: self._deactivate_fmt_brush(); return
            
        raw = lines[target_idx].strip()
        clean = re.sub(r'^[-*•]\s*', '', raw)
        clean = re.sub(r'^\*\*(.+)\*\*$', r'\1', clean).strip()
        if not clean: self._deactivate_fmt_brush(); return
            
        lines[target_idx] = f"**{clean}**"
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt)
        self._save_undo_state(); self.refresh_preview()
        self._deactivate_fmt_brush()

    def _downgrade_header(self, raw_text):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        target_idx = -1
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                target_idx = i; break
        
        if target_idx != -1:
            clean = self.clean_markdown(lines[target_idx])
            lines[target_idx] = clean
            txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt)
            self._save_undo_state(); self.refresh_preview()

    def _on_canvas_right_click(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        hit = self._find_edit_item(cx, cy)
        if not hit: return
        cid, line_idx, bbox, text, line_type, raw_text = hit[:6]
        self._right_click_item = hit
        self._cursor_raw_text = raw_text

        self.canvas.delete("cursor_highlight")
        x0, y0, x1, y1 = bbox
        self.canvas.create_rectangle(x0-2, y0-1, x1+2, y1+1, outline=self.C_ACCENT_RED, width=2, fill="", tags="cursor_highlight")

        self._ctx_menu.delete(0, tk.END)
        if line_type == 'table_cell':
            self._ctx_menu.add_command(label="✨ 智表：智能补全该行指标", command=lambda r=raw_text: self._ai_smart_table(r))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="➕ 在上方插入行", command=lambda: self._insert_table_row(line_idx, True))
            self._ctx_menu.add_command(label="➕ 在下方插入行", command=lambda: self._insert_table_row(line_idx, False))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="➕ 在左侧插入列", command=lambda: self._insert_table_col(line_idx, cx, True))
            self._ctx_menu.add_command(label="➕ 在右侧插入列", command=lambda: self._insert_table_col(line_idx, cx, False))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="🗑 删除此表格行", command=lambda: self._delete_table_row(line_idx))
            self._ctx_menu.add_command(label="🗑 删除此表格列", command=lambda: self._delete_table_col(line_idx, cx))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="🗑 删除整个表格", command=lambda: self._delete_entire_table(line_idx))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="🔀 行列互换此表格", command=self.transpose_table)
        elif line_type in ('img_frame', 'float_img'):
            self._ctx_menu.add_command(label="🗑 删除此图片", command=lambda r=raw_text: self._delete_line_by_exact_text(r))
        elif line_type == 'opt_section':
            self._ctx_menu.add_command(label="✨ 智撰：自动生成此节正文", command=lambda r=raw_text: self._ai_smart_write(r))
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label=f"🗑 移除章节: {text}", command=lambda: self._remove_opt_section(line_idx, text))
        elif line_type == 'cover':
            return
        else:
            if line_type == 'header':
                self._ctx_menu.add_command(label="✨ 智撰：自动生成此节正文", command=lambda r=raw_text: self._ai_smart_write(r))
                self._ctx_menu.add_separator()
            elif line_type in ('text', 'bullet'):
                self._ctx_menu.add_command(label="✨ 智行：智能续写下一行", command=lambda r=raw_text: self._ai_smart_line(r))
                self._ctx_menu.add_separator()

            self._ctx_menu.add_command(label="✏️ 编辑此行", command=lambda: self._on_canvas_double_click_at(cx, cy))
            if line_type == 'header':
                self._ctx_menu.add_command(label="⬇️ 降级为普通正文", command=lambda r=raw_text: self._downgrade_header(r))
            else:
                self._ctx_menu.add_command(label="🖌️ 格式刷：转为标题样式", command=lambda r=raw_text: self._apply_fmt_brush(line_type, r))
            self._ctx_menu.add_command(label="🗑 删除此行", command=lambda r=raw_text: self._delete_line_by_exact_text(r))
            if line_type == 'header':
                self._ctx_menu.add_command(label="🗑 删除整个章节", command=lambda r=raw_text: self._delete_section_by_exact_text(r))

        if line_type != 'cover':
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="➕ 在下方插入正文", command=lambda r=raw_text: self._insert_text_below(r))
            
            if line_type not in ('opt_section', 'img_frame', 'float_img'):
                frame_menu = tk.Menu(self._ctx_menu, tearoff=0, font=("微软雅黑", 10), bg=self.C_BG_PANEL)
                for n in [1, 2, 3, 4, 6]:
                    frame_menu.add_command(label=f"插入 {n} 图框", command=lambda c=n, r=raw_text: self._insert_photo_frame_at(r, c))
                self._ctx_menu.add_cascade(label="📷 在下方插入图框...", menu=frame_menu)

        self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _insert_photo_frame_at(self, raw_text, cols):
        sel = self.get_selected_paths()
        if not sel: return messagebox.showinfo("提示", "请先在右侧图库中选择图片")
        imgs = (sel + [None] * cols)[:cols]
        tag = self._build_img_frame_tag(cols, imgs)
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        target_idx = len(lines)
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                target_idx = i + 1; break
        lines.insert(target_idx, tag)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state()
        self._cursor_raw_text = tag; self.refresh_preview()

    def _remove_opt_section(self, line_idx, text):
        if text in ("产品图片", "Product Images"): self.opt_product_image.set(False)
        elif text in ("产品包装", "Product Packaging"): self.opt_product_package.set(False)
        else:
            for enabled, cn_var, en_var in self.custom_sections:
                if cn_var.get() == text or en_var.get() == text:
                    enabled.set(False); break
        self._sync_opts_to_text()

    def _sync_opts_to_text(self):
        changed = False
        for lang in ('cn', 'en'):
            txt = self.txt_cn if lang == 'cn' else self.txt_en
            lines = txt.get("1.0", "end-1c").split('\n')
            lines_strip = [l.strip() for l in lines]
            
            active = []
            if self.opt_product_image.get(): active.append('产品图片' if lang=='cn' else 'Product Images')
            if self.opt_product_package.get(): active.append('产品包装' if lang=='cn' else 'Product Packaging')
            for enabled, cn_v, en_v in self.custom_sections:
                if enabled.get(): active.append(cn_v.get().strip() if lang=='cn' else en_v.get().strip())
            
            lang_changed = False
            for lbl in active:
                if not lbl: continue
                h = f"**{lbl}**"
                if h not in lines_strip:
                    lines.extend(["", h])
                    lang_changed = True
            
            if lang_changed:
                txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
                self._highlight_syntax(txt)
                changed = True
                
        if changed: self._save_undo_state()
        self.refresh_preview()

    def _set_float_mode(self, mode):
        sel = self.get_selected_paths()
        if len(sel) != 1: return messagebox.showinfo("提示", "请先在右侧图库中选中【1张】图片")
        self._float_insert_mode = mode
        self._float_path = sel[0]
        self.canvas.config(cursor="crosshair")
        self.lbl_float_hint.config(text=f"点击正文插入{'左' if mode=='left' else '右'}浮动图")

    def _cancel_float_mode(self):
        self._float_insert_mode = None; self._float_path = None
        self.canvas.config(cursor="")
        self.lbl_float_hint.config(text="")

    # ══════════════════════════════════════════
    # ★ 统一拖拽系统
    # ══════════════════════════════════════════
    def _on_drag_start(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)

        if self._fmt_brush_active:
            hit = self._find_edit_item(cx, cy)
            if hit: self._apply_fmt_brush(hit[4], hit[5])
            return

        if self._float_insert_mode:
            txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
            lines = txt.get("1.0", "end-1c").split('\n')
            target_idx = len(lines)
            ox = self._canvas_offset(); ml = int(self.M_LR * self.PX_PER_MM)
            cw = self.A4_W - ml * 2; iw = int(cw * 0.40); ih = int(iw * 0.75) 
            ideal_y0 = cy - ih / 2
            valid_items = [i for i in self._edit_items if i[4] not in ('cover', 'float_img')]
            
            if valid_items:
                closest_item = None; min_dist = float('inf')
                for item in valid_items:
                    dist = abs((item[2][1] + item[2][3]) / 2 - ideal_y0)
                    if dist < min_dist: min_dist = dist; closest_item = item
                if closest_item:
                    raw = closest_item[5]
                    for i, l in enumerate(lines):
                        if l.strip() == raw.strip() and l.strip():
                            target_idx = i
                            if ideal_y0 > (closest_item[2][1] + closest_item[2][3]) / 2: target_idx = i + 1
                            break

            tag = f"[FLOAT_IMG:{self._float_insert_mode}:40:{self._float_path}]"
            lines.insert(target_idx, tag)
            txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt)
            self._cancel_float_mode(); self._save_undo_state()
            self._cursor_raw_text = tag; self.refresh_preview()
            return

        self._drag_activated = False; self._drag_mode = None
        self._drag_hit = self._find_edit_item(cx, cy)
        self._drag_press_cy = cy; self._drag_start_y = cy
        
        if self._drag_hit and self._drag_hit[4] != 'opt_section':
            self._on_canvas_click(event)
        elif not self._drag_hit:
            self._dismiss_editor()
            self._cursor_raw_text = None
            self.canvas.delete("cursor_highlight")

    def _on_drag_motion(self, event):
        cy = self.canvas.canvasy(event.y)
        if not self._drag_activated:
            if abs(cy - self._drag_press_cy) < 8: return
            if not self._drag_hit: return
            _, line_idx, bbox, text, line_type, raw_text = self._drag_hit[:6]
            if line_type in ('header', 'opt_section'): self._drag_mode = 'section'
            elif line_type in ('bullet', 'text', 'img_frame', 'float_img'): self._drag_mode = 'row'
            else: return  
            self._drag_activated = True
            self.canvas.config(cursor="fleur")
            self._create_drag_ghost()

        if not self._drag_activated: return
        self._update_ghost_pos(cy)
        self.canvas.delete("drag_insert_line")
        insert_li, line_y = self._calc_insert_pos(cy)
        if line_y is not None:
            ox = self._canvas_offset(); ml = int(self.M_LR * self.PX_PER_MM)
            color = {"section": self.C_ACCENT_BLUE, "row": self.C_ACCENT_GREEN}.get(self._drag_mode, self.C_ACCENT_ORANGE)
            self.canvas.create_line(ox+ml, line_y, ox+self.A4_W-ml, line_y, fill=color, width=3, dash=(8,4), tags="drag_insert_line")

    def _on_drag_release(self, event):
        self.canvas.config(cursor=""); self.canvas.delete("drag_insert_line")
        if self._drag_ghost:
            self.canvas.delete(self._drag_ghost); self._drag_ghost = None
        if self._drag_activated and self._drag_hit and self._drag_mode:
            cy = self.canvas.canvasy(event.y)
            insert_li, _ = self._calc_insert_pos(cy)
            if self._drag_mode == 'section': self._finish_drag_section(insert_li)
            elif self._drag_mode == 'row': self._finish_drag_row(insert_li)

        self._drag_activated = False
        self._drag_mode = None; self._drag_hit = None

    def _create_drag_ghost(self):
        self.canvas.delete("drag_ghost")
        if not self._drag_hit: return
        _, line_idx, bbox, text, line_type, raw_text = self._drag_hit[:6]
        if self._drag_mode == 'section': x0, y0, x1, y1 = self._get_section_canvas_bbox(line_idx)
        else: x0, y0, x1, y1 = bbox
        oc, fc = {"section": (self.C_ACCENT_BLUE, "#e8f0fe"), "row": (self.C_ACCENT_GREEN, "#eafaf1")}.get(self._drag_mode, (self.C_ACCENT_ORANGE, "#fff3e0"))
        self._drag_ghost = self.canvas.create_rectangle(x0-2, y0-2, x1+2, y1+2, outline=oc, width=2, dash=(6,3), fill=fc, stipple="gray25", tags="drag_ghost")
        self._ghost_orig_y0, self._ghost_orig_y1 = y0, y1

    def _update_ghost_pos(self, cy):
        if not self._drag_ghost: return
        dy = cy - self._drag_press_cy
        coords = self.canvas.coords(self._drag_ghost)
        if coords: self.canvas.coords(self._drag_ghost, coords[0], self._ghost_orig_y0+dy-2, coords[2], self._ghost_orig_y1+dy+2)

    def _calc_insert_pos(self, mouse_cy):
        candidates = [(i[1], i[2]) for i in self._edit_items if i[4] not in ('cover',) and i[1] != self._drag_hit[1]]
        if not candidates: return None, None
        candidates.sort(key=lambda x: (x[1][1]+x[1][3])/2)
        for li, bbox in candidates:
            if mouse_cy < (bbox[1]+bbox[3])/2: return li, bbox[1] - 3
        return None, candidates[-1][1][3] + 3

    def _get_section_canvas_bbox(self, header_line_idx):
        ox = self._canvas_offset(); ml = int(self.M_LR * self.PX_PER_MM)
        x0 = ox + ml; x1 = ox + self.A4_W - ml
        y0 = y1 = None; in_sec = False
        for item in self._edit_items:
            _, li, bbox, _, lt_type, raw_text = item[:6]
            if li == header_line_idx and lt_type in ('header','opt_section'):
                y0 = bbox[1]; y1 = bbox[3]; in_sec = True; continue
            if in_sec:
                if lt_type in ('header','opt_section','cover'): break
                y1 = max(y1, bbox[3])
        if y0 is None: y0, y1 = self._drag_hit[2][1], self._drag_hit[2][3]
        return x0, y0, x1, y1

    def _finish_drag_section(self, insert_li):
        raw_text = self._drag_hit[5]
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        src_start = -1
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                src_start = i; break
        if src_start == -1: return
        src_end = src_start + 1
        while src_end < len(lines):
            ls = lines[src_end].strip()
            if ls and self.is_header(ls): break
            src_end += 1
        src_block = lines[src_start:src_end]
        while src_block and not src_block[-1].strip(): src_block.pop()
        remaining = lines[:src_start] + lines[src_end:]
        al = self._get_all_lines_fixed()
        if insert_li is None or insert_li >= len(al):
            while remaining and not remaining[-1].strip(): remaining.pop()
            remaining.extend([''] + src_block)
        else:
            target_text = al[insert_li].strip()
            target_idx = -1
            for i, l in enumerate(remaining):
                if l.strip() == target_text and l.strip():
                    target_idx = i; break
            if target_idx != -1:
                insert_block = src_block + ['']
                for i, row in enumerate(insert_block): remaining.insert(target_idx + i, row)
            else:
                while remaining and not remaining[-1].strip(): remaining.pop()
                remaining.extend([''] + src_block)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(remaining).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()

    def _finish_drag_row(self, insert_li):
        raw_text = self._drag_hit[5]
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        src_idx = -1
        for i, l in enumerate(lines):
            if l.strip() == raw_text.strip() and l.strip():
                src_idx = i; break
        if src_idx == -1: return
        al = self._get_all_lines_fixed()
        target_text = al[insert_li].strip() if insert_li is not None and insert_li < len(al) else None
        src_row = lines.pop(src_idx)
        if not target_text: lines.append(src_row)
        else:
            target_idx = -1
            for i, l in enumerate(lines):
                if l.strip() == target_text and l.strip():
                    target_idx = i; break
            if target_idx != -1: lines.insert(target_idx, src_row)
            else: lines.append(src_row)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()

    def _on_canvas_double_click_at(self, cx, cy):
        hit = self._find_edit_item(cx, cy)
        if not hit: return
        cid, line_idx, bbox, text, line_type, raw_text = hit[:6]
        if line_type in ('img_frame', 'float_img', 'opt_section', 'cover'): return
        self._dismiss_editor()
        
        col_idx = hit[6] if line_type == 'table_cell' and len(hit) > 6 else None
        
        x0, y0, x1, y1 = bbox
        ew, eh = max(x1-x0, 200), max(y1-y0, 24)
        entry = tk.Entry(self.canvas, font=self._get_edit_font(line_type),
                         bg="#ffffdd", fg="#000", relief="solid", borderwidth=2,
                         insertbackground="#000")
        entry.insert(0, text); entry.select_range(0, tk.END)
        win_id = self.canvas.create_window(x0, y0, window=entry, anchor="nw", width=ew, height=eh+6)
        self._active_editor = {'entry': entry, 'win_id': win_id, 'line_idx': line_idx,
                               'line_type': line_type, 'original_text': text, 'raw_text': raw_text, 'col_idx': col_idx}
        entry.focus_set(); entry.bind("<Return>", self._commit_edit)
        entry.bind("<Escape>", lambda e: self._dismiss_editor()); entry.bind("<FocusOut>", self._commit_edit)

    def _delete_table_row(self, line_idx):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        all_lines = txt.get("1.0", "end-1c").split('\n')
        if not (0 <= line_idx < len(all_lines)): return
        del_indices = [line_idx]
        is_first_row = True
        for i in range(line_idx - 1, -1, -1):
            lt = all_lines[i].strip()
            if '|' in lt and lt:
                is_first_row = False; break
            else: break
        if is_first_row and line_idx + 1 < len(all_lines):
            next_line = all_lines[line_idx + 1].strip()
            if '|' in next_line and '---' in next_line: del_indices.append(line_idx + 1)
        for i in sorted(del_indices, reverse=True):
            if 0 <= i < len(all_lines): del all_lines[i]
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(all_lines).rstrip('\n'))
        self._save_undo_state(); self.refresh_preview()

    def _delete_table_col(self, line_idx, click_cx):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        all_lines = txt.get("1.0", "end-1c").split('\n')
        if not (0 <= line_idx < len(all_lines)): return
        tbl_start = line_idx
        while tbl_start > 0 and '|' in all_lines[tbl_start - 1].strip(): tbl_start -= 1
        tbl_end = line_idx
        while tbl_end < len(all_lines) - 1 and '|' in all_lines[tbl_end + 1].strip(): tbl_end += 1
        
        ox = self._canvas_offset(); m_l = int(self.M_LR * self.PX_PER_MM); c_w = self.A4_W - m_l * 2
        cur_parts = all_lines[line_idx].strip().strip('|').split('|')
        nc = len(cur_parts)
        if nc <= 1: return messagebox.showinfo("提示", "表格仅剩一列，无法再删除")
        
        col_w = c_w / nc; rel_x = click_cx - ox - m_l
        col_idx = max(0, min(int(rel_x / col_w), nc - 1))
        
        for i in range(tbl_start, tbl_end + 1):
            lt = all_lines[i].strip()
            if '|' in lt:
                parts = lt.strip().strip('|').split('|')
                if col_idx < len(parts): parts.pop(col_idx)
                all_lines[i] = "| " + " | ".join([p.strip() for p in parts]) + " |"
                
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(all_lines).rstrip('\n'))
        self._save_undo_state(); self.refresh_preview()

    def _insert_table_row(self, line_idx, above=True):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        all_lines = txt.get("1.0", "end-1c").split('\n')
        if not (0 <= line_idx < len(all_lines)): return
        row_text = all_lines[line_idx]
        
        parts = row_text.strip().strip('|').split('|')
        nc = len(parts) if parts else 1
        
        new_row = "| " + " | ".join([" "] * nc) + " |"
        insert_idx = line_idx if above else line_idx + 1
        if not above and line_idx + 1 < len(all_lines) and '---' in all_lines[line_idx+1]: insert_idx = line_idx + 2
        all_lines.insert(insert_idx, new_row)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(all_lines).rstrip('\n'))
        self._save_undo_state(); self.refresh_preview()

    def _insert_table_col(self, line_idx, click_cx, left=True):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        all_lines = txt.get("1.0", "end-1c").split('\n')
        if not (0 <= line_idx < len(all_lines)): return
        tbl_start = line_idx
        while tbl_start > 0 and '|' in all_lines[tbl_start - 1].strip(): tbl_start -= 1
        tbl_end = line_idx
        while tbl_end < len(all_lines) - 1 and '|' in all_lines[tbl_end + 1].strip(): tbl_end += 1
        ox = self._canvas_offset(); m_l = int(self.M_LR * self.PX_PER_MM); c_w = self.A4_W - m_l * 2
        
        cur_parts = all_lines[line_idx].strip().strip('|').split('|')
        nc = len(cur_parts) if cur_parts else 1
        col_w = c_w / nc; rel_x = click_cx - ox - m_l
        col_idx = max(0, min(int(rel_x / col_w), nc - 1))
        insert_pos = col_idx if left else col_idx + 1
        
        for i in range(tbl_start, tbl_end + 1):
            lt = all_lines[i].strip()
            if '|' in lt:
                is_sep = '---' in lt
                parts = lt.strip().strip('|').split('|')
                if is_sep: parts.insert(insert_pos, "---")
                else: parts.insert(insert_pos, " ")
                all_lines[i] = "| " + " | ".join([p.strip() for p in parts]) + " |"
                
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(all_lines).rstrip('\n'))
        self._save_undo_state(); self.refresh_preview()

    def add_custom_section(self):
        self.custom_sections.append((tk.BooleanVar(value=True), tk.StringVar(value="自定义章节"), tk.StringVar(value="Custom Section")))
        self._render_custom_sections(); self._sync_opts_to_text()

    def remove_custom_section(self, index):
        if 0<=index<len(self.custom_sections):
            self.custom_sections.pop(index); self._render_custom_sections()
            self._sync_opts_to_text()

    def clear_custom_sections(self):
        self.custom_sections.clear(); self._render_custom_sections()
        self._sync_opts_to_text()

    def _render_custom_sections(self):
        for w in self.custom_frame.winfo_children(): w.destroy()
        for i,(enabled,cn_var,en_var) in enumerate(self.custom_sections):
            row_f=tk.Frame(self.custom_frame,bg=self.C_BG_HDR,padx=4,pady=2); row_f.pack(fill="x", padx=6, pady=1)
            tk.Checkbutton(row_f,variable=enabled,bg=self.C_BG_HDR,activebackground=self.C_BG_HDR, command=self._on_custom_toggle).pack(side=tk.LEFT)
            tk.Label(row_f,text="中:",bg=self.C_BG_HDR,font=("微软雅黑",8)).pack(side=tk.LEFT)
            e_cn=tk.Entry(row_f,textvariable=cn_var,font=("微软雅黑",8),width=12,bg=self.C_BG_PANEL,relief="flat")
            e_cn.pack(side=tk.LEFT,padx=2)
            tk.Label(row_f,text="EN:",bg=self.C_BG_HDR,font=("微软雅黑",8)).pack(side=tk.LEFT,padx=(4,0))
            e_en=tk.Entry(row_f,textvariable=en_var,font=("微软雅黑",8),width=12,bg=self.C_BG_PANEL,relief="flat")
            e_en.pack(side=tk.LEFT,padx=2)
            tk.Button(row_f,text="✕",command=lambda idx=i:self.remove_custom_section(idx), bg=self.C_ACCENT_RED,fg="white",font=("微软雅黑",8),relief="flat").pack(side=tk.LEFT,padx=(6,0))

    def _on_custom_toggle(self):
        self._sync_opts_to_text()

    # ══════════════════════════════════════════
    # 工具栏
    # ══════════════════════════════════════════
    def _build_toolbar(self, parent):
        TB = self.C_BG_PANEL
        bar_outer=tk.Frame(parent,bg=TB); bar_outer.pack(fill="x",side=tk.TOP)
        
        # 第一行
        bar1=tk.Frame(bar_outer,bg=TB,height=40); bar1.pack(fill="x"); bar1.pack_propagate(False)
        def sep(b): tk.Frame(b,bg=self.C_BORDER,width=1).pack(side=tk.LEFT,fill=tk.Y,padx=6,pady=8)
        def lbl(b,t): tk.Label(b,text=t,bg=TB,fg=self.C_TEXT,font=("微软雅黑",8)).pack(side=tk.LEFT,padx=(4,1))
        def spb(b,var,lo,hi,w=3):
            sp=tk.Spinbox(b,textvariable=var,from_=lo,to=hi,width=w,bg=self.C_BG_MAIN,fg=self.C_TEXT,font=("微软雅黑",9),relief="flat",command=self.refresh_preview)
            sp.pack(side=tk.LEFT,padx=(0,2),pady=6)
            sp.bind("<Return>",lambda e:self.refresh_preview()); sp.bind("<FocusOut>",lambda e:self.refresh_preview())

        lbl(bar1,"中文")
        cb_cn=ttk.Combobox(bar1,textvariable=self.var_cn_font,values=FONT_CHOICES,width=8,state="readonly")
        cb_cn.pack(side=tk.LEFT,padx=(0,3),pady=6); cb_cn.bind("<<ComboboxSelected>>",lambda e:self.refresh_preview()); sep(bar1)
        lbl(bar1,"英文")
        cb_en=ttk.Combobox(bar1,textvariable=self.var_en_font,values=FONT_CHOICES,width=8,state="readonly")
        cb_en.pack(side=tk.LEFT,padx=(0,3),pady=6); cb_en.bind("<<ComboboxSelected>>",lambda e:self.refresh_preview()); sep(bar1)
        lbl(bar1,"标题");spb(bar1,self.var_title_size,8,36); lbl(bar1,"正文");spb(bar1,self.var_body_size,6,24)
        lbl(bar1,"行距");spb(bar1,self.var_line_gap,0,40); lbl(bar1,"标题距");spb(bar1,self.var_title_gap,0,50); sep(bar1)
        lbl(bar1,"符号")
        cb_bul=ttk.Combobox(bar1,textvariable=self.var_bullet,values=list(BULLET_STYLES.keys()),width=10,state="readonly")
        cb_bul.pack(side=tk.LEFT,padx=(0,3),pady=6); cb_bul.bind("<<ComboboxSelected>>",lambda e:self.refresh_preview())
        tk.Checkbutton(bar1,text="自动缩字",variable=self.var_auto_shrink,bg=TB,fg=self.C_TEXT, activebackground=TB,font=("微软雅黑",8),command=self.refresh_preview).pack(side=tk.LEFT,padx=3,pady=6)
        tk.Button(bar1,text="↺ 重置",command=self._reset_style,bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",8,"bold"),relief="flat",padx=5).pack(side=tk.LEFT,padx=4,pady=6)

        # 第二行 (插入与章节)
        bar2=tk.Frame(bar_outer,bg=self.C_BG_MAIN,height=36); bar2.pack(fill="x"); bar2.pack_propagate(False)
        tk.Label(bar2,text="📷 图框:",bg=self.C_BG_MAIN,fg=self.C_TEXT,font=("微软雅黑",8)).pack(side=tk.LEFT,padx=(6,2))
        for n in [1, 2, 3, 4, 6]:
            tk.Button(bar2,text=f" {n}框 ",command=lambda c=n:self.insert_photo_frame(c), bg=self.C_ACCENT_BLUE,fg="white",font=("微软雅黑",8),relief="flat",padx=4).pack(side=tk.LEFT,padx=2,pady=4)
        tk.Button(bar2,text="✕ 清除",command=self.clear_photo_frames,bg=self.C_BG_HDR,fg=self.C_TEXT, font=("微软雅黑",8),relief="flat",padx=5).pack(side=tk.LEFT,padx=4,pady=4)
        sep(bar2)

        tk.Button(bar2,text="📌左浮动",command=lambda:self._set_float_mode('left'),bg=self.C_ACCENT_GREEN,fg="white",font=("微软雅黑",8),relief="flat",padx=4).pack(side=tk.LEFT,padx=2,pady=4)
        tk.Button(bar2,text="📌右浮动",command=lambda:self._set_float_mode('right'),bg=self.C_ACCENT_BLUE,fg="white",font=("微软雅黑",8),relief="flat",padx=4).pack(side=tk.LEFT,padx=2,pady=4)
        self.lbl_float_hint=tk.Label(bar2,text="",bg=self.C_BG_MAIN,fg=self.C_ACCENT_ORANGE,font=("微软雅黑",8)); self.lbl_float_hint.pack(side=tk.LEFT,padx=4)

        sep(bar2)
        tk.Label(bar2, text="➕附加章节:", bg=self.C_BG_MAIN, fg=self.C_TEXT, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        tk.Checkbutton(bar2, text="图片", variable=self.opt_product_image, font=("微软雅黑", 8), bg=self.C_BG_MAIN, activebackground=self.C_BG_MAIN, command=self._on_opt_toggle).pack(side=tk.LEFT, padx=1)
        tk.Checkbutton(bar2, text="包装", variable=self.opt_product_package, font=("微软雅黑", 8), bg=self.C_BG_MAIN, activebackground=self.C_BG_MAIN, command=self._on_opt_toggle).pack(side=tk.LEFT, padx=1)
        tk.Button(bar2, text="+自定义", command=self.add_custom_section, bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=2, pady=6)
        tk.Button(bar2, text="清空", command=self.clear_custom_sections, bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=2, pady=6)

        self.custom_frame = tk.Frame(bar_outer, bg=self.C_BG_MAIN)
        self.custom_frame.pack(fill="x")

        # 第三行 (排版与重写)
        bar3=tk.Frame(bar_outer,bg=self.C_BG_MAIN,height=36); bar3.pack(fill="x"); bar3.pack_propagate(False)
        tk.Button(bar3,text="🔀 行列互换",command=self.transpose_table,bg=self.C_ACCENT_BLUE,fg="white", font=("微软雅黑",8),relief="flat",padx=8).pack(side=tk.LEFT,padx=4,pady=4)

        sep(bar3)
        self.btn_fmt_brush=tk.Button(bar3,text="🖌️ 格式刷",command=self.toggle_fmt_brush, bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",8),relief="flat",padx=8)
        self.btn_fmt_brush.pack(side=tk.LEFT,padx=4,pady=4)
        
        sep(bar3)
        self.btn_undo=tk.Button(bar3,text="↩ 撤销",font=("微软雅黑",8),relief="flat",bg=self.C_BG_HDR,fg=self.C_TEXT,padx=6,command=self.do_undo,state=tk.DISABLED)
        self.btn_undo.pack(side=tk.LEFT,padx=2,pady=4)
        self.btn_redo=tk.Button(bar3,text="↪ 重做",font=("微软雅黑",8),relief="flat",bg=self.C_BG_HDR,fg=self.C_TEXT,padx=6,command=self.do_redo,state=tk.DISABLED)
        self.btn_redo.pack(side=tk.LEFT,padx=2,pady=4)

        sep(bar3)
        tk.Button(bar3,text="➕ 插入正文",command=self.insert_new_text_line,bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",8),relief="flat",padx=8).pack(side=tk.LEFT,padx=4,pady=4)
        
        tk.Label(bar3, text="符号:", bg=self.C_BG_MAIN, fg=self.C_TEXT, font=("微软雅黑", 8)).pack(side=tk.LEFT, padx=(2,0))
        self.var_insert_symbol = tk.StringVar()
        cb_ins_sym = ttk.Combobox(bar3, textvariable=self.var_insert_symbol, values=list(BULLET_STYLES.keys()), width=10, state="readonly")
        cb_ins_sym.pack(side=tk.LEFT, padx=(0,4), pady=4)
        
        tk.Button(bar3, text="🔄 一键去序号", command=self.replace_numbers_with_bullets, bg=self.C_ACCENT_ORANGE, fg="white", font=("微软雅黑", 8), relief="flat", padx=4).pack(side=tk.LEFT, padx=4, pady=4)

        tk.Button(bar3, text="✂️ 一键去说明", command=self.remove_descriptions_after_colon, bg=self.C_ACCENT_RED, fg="white", font=("微软雅黑", 8), relief="flat", padx=4).pack(side=tk.LEFT, padx=4, pady=4)
        
        def on_insert_sym(e):
            sym_key = self.var_insert_symbol.get()
            if not sym_key: return
            sym_val = BULLET_STYLES.get(sym_key, "● ")
            if sym_val in ("__NUM__", "__DOT__"):
                sym_val = "① " if sym_val == "__NUM__" else "1. "
            
            if self._active_editor:
                self._active_editor['entry'].insert(tk.INSERT, sym_val)
            elif self._cursor_raw_text:
                txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
                lines = txt.get("1.0", "end-1c").split('\n')
                for i, l in enumerate(lines):
                    if l.strip() == self._cursor_raw_text.strip() and l.strip():
                        lines[i] = sym_val + l
                        self._cursor_raw_text = lines[i]
                        break
                txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
                self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()
            else:
                messagebox.showinfo("提示", "请先在预览区点击选中一行正文，或双击进入编辑模式后使用。")
            self.var_insert_symbol.set('')
            self.canvas.focus_set()
            
        cb_ins_sym.bind("<<ComboboxSelected>>", on_insert_sym)

        sep(bar3)
        tk.Label(bar3, text="🖼️背景:", bg=self.C_BG_MAIN, fg=self.C_TEXT, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        tk.Button(bar3, text="封面", command=lambda:self.pick_bg('cover'), bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=1, pady=6)
        tk.Button(bar3, text="水印", command=lambda:self.pick_bg('body'), bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=1, pady=6)

        # 第四行 (翻译工具栏与重写)
        bar4=tk.Frame(bar_outer,bg=self.C_BG_MAIN,height=36); bar4.pack(fill="x"); bar4.pack_propagate(False)
        tk.Label(bar4,text="🌍 智能翻译:",bg=self.C_BG_MAIN,fg=self.C_TEXT,font=("微软雅黑",8,"bold")).pack(side=tk.LEFT,padx=(6,2))
        
        self.var_trans_from = tk.StringVar(value="自动检测")
        cb_from = ttk.Combobox(bar4, textvariable=self.var_trans_from, values=list(LANGUAGES.keys()), width=10, state="readonly")
        cb_from.pack(side=tk.LEFT, padx=2, pady=4)
        
        tk.Label(bar4,text="➡",bg=self.C_BG_MAIN,fg=self.C_TEXT,font=("微软雅黑",8)).pack(side=tk.LEFT)
        
        self.var_trans_to = tk.StringVar(value="英语")
        cb_to = ttk.Combobox(bar4, textvariable=self.var_trans_to, values=[k for k in LANGUAGES.keys() if k != "自动检测"], width=10, state="readonly")
        cb_to.pack(side=tk.LEFT, padx=2, pady=4)
        
        self.btn_do_trans = tk.Button(bar4, text="实时翻译预览区", command=self.start_translation, bg=self.C_ACCENT_GREEN, fg="white", font=("微软雅黑", 8), relief="flat", padx=8)
        self.btn_do_trans.pack(side=tk.LEFT, padx=6, pady=4)
        
        self.lbl_trans_status = tk.Label(bar4, text="", bg=self.C_BG_MAIN, fg=self.C_ACCENT_ORANGE, font=("微软雅黑", 8))
        self.lbl_trans_status.pack(side=tk.LEFT, padx=4)

        sep(bar4)
        tk.Label(bar4, text="🔄选段重写:", bg=self.C_BG_MAIN, fg=self.C_TEXT, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        tk.Button(bar4, text="AI重写", command=self.ai_rewrite_selection, bg=self.C_ACCENT_ORANGE, fg="white", font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=1, pady=6)
        tk.Button(bar4, text="精炼", command=lambda:self.ai_rewrite_selection("更精炼简洁"), bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=1, pady=6)
        tk.Button(bar4, text="详细", command=lambda:self.ai_rewrite_selection("更详细丰富"), bg=self.C_BG_HDR, fg=self.C_TEXT, font=("微软雅黑", 8), relief="flat", padx=2).pack(side=tk.LEFT, padx=1, pady=6)
        self.rewrite_hint = tk.Label(bar4, text="", fg=self.C_TEXT_MUTED, bg=self.C_BG_MAIN, font=("微软雅黑", 8))
        self.rewrite_hint.pack(side=tk.LEFT, padx=2)

    def remove_descriptions_after_colon(self):
        """核心：一键删除产品特点等章节中冒号后面的说明文字，支持鼠标选中区域局部处理"""
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        
        try:
            sel_first = txt.index(tk.SEL_FIRST)
            sel_last = txt.index(tk.SEL_LAST)
            start_line = int(sel_first.split('.')[0]) - 1
            end_line = int(sel_last.split('.')[0]) - 1
        except tk.TclError:
            start_line = 0
            end_line = float('inf')

        lines = txt.get("1.0", "end-1c").split('\n')
        changed = False

        for i, l in enumerate(lines):
            if i < start_line or i > end_line: continue
            
            s = l.strip()
            if not s or s.startswith('|') or self.is_header(l): continue
            if self._parse_img_frame_tag(l) or self._parse_float_tag(l): continue
            if 'http://' in l or 'https://' in l: continue

            m = re.search(r'^(.*?)(：|:)', l)
            if m:
                prefix = m.group(1)
                clean_prefix = re.sub(r'^[-*•●■▶◆○①-⑳0-9.、\s]+', '', prefix)
                if len(clean_prefix) <= 30:
                    colon_idx = m.end()
                    new_l = l[:colon_idx] 
                    if new_l != l:
                        lines[i] = new_l
                        changed = True

        if changed:
            txt.delete("1.0", tk.END)
            txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt)
            self._save_undo_state()
            self.refresh_preview()
            messagebox.showinfo("成功", "已一键删除了冒号后面的说明文字！")
        else:
            messagebox.showinfo("提示", "没有找到符合条件的冒号说明文案。\n\n💡 小技巧：如果想更精确，您可以先用鼠标在左侧文本框【选中要处理的几段文字】，然后再点击这个按钮。")

    def replace_numbers_with_bullets(self):
        """核心：一键将开头的数字序号（1. 2. 1、①等）替换为项目符号"""
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        
        sym_key = self.var_insert_symbol.get() or self.var_bullet.get()
        sym_val = BULLET_STYLES.get(sym_key, "● ")
        if sym_val in ("__NUM__", "__DOT__"):
            sym_val = "- " 
            
        changed = False
        for i, l in enumerate(lines):
            s = l.strip()
            if not s: continue
            if self._parse_img_frame_tag(l) or self._parse_float_tag(l): continue
            if s.startswith('|'): continue
            if self.is_header(l): continue
            
            new_l = re.sub(r'^(\d+[\.、]|[①-⑳])\s*', sym_val, l.lstrip())
            if new_l != l.lstrip():
                indent = l[:len(l)-len(l.lstrip())]
                lines[i] = indent + new_l
                changed = True
                
        if changed:
            txt.delete("1.0", tk.END)
            txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
            self._highlight_syntax(txt)
            self._save_undo_state()
            self.refresh_preview()
            messagebox.showinfo("成功", "已将数字序号一键转换为选定的项目符号！")
        else:
            messagebox.showinfo("提示", "未在当前文案中找到需要替换的数字序号。")

    def start_translation(self):
        if not self.baidu_appid or not self.baidu_appkey:
            return messagebox.showerror("配置错误", "请先在右上角【⚙️ 系统设置】中配置百度翻译的 APP ID 和 密钥。")
            
        if hasattr(self, '_translating') and self._translating: return
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        raw_content = txt.get("1.0", "end-1c")
        if not raw_content.strip(): return messagebox.showinfo("提示", "当前预览区没有文本可以翻译")
        
        from_name = self.var_trans_from.get()
        to_name = self.var_trans_to.get()
        if from_name not in LANGUAGES or to_name not in LANGUAGES: return
        if from_name == to_name and from_name != "自动检测": return
        
        from_lang = LANGUAGES[from_name]
        to_lang = LANGUAGES[to_name]
        
        self._translating = True
        self.btn_do_trans.config(state=tk.DISABLED, text="正在后台翻译...", bg="#999")
        self.lbl_trans_status.config(text="正在分析排版元素...", fg=self.C_ACCENT_ORANGE)
        
        def run():
            try:
                lines = raw_content.split('\n')
                to_translate_lines = []
                mapping = {}
                formatting_info = {}
                
                for i, line in enumerate(lines):
                    stripped = line.strip()
                    if not stripped: continue
                    if self._parse_img_frame_tag(stripped) or self._parse_float_tag(stripped): continue
                    if stripped.replace('|', '').replace('-', '').strip() == '': continue
                    
                    is_table = '|' in stripped
                    is_hdr = self.is_header(stripped) 
                    prefix = ""
                    clean_text = stripped

                    if not is_table:
                        clean_text = re.sub(r'^#+\s*', '', clean_text)
                        m_prefix = re.match(r'^((?:[-•●■▶◆○—]|\*(?!\*))\s*|\d+\.\s*|[①-⑳]\s*)(.*)', clean_text)
                        if m_prefix:
                            prefix = m_prefix.group(1)
                            clean_text = m_prefix.group(2).strip()
                        m_bold = re.match(r'^\*\*(.*?)\*\*$', clean_text)
                        if m_bold:
                            clean_text = m_bold.group(1).strip()
                    
                    if not clean_text: continue
                    
                    to_translate_lines.append(clean_text)
                    mapping[len(to_translate_lines)-1] = i
                    needs_bold = True if is_hdr else (m_bold is not None)
                    formatting_info[len(to_translate_lines)-1] = (is_table, prefix, needs_bold)
                
                if not to_translate_lines:
                    self.root.after(0, lambda: self._trans_done(txt, lines, "无可翻译内容"))
                    return
                
                def progress_cb(done, total):
                    self.root.after(0, lambda: self.lbl_trans_status.config(text=f"进度: 已处理 {done}/{total} 段落..."))
                
                translated_segments = BaiduTranslator.translate_paragraphs(to_translate_lines, from_lang, to_lang, self.baidu_appid, self.baidu_appkey, progress_cb)
                
                res_lines = lines.copy()
                for i, trans_text in enumerate(translated_segments):
                    orig_idx = mapping[i]
                    is_table, orig_prefix, needs_bold = formatting_info[i]
                    
                    clean_trans = trans_text.strip()
                    
                    if not is_table:
                        clean_trans = re.sub(r'^\*+\s*', '', clean_trans)
                        clean_trans = re.sub(r'\s*\*+$', '', clean_trans)
                        clean_trans = re.sub(r'^#+\s*', '', clean_trans)
                        
                    if needs_bold:
                        clean_trans = f"**{clean_trans}**"
                    
                    res_lines[orig_idx] = orig_prefix + clean_trans
                
                self.root.after(0, lambda: self._trans_done(txt, res_lines, "✅ 翻译并重新排版完成！"))
            except Exception as e:
                self.root.after(0, lambda: self._trans_done(None, None, f"❌ 失败: {str(e)}"))
        
        threading.Thread(target=run, daemon=True).start()

    def _trans_done(self, txt_widget, res_lines, msg):
        self._translating = False
        self.btn_do_trans.config(state=tk.NORMAL, text="实时翻译预览区", bg=self.C_ACCENT_GREEN)
        self.lbl_trans_status.config(text=msg, fg=self.C_ACCENT_GREEN if "完成" in msg else self.C_ACCENT_RED)
        
        if res_lines and txt_widget:
            self._save_undo_state() 
            txt_widget.delete("1.0", tk.END)
            txt_widget.insert("1.0", '\n'.join(res_lines).rstrip('\n'))
            self._highlight_syntax(txt_widget)
            self.refresh_preview()
            
        if "完成" in msg:
            self.root.after(3000, lambda: self.lbl_trans_status.config(text=""))

    def _reset_style(self):
        self.var_title_size.set(14);self.var_body_size.set(11);self.var_line_gap.set(8);self.var_title_gap.set(14)
        self.var_cn_font.set("微软雅黑");self.var_en_font.set("Arial");self.var_bullet.set("●  实心圆")
        self.var_auto_shrink.set(True);self.refresh_preview()

    def insert_new_text_line(self):
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        target_idx = len(lines)
        
        if hasattr(self, '_cursor_raw_text') and self._cursor_raw_text:
            for i, l in enumerate(lines):
                if l.strip() == self._cursor_raw_text.strip() and l.strip():
                    target_idx = i + 1; break
            
        new_text = "请输入正文..."
        lines.insert(target_idx, new_text)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt)
        self._save_undo_state(); self._cursor_raw_text = new_text; self.refresh_preview()
        self.root.after(150, lambda: self._auto_edit_by_raw(new_text))

    def _auto_edit_by_raw(self, raw_text):
        for item in self._edit_items:
            if item[5].strip() == raw_text.strip():
                bbox = item[2]
                cx = (bbox[0] + bbox[2]) / 2; cy = (bbox[1] + bbox[3]) / 2
                self._on_canvas_double_click_at(cx, cy)
                break

    # ══════════════════════════════════════════
    # ★ 表格行列互换
    # ══════════════════════════════════════════
    def transpose_table(self):
        txt=self.txt_cn if self.current_lang=='cn' else self.txt_en
        lines=txt.get("1.0", "end-1c").split('\n')
        def is_tbl(l): return '|' in l.strip() and l.strip()!=''
        def is_sep(l): return '|' in l.strip() and '-' in l.strip() and not any(c.isalnum() for c in l.strip())
        
        cursor_line = -1
        if hasattr(self, '_cursor_raw_text') and self._cursor_raw_text:
            for i, l in enumerate(lines):
                if l.strip() == self._cursor_raw_text.strip() and l.strip() and is_tbl(l):
                    cursor_line = i; break

        if cursor_line == -1 and hasattr(self, '_right_click_item') and self._right_click_item:
            raw = self._right_click_item[5]
            for i, l in enumerate(lines):
                if l.strip() == raw.strip() and l.strip() and is_tbl(l):
                    cursor_line = i; break

        if cursor_line == -1: cursor_line=int(txt.index(tk.INSERT).split('.')[0])-1
        if cursor_line<0 or cursor_line>=len(lines) or not is_tbl(lines[cursor_line]):
            return messagebox.showinfo("提示","请先在预览区点击选中一个表格！")
            
        s=cursor_line
        while s>0 and is_tbl(lines[s-1]): s-=1
        e=cursor_line
        while e<len(lines)-1 and is_tbl(lines[e+1]): e+=1
        
        data=[]
        for i in range(s,e+1):
            if is_sep(lines[i]): continue  
            parts = lines[i].strip().strip('|').split('|')
            cells = [c.strip() for c in parts]
            if cells and (len(cells) > 1 or cells[0] != ''): data.append(cells)
            
        if len(data)<1: return messagebox.showinfo("提示","未找到有效表格数据")
        mc=max(len(r) for r in data)
        for r in data:
            while len(r)<mc: r.append("")
            
        tr=[[data[r][c] for r in range(len(data))] for c in range(mc)]
        nl=[]
        for ri,row in enumerate(tr):
            clean_row = [c.strip() for c in row]
            nl.append("| "+" | ".join(clean_row)+" |")
            if ri==0: nl.append("| "+" | ".join(["---"]*len(clean_row))+" |")
                
        result=lines[:s]+nl+lines[e+1:]
        txt.delete("1.0",tk.END); txt.insert("1.0",'\n'.join(result).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state(); self.refresh_preview()

    # ══════════════════════════════════════════
    # ★ 底部面板（文档提取与网络搜索）
    # ══════════════════════════════════════════
    def _build_bottom_panel(self, parent):
        PBGD = "#202124"; HBGD = "#2D2E32"
        self.btm_outer = tk.Frame(parent, bg=HBGD)
        self.btm_outer.pack(side=tk.BOTTOM, fill=tk.X)

        hdr = tk.Frame(self.btm_outer, bg=HBGD, height=34)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        self.tab_doc_btn = tk.Button(hdr, text="📂 文档图片提取", bg=self.C_ACCENT_BLUE, fg="#fff", font=("微软雅黑", 9, "bold"), relief="flat", padx=12, command=lambda: self._switch_bottom_tab('doc'))
        self.tab_doc_btn.pack(side=tk.LEFT, padx=(8,2), pady=4)
        self.tab_web_btn = tk.Button(hdr, text="🌐 网络图片搜索", bg="#444", fg="#aaa", font=("微软雅黑", 9), relief="flat", padx=12, command=lambda: self._switch_bottom_tab('web'))
        self.tab_web_btn.pack(side=tk.LEFT, padx=2, pady=4)

        self.doc_fold_btn = tk.Button(hdr, text="▲ 收起底部栏", bg=HBGD, fg="#aaa", font=("微软雅黑", 8), relief="flat", cursor="hand2", command=self._toggle_bottom_panel)
        self.doc_fold_btn.pack(side=tk.RIGHT, padx=8)

        self.doc_progress_var = tk.IntVar(value=0)
        self.doc_progress = ttk.Progressbar(hdr, variable=self.doc_progress_var, mode="determinate", length=160, maximum=100)
        self.doc_progress.pack(side=tk.RIGHT, padx=6, pady=6)
        self.doc_status_lbl = tk.Label(hdr, text="", bg=HBGD, fg=self.C_ACCENT_ORANGE, font=("微软雅黑", 8))
        self.doc_status_lbl.pack(side=tk.RIGHT, padx=4)

        self.btm_body = tk.Frame(self.btm_outer, bg=PBGD, height=self.layout_config["bottom_h"])
        self.btm_body.pack(fill="x"); self.btm_body.pack_propagate(False)

        self.doc_extract_body = tk.Frame(self.btm_body, bg=PBGD)
        ctrl_col = tk.Frame(self.doc_extract_body, bg=PBGD, width=130)
        ctrl_col.pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=6); ctrl_col.pack_propagate(False)
        self.doc_extract_files_lbl = tk.Label(ctrl_col, text="未选择文件", bg=PBGD, fg="#7ec8e3", font=("微软雅黑", 8), wraplength=120)
        self.doc_extract_files_lbl.pack(pady=(0, 4))
        tk.Button(ctrl_col, text="🗑 清空", command=self._doc_clear_results, bg="#6b7280", fg="white", font=("微软雅黑", 8), relief="flat", cursor="hand2").pack(fill="x", pady=(0,4))
        tk.Label(ctrl_col, text="上传资料后自动提取图片，点击加入图库", bg=PBGD, fg="#888", font=("微软雅黑", 7), wraplength=118, justify="left").pack(anchor="w", pady=(4,0))

        preview_frame = tk.Frame(self.doc_extract_body, bg=PBGD)
        preview_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0,6), pady=6)
        h_scr = tk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        h_scr.pack(side=tk.BOTTOM, fill=tk.X)
        self.doc_img_canvas = tk.Canvas(preview_frame, bg="#111122", highlightthickness=0, xscrollcommand=h_scr.set, height=self.layout_config["bottom_h"] - 20)
        self.doc_img_canvas.pack(fill="both", expand=True)
        h_scr.config(command=self.doc_img_canvas.xview)
        self.doc_img_inner = tk.Frame(self.doc_img_canvas, bg="#111122")
        self.doc_img_canvas.create_window((0,0), window=self.doc_img_inner, anchor="nw")
        self.doc_img_inner.bind("<Configure>", lambda e: self.doc_img_canvas.config(scrollregion=self.doc_img_canvas.bbox("all")))
        self.doc_img_canvas.bind("<MouseWheel>", lambda e: self.doc_img_canvas.xview_scroll(int(-1*(e.delta/120)), "units"))
        self._doc_show_empty_hint()

        self.web_search_body = tk.Frame(self.btm_body, bg=PBGD)
        sf=tk.Frame(self.web_search_body,bg=PBGD,width=220); sf.pack(side=tk.LEFT,fill=tk.Y,padx=6,pady=6); sf.pack_propagate(False)
        self.wsv=tk.StringVar(); se=tk.Entry(sf,textvariable=self.wsv,font=("微软雅黑",10),width=18); se.pack(fill="x",pady=(0,6)); se.bind("<Return>",lambda e:self._web_search(True))
        tk.Button(sf,text="🔍 搜索",command=lambda:self._web_search(True),bg=self.C_ACCENT_BLUE,fg="white",font=("微软雅黑",9),relief="flat").pack(fill="x",pady=(0,4))
        tk.Button(sf,text="▶ 下一批",command=lambda:self._web_search(False),bg=self.C_ACCENT_ORANGE,fg="white",font=("微软雅黑",9),relief="flat").pack(fill="x")
        self.wsl=tk.Label(sf,text="输入关键词搜索",bg=PBGD,fg="#888",font=("微软雅黑",7)); self.wsl.pack(anchor="w",pady=4)
        
        wpf=tk.Frame(self.web_search_body,bg=PBGD); wpf.pack(side=tk.LEFT,fill="both",expand=True,padx=(0,6),pady=6)
        whs=tk.Scrollbar(wpf,orient=tk.HORIZONTAL); whs.pack(side=tk.BOTTOM,fill=tk.X)
        self.wicv=tk.Canvas(wpf,bg="#111122",highlightthickness=0,xscrollcommand=whs.set, height=self.layout_config["bottom_h"] - 20); self.wicv.pack(fill="both",expand=True); whs.config(command=self.wicv.xview)
        self.wii=tk.Frame(self.wicv,bg="#111122"); self.wicv.create_window((0,0),window=self.wii,anchor="nw")
        self.wii.bind("<Configure>",lambda e:self.wicv.config(scrollregion=self.wicv.bbox("all")))
        self.wicv.bind("<MouseWheel>", lambda e: self.wicv.xview_scroll(int(-1*(e.delta/120)), "units"))
        self._switch_bottom_tab('doc')

    def _switch_bottom_tab(self, tab):
        self.doc_extract_body.pack_forget(); self.web_search_body.pack_forget()
        self.tab_doc_btn.config(bg="#444", fg="#aaa", font=("微软雅黑", 9, "bold"))
        self.tab_web_btn.config(bg="#444", fg="#aaa", font=("微软雅黑", 9, "bold"))
        if tab == 'doc':
            self.tab_doc_btn.config(bg=self.C_ACCENT_BLUE, fg="#fff", font=("微软雅黑", 9, "bold"))
            self.doc_extract_body.pack(fill="both", expand=True)
        else:
            self.tab_web_btn.config(bg=self.C_ACCENT_BLUE, fg="#fff", font=("微软雅黑", 9, "bold"))
            self.web_search_body.pack(fill="both", expand=True)

    def _toggle_bottom_panel(self):
        if self._bottom_panel_expanded:
            self.btm_body.pack_forget()
            self.doc_fold_btn.config(text="▼ 展开底部栏")
            self._bottom_panel_expanded = False
            if not self.right_panel_visible: self.toggle_right_panel()
        else:
            self.btm_body.pack(fill="x")
            self.doc_fold_btn.config(text="▲ 收起底部栏")
            self._bottom_panel_expanded = True

    # ══════════════════════════════════════════
    # ★ 文档图片提取逻辑
    # ══════════════════════════════════════════
    def _doc_show_empty_hint(self):
        for w in self.doc_img_inner.winfo_children(): w.destroy()
        self._doc_extract_thumbs.clear()
        tk.Label(self.doc_img_inner, text="← 选择文档上传后 → 提取的图片预览于此 → 点击加入图片库", bg="#111122", fg="#444", font=("微软雅黑", 9), justify="left").pack(padx=20, pady=40)

    def _doc_clear_results(self):
        self._doc_extract_imgs.clear(); self._doc_show_empty_hint()
        self.doc_status_lbl.config(text=""); self.doc_progress_var.set(0)

    def _doc_do_extract_and_done(self):
        self.root.after(0, lambda: (
            self.doc_progress.config(mode="indeterminate"),
            self.doc_progress.start(12),
            self.doc_status_lbl.config(text="提取图片中...", fg=self.C_ACCENT_ORANGE)
        ))
        tmp_dir = tempfile.mkdtemp(prefix="huamai_docimg_")
        total = 0
        try:
            for file_path in self._doc_extract_files:
                ext = os.path.splitext(file_path)[1].lower()
                stem = os.path.splitext(os.path.basename(file_path))[0]
                safe_stem = re.sub(r'[\\/:*?"<>|]', '_', stem)
                if ext == ".pdf": total += self._doc_extract_from_pdf(file_path, tmp_dir, safe_stem)
                elif ext in (".docx", ".doc"): total += self._doc_extract_from_docx(file_path, tmp_dir, safe_stem)
        except Exception as e: self.root.after(0, lambda: messagebox.showerror("提取失败", str(e)))
        self.root.after(0, lambda: self._doc_extract_done(total))

    def _doc_extract_from_pdf(self, pdf_path, out_dir, safe_stem):
        count = 0
        try:
            import fitz
            doc = fitz.open(pdf_path)
            seen = set()
            for page_num in range(len(doc)):
                for img in doc[page_num].get_images(full=True):
                    xref = img[0]
                    if xref in seen: continue
                    seen.add(xref)
                    try:
                        bi = doc.extract_image(xref)
                        data = bi.get("image", b"")
                        ext  = bi.get("ext", "png")
                        if len(data) < 5000: continue
                        fname = f"{safe_stem}_P{page_num+1}_IMG{count+1}.{ext}"
                        fpath = os.path.join(out_dir, fname)
                        with open(fpath, "wb") as f: f.write(data)
                        self._doc_extract_imgs.append(fpath)
                        count += 1
                    except: pass
            doc.close()
        except ImportError: pass
        return count

    def _doc_extract_from_docx(self, docx_path, out_dir, safe_stem):
        import zipfile
        count = 0
        valid = {".png",".jpg",".jpeg",".gif",".bmp",".tiff",".tif",".webp"}
        try:
            with zipfile.ZipFile(docx_path, "r") as z:
                media = [f for f in z.namelist() if f.startswith("word/media/")]
                for mf in media:
                    suf = os.path.splitext(mf)[1].lower()
                    if suf not in valid: continue
                    data = z.read(mf)
                    if len(data) < 5000: continue
                    orig = os.path.basename(mf)
                    fname = f"{safe_stem}_{orig}"
                    fpath = os.path.join(out_dir, fname)
                    with open(fpath, "wb") as f: f.write(data)
                    self._doc_extract_imgs.append(fpath)
                    count += 1
        except: pass
        return count

    def _doc_extract_done(self, total):
        self.doc_progress.stop(); self.doc_progress.config(mode="determinate")
        self.doc_progress_var.set(100); self._doc_extract_running = False
        self.doc_status_lbl.config(text=f"✅ 共提取 {total} 张" if total > 0 else "⚠ 未提取到", fg=self.C_ACCENT_GREEN if total > 0 else self.C_ACCENT_ORANGE)
        if total > 0: self._doc_render_thumbs()
        else: self._doc_show_empty_hint()

    def _doc_render_thumbs(self):
        for w in self.doc_img_inner.winfo_children(): w.destroy()
        self._doc_extract_thumbs.clear()
        TH = 75 
        for i, path in enumerate(self._doc_extract_imgs):
            try:
                img = Image.open(path).convert("RGB"); w_orig, h_orig = img.size
                scale = TH / h_orig; new_w = max(1, int(w_orig * scale))
                img = img.resize((new_w, TH), Image.LANCZOS)
                ti = ImageTk.PhotoImage(img)
            except: ti = ImageTk.PhotoImage(Image.new("RGB", (TH, TH), "#333"))
            self._doc_extract_thumbs.append(ti)
            card = tk.Frame(self.doc_img_inner, bg="#1e2040", padx=3, pady=3, cursor="hand2")
            card.pack(side=tk.LEFT, padx=4, pady=4)
            lbl = tk.Label(card, image=ti, bg="#1e2040", cursor="hand2", relief="flat", borderwidth=2); lbl.pack()
            fname = os.path.basename(path)
            tk.Label(card, text=fname[:14]+"…" if len(fname)>15 else fname, bg="#1e2040", fg="#aaa", font=("微软雅黑", 7), wraplength=TH+20).pack()
            def on_click(event, p=path, c=card): self._doc_add_to_gallery(p, c)
            card.bind("<Button-1>", on_click); lbl.bind("<Button-1>",  on_click)

    def _doc_add_to_gallery(self, path, card=None):
        if path not in self.gallery_paths: self.gallery_paths.append(path)
        idx = self.gallery_paths.index(path)
        self.gallery_sel.add(idx)
        self._render_gallery()
        if card:
            try:
                card.config(bg=self.C_ACCENT_GREEN)
                for w in card.winfo_children(): w.config(bg=self.C_ACCENT_GREEN)
                card.after(400, lambda: (card.config(bg="#1e2040"), [w.config(bg="#1e2040") for w in card.winfo_children()]))
            except: pass

    # ══════════════════════════════════════════
    # ★ 网络搜索逻辑
    # ══════════════════════════════════════════
    def _web_search(self, new):
        kw = self.wsv.get().strip()
        if not kw: return messagebox.showinfo("提示", "请输入关键词")
        if self._web_searching: return
        if new: self._web_search_keyword = kw; self._web_search_page = 0
        else: self._web_search_page += 1
        self._web_searching = True
        self.wsl.config(text="🔍 搜索中...", fg=self.C_ACCENT_ORANGE)
        threading.Thread(target=self._web_fetch_thread, daemon=True).start()

    def _web_fetch_thread(self):
        kw = self._web_search_keyword
        results = baidu_image_search(kw, pn=self._web_search_page * 20, rn=20)
        if not results:
            self.root.after(0, lambda: self.wsl.config(text="⚠ 未搜到图片，请重试", fg=self.C_ACCENT_ORANGE))
            self._web_searching = False; return
            
        td = tempfile.mkdtemp(prefix="hm_img_")
        dl = []; headers = {'User-Agent': WebFetcher.UA, 'Referer': 'https://image.baidu.com/'}
        for i, item in enumerate(results):
            try:
                r = requests.get(item['thumb'], headers=headers, timeout=12, verify=False)
                if r.status_code == 200 and len(r.content) > 2000:
                    fp = os.path.join(td, f"img_{self._web_search_page}_{i}.jpg")
                    with open(fp, 'wb') as f: f.write(r.content)
                    dl.append({'path': fp, 'desc': item.get('desc', kw)[:20]})
            except: pass
        self.root.after(0, lambda: self._web_fetch_done(dl))

    def _web_fetch_done(self, dl):
        self._web_searching = False
        if not dl: return self.wsl.config(text="⚠ 下载失败", fg=self.C_ACCENT_ORANGE)
        self.wsl.config(text=f"✅ {len(dl)}张（点击加入图库）", fg=self.C_ACCENT_GREEN)
        for w in self.wii.winfo_children(): w.destroy()
        self._web_search_thumbs = []; TH = 75
        for item in dl:
            try:
                img = Image.open(item['path']).convert("RGB"); w0, h0 = img.size
                img = img.resize((max(1, int(w0 * TH / h0)), TH), Image.LANCZOS)
                ti = ImageTk.PhotoImage(img)
            except: ti = ImageTk.PhotoImage(Image.new("RGB", (TH, TH), "#333"))
            self._web_search_thumbs.append(ti)
            card = tk.Frame(self.wii, bg="#1e2040", padx=3, pady=3, cursor="hand2")
            card.pack(side=tk.LEFT, padx=4, pady=4)
            lbl = tk.Label(card, image=ti, bg="#1e2040", cursor="hand2"); lbl.pack()
            tk.Label(card, text=item.get('desc','')[:14], bg="#1e2040", fg="#aaa", font=("微软雅黑", 7)).pack()
            def oc(e, p=item['path'], c=card): self._doc_add_to_gallery(p, c)
            card.bind("<Button-1>", oc); lbl.bind("<Button-1>", oc)

    # ══════════════════════════════════════════
    # 图片库 
    # ══════════════════════════════════════════
    def _build_gallery_panel(self):
        PBG=self.C_BG_PANEL; HBG=self.C_BG_HDR
        self.gallery_panel=tk.Frame(self.main_paned,width=self.layout_config["gallery_w"],bg=PBG)
        self.gallery_panel.pack_propagate(False)
        hdr=tk.Frame(self.gallery_panel,bg=HBG,height=36); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="📷 图库",bg=HBG,fg=self.C_TEXT,font=("微软雅黑",10,"bold")).pack(side=tk.LEFT,padx=8,pady=6)
        f_btn=tk.Frame(self.gallery_panel,bg=PBG); f_btn.pack(fill="x",padx=4,pady=4)
        tk.Button(f_btn,text="＋ 上传",command=self.gallery_upload,bg=self.C_ACCENT_GREEN,fg="white", font=("微软雅黑",8),relief="flat").pack(fill="x",pady=(0,2))
        tk.Button(f_btn,text="🗑 删除",command=self.gallery_delete_selected,bg=self.C_ACCENT_RED,fg="white", font=("微软雅黑",8),relief="flat").pack(fill="x",pady=(0,2))
        tk.Button(f_btn,text="☐ 全选",command=self.gallery_toggle_all,bg=self.C_BORDER,fg=self.C_TEXT, font=("微软雅黑",8),relief="flat").pack(fill="x")
        self.lbl_gallery_tip=tk.Label(self.gallery_panel,text="", bg=PBG,fg=self.C_TEXT_MUTED,font=("微软雅黑",8),justify="center",wraplength=self.layout_config["gallery_w"]-8)
        self.lbl_gallery_tip.pack(pady=2)
        sa=tk.Frame(self.gallery_panel,bg=PBG); sa.pack(fill="both",expand=True)
        self.gal_vscr=tk.Scrollbar(sa,orient=tk.VERTICAL); self.gal_vscr.pack(side=tk.RIGHT,fill=tk.Y)
        self.gal_canvas=tk.Canvas(sa,bg=PBG,highlightthickness=0,yscrollcommand=self.gal_vscr.set)
        self.gal_canvas.pack(fill="both",expand=True); self.gal_vscr.config(command=self.gal_canvas.yview)
        self.gal_inner=tk.Frame(self.gal_canvas,bg=PBG)
        self.gal_canvas.create_window((0,0),window=self.gal_inner,anchor="nw")
        self.gal_inner.bind("<Configure>",lambda e:self.gal_canvas.config(scrollregion=self.gal_canvas.bbox("all")))
        self.gal_canvas.bind("<MouseWheel>",lambda e:self.gal_canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

    def gallery_upload(self):
        paths=filedialog.askopenfilenames(title="选择图片",filetypes=[("图片","*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tiff")])
        for p in paths:
            if p not in self.gallery_paths: self.gallery_paths.append(p)
        self._render_gallery()
    def gallery_delete_selected(self):
        if not self.gallery_sel: return
        for i in sorted(self.gallery_sel,reverse=True):
            if 0<=i<len(self.gallery_paths): self.gallery_paths.pop(i)
        self.gallery_sel.clear(); self._render_gallery(); self.refresh_preview()
    def gallery_toggle_all(self):
        if len(self.gallery_sel)==len(self.gallery_paths): self.gallery_sel.clear()
        else: self.gallery_sel=set(range(len(self.gallery_paths)))
        self._render_gallery()
    def _render_gallery(self):
        for w in self.gal_inner.winfo_children(): w.destroy()
        self.gallery_thumbs=[]
        COLS=1; cw=(self.layout_config["gallery_w"]-10) 
        for i,path in enumerate(self.gallery_paths):
            row,col=i//COLS,i%COLS; sel=i in self.gallery_sel
            try:
                img=Image.open(path).convert("RGB"); img.thumbnail((cw-4,cw-4),Image.LANCZOS)
                if sel:
                    b=Image.new("RGB",(cw,cw),self.C_ACCENT_BLUE); bw,bh=img.size
                    b.paste(img,((cw-bw)//2,(cw-bh)//2)); img=b
                ti=ImageTk.PhotoImage(img)
            except: ti=ImageTk.PhotoImage(Image.new("RGB",(cw-4,cw-4),"#444"))
            self.gallery_thumbs.append(ti)
            cf=tk.Frame(self.gal_inner,bg=self.C_ACCENT_BLUE if sel else self.C_BG_HDR,padx=2,pady=2); cf.grid(row=row,column=col,padx=2,pady=2)
            li=tk.Label(cf,image=ti,bg=self.C_ACCENT_BLUE if sel else self.C_BG_HDR,cursor="hand2"); li.pack()
            def oc(event,idx=i): self.gallery_sel.symmetric_difference_update({idx}); self._render_gallery()
            li.bind("<Button-1>",oc); cf.bind("<Button-1>",oc)
        n,ns=len(self.gallery_paths),len(self.gallery_sel)
        self.lbl_gallery_tip.config(text=f"选{ns}张" if n else "")
        self.gal_canvas.config(scrollregion=self.gal_canvas.bbox("all"))
    def get_selected_paths(self):
        return [self.gallery_paths[i] for i in sorted(self.gallery_sel) if i<len(self.gallery_paths)]

    # ══════════════════════════════════════════
    # 图片框
    # ══════════════════════════════════════════
    def _build_img_frame_tag(self,cols,img_paths, scale=100):
        base = f"[IMG_FRAME:{cols}:{'|'.join(p if p else '' for p in img_paths)}]"
        if scale != 100:
            base = base[:-1] + f"|S:{scale}]"
        return base
        
    def _parse_img_frame_tag(self,line):
        m=IMG_FRAME_PATTERN.match(line.strip())
        if m:
            cols=int(m.group(1)); ps=m.group(2)
            scale = int(m.group(3)) if m.group(3) else 100
            return cols,[p if p else None for p in ps.split("|")] if ps else [], scale
        return None
        
    def _parse_float_tag(self, line):
        m = FLOAT_IMG_PATTERN.match(line.strip())
        return (m.group(1), int(m.group(2)), m.group(3)) if m else None

    def _count_img_frames_in_text(self,lang=None):
        if lang is None: lang=self.current_lang
        txt=self.txt_cn if lang=='cn' else self.txt_en
        return sum(1 for l in txt.get("1.0", "end-1c").split('\n') if self._parse_img_frame_tag(l))

    def insert_photo_frame(self, cols):
        sel = self.get_selected_paths()
        if not sel: return messagebox.showinfo("提示", "请先在图片库中选择图片")
        imgs = (sel + [None] * cols)[:cols]
        tag = self._build_img_frame_tag(cols, imgs)
        txt = self.txt_cn if self.current_lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        target_idx = len(lines)
        if hasattr(self, '_cursor_raw_text') and self._cursor_raw_text:
            for i, l in enumerate(lines):
                if l.strip() == self._cursor_raw_text.strip() and l.strip():
                    target_idx = i + 1; break
        lines.insert(target_idx, tag)
        txt.delete("1.0", tk.END); txt.insert("1.0", '\n'.join(lines).rstrip('\n'))
        self._highlight_syntax(txt); self._save_undo_state(); self._cursor_raw_text = tag; self.refresh_preview()

    def clear_photo_frames(self):
        for lk in ['cn','en']:
            txt=self.txt_cn if lk=='cn' else self.txt_en
            lines=txt.get("1.0", "end-1c").split('\n')
            txt.delete("1.0",tk.END); txt.insert("1.0",'\n'.join(l for l in lines if self._parse_img_frame_tag(l) is None).strip())
        self._save_undo_state(); self.refresh_preview()

    # ══════════════════════════════════════════
    # ★ AI局部重写 
    # ══════════════════════════════════════════
    def ai_rewrite_selection(self,style_hint=None):
        if not self.kimi_api_key:
            return messagebox.showerror("配置错误", "请先在右上角【⚙️ 系统设置】中配置 Kimi API Key。")
            
        txt=self.txt_cn if self.current_lang=='cn' else self.txt_en
        try: sel_text=txt.get(tk.SEL_FIRST,tk.SEL_LAST).strip()
        except tk.TclError: return messagebox.showinfo("提示","请先选中要重写的文字")
        if not sel_text: return
        self.rewrite_hint.config(text="🔄 AI重写中...",fg=self.C_ACCENT_ORANGE)
        ss=txt.index(tk.SEL_FIRST); se=txt.index(tk.SEL_LAST); lang=self.current_lang
        def run():
            h=style_hint or ""
            if style_hint and "详细" in style_hint:
                DESC_KEYS_CN  = ["产品描述","产品介绍","产品指标"]
                DESC_KEYS_EN  = ["Product Description","Product Specifications","Technical Specifications"]
                OTHER_KEYS_CN = ["产品特点","应用场景","安装方式","安装方法","使用说明","注意事项","技术参数"]
                OTHER_KEYS_EN = ["Product Features","Applications","Application Scenarios","Installation","Instructions","Notes"]
                ALL_KEYS = DESC_KEYS_CN + DESC_KEYS_EN + OTHER_KEYS_CN + OTHER_KEYS_EN
                all_lines = txt.get("1.0", "end-1c").split('\n')
                sel_start_line = int(txt.index(tk.SEL_FIRST).split('.')[0])
                nearest_header = ""
                for i in range(min(sel_start_line - 1, len(all_lines) - 1), -1, -1):
                    line_clean = re.sub(r'\*\*(.+?)\*\*', r'\1', all_lines[i]).strip()
                    if line_clean in ALL_KEYS:
                        nearest_header = line_clean; break
                orig_len  = len(sel_text)
                add_chars = 50 if nearest_header in DESC_KEYS_CN + DESC_KEYS_EN else 20
                target_len = orig_len + add_chars
                
                # 修复换行问题：强制去除回车换行指令
                if lang == 'cn':
                    p = (f"请在原文基础上扩写，补充细节，输出字数控制在{target_len}字左右。绝对不要换行，必须输出为单行纯文本，不要带有任何多余的项目符号，直接输出扩写后的内容：\n\n{sel_text}")
                else:
                    p = (f"Expand text adding ~{add_chars} characters. Target: ~{target_len} chars. DO NOT use line breaks. Output as a single line of plain text without bullet points. Output only:\n\n{sel_text}")
            else:
                p=(f"请重写以下片段，保持专业性和格式。{h}\n直接输出：\n\n{sel_text}" if lang=='cn' else f"Rewrite professionally. {h}\nOutput only:\n\n{sel_text}")
            try:
                res=requests.post(KIMI_API_URL,headers={"Authorization":f"Bearer {self.kimi_api_key}"},
                    json={"model":"moonshot-v1-8k","messages":[{"role":"user","content":p}],"temperature":0.3})
                rj=res.json()
                if "choices" in rj:
                    nt=rj["choices"][0]["message"]["content"].strip()
                    self.root.after(0,lambda:self._apply_rewrite(txt,ss,se,nt))
                else: self.root.after(0,lambda:messagebox.showerror("AI错误",rj.get("error",{}).get("message","")))
            except Exception as e: self.root.after(0,lambda:messagebox.showerror("网络异常",str(e)))
            self.root.after(0,lambda:self.rewrite_hint.config(text=""))
        threading.Thread(target=run,daemon=True).start()

    def _apply_rewrite(self,txt,s,e,nt):
        # 强制替换换行符，保证项目符号不断行
        nt = nt.replace('\n', '').replace('\r', '').strip()
        txt.delete(s,e); txt.insert(s,nt); txt.tag_add(tk.SEL,s,f"{s}+{len(nt)}c")
        self._highlight_syntax(txt)
        self._save_undo_state(); self.refresh_preview()
        self.rewrite_hint.config(text="✅ 完成！",fg=self.C_ACCENT_GREEN)

    # ══════════════════════════════════════════
    # 预览区编辑
    # ══════════════════════════════════════════
    def _on_canvas_click(self,event):
        cx,cy=self.canvas.canvasx(event.x),self.canvas.canvasy(event.y)
        hit=self._find_edit_item(cx,cy)
        if hit and hit[4]=='opt_section': return
        self._dismiss_editor()
        if hit:
            _,li,bbox,_,_,raw = hit[:6]
            self._cursor_raw_text=raw
            self.canvas.delete("cursor_highlight"); x0,y0,x1,y1=bbox
            self.canvas.create_rectangle(x0-2,y0-1,x1+2,y1+1,outline=self.C_ACCENT_BLUE,width=2,fill="",tags="cursor_highlight")
        else:
            self._cursor_raw_text=None

    def _on_canvas_double_click(self,event):
        self._dismiss_editor()
        cx,cy=self.canvas.canvasx(event.x),self.canvas.canvasy(event.y)
        self._on_canvas_double_click_at(cx, cy)

    def _get_edit_font(self,lt):
        face=self.var_cn_font.get() if self.current_lang=='cn' else self.var_en_font.get()
        if lt=='header': return (face,self.var_title_size.get(),"bold")
        elif lt=='cover': return (face,max(self.var_title_size.get()+14,24),"bold")
        return (face,self.var_body_size.get())

    def _find_edit_item(self,cx,cy):
        for item in self._edit_items:
            _,_,bbox,_,_,_ = item[:6]
            if bbox[0]<=cx<=bbox[2] and bbox[1]<=cy<=bbox[3]: return item
        return None

    def _commit_edit(self,event=None):
        if not self._active_editor: return
        e=self._active_editor; nt=e['entry'].get().strip()
        li,lt,orig,raw=e['line_idx'],e['line_type'],e['original_text'],e['raw_text']
        if nt and nt!=orig:
            txt=self.txt_cn if self.current_lang=='cn' else self.txt_en
            lines=txt.get("1.0", "end-1c").split('\n')
            target_idx = -1
            for i, l in enumerate(lines):
                if l.strip() == raw.strip() and l.strip():
                    target_idx = i; break
            if target_idx != -1:
                old=lines[target_idx]
                if lt=='header': nl=f"**{nt}**"
                elif lt=='bullet':
                    pf=""; s=old.strip()
                    if s.startswith('-'): pf="- "
                    elif s.startswith('*'): pf="* "
                    elif s.startswith('•'): pf="• "
                    t2=nt
                    for sym in BULLET_STYLES.values():
                        if sym not in ("__NUM__","__DOT__") and t2.startswith(sym): t2=t2[len(sym):]; break
                    nl=pf+t2
                elif lt=='cover': nl=f"**{nt}**" if '**' in old else nt
                elif lt=='table_cell':
                    col_idx = e.get('col_idx')
                    parts = old.strip().strip('|').split('|')
                    if col_idx is not None and col_idx < len(parts):
                        parts[col_idx] = " " + nt + " "
                    nl = "| " + " | ".join(p.strip() for p in parts) + " |"
                else: nl=nt
                lines[target_idx]=nl; txt.delete("1.0",tk.END); txt.insert("1.0",'\n'.join(lines).rstrip('\n'))
                self._highlight_syntax(txt)
                self._save_undo_state()
        self._dismiss_editor(); self.refresh_preview()

    def _dismiss_editor(self):
        if self._active_editor:
            try: self.canvas.delete(self._active_editor['win_id']); self._active_editor['entry'].destroy()
            except: pass
            self._active_editor=None
        self.canvas.delete("cursor_highlight")

    def start_ai_cn(self):
        if not self.kimi_api_key:
            return messagebox.showerror("配置错误", "请先在右上角【⚙️ 系统设置】中配置 Kimi API Key。")
        if not self.raw_text and not self.var_custom_prompt.get():
            return messagebox.showwarning("提示","请先提取资料内容或启用自定义提示词")
        self.lbl_info.config(text="🚀 AI撰写中...",fg=self.C_ACCENT_ORANGE)
        threading.Thread(target=self._run_kimi_cn,daemon=True).start()
    def _run_kimi_cn(self):
        prompt = self._build_prompt('cn'); self._call_kimi(prompt, 'cn')
    def start_ai_en(self):
        if not self.kimi_api_key:
            return messagebox.showerror("配置错误", "请先在右上角【⚙️ 系统设置】中配置 Kimi API Key。")
        if not self.raw_text and not self.var_custom_prompt.get():
            return messagebox.showwarning("Tips","Please upload source material first or enable custom prompt.")
        self.lbl_info.config(text="🌐 AI is writing...",fg=self.C_ACCENT_BLUE)
        threading.Thread(target=self._run_kimi_en,daemon=True).start()
    def _run_kimi_en(self):
        prompt = self._build_prompt('en'); self._call_kimi(prompt, 'en')
        
    def _call_kimi(self,prompt,target):
        try:
            res=requests.post(KIMI_API_URL,headers={"Authorization":f"Bearer {self.kimi_api_key}"},
                json={"model":"moonshot-v1-8k","messages":[{"role":"user","content":prompt}],"temperature":0.2})
            rj=res.json()
            if "choices" in rj:
                text = rj["choices"][0]["message"]["content"]
                text = self._sanitize_ai_output(text)
                self.root.after(0,lambda:self._apply_ai(text,target))
            else: self.root.after(0,lambda:messagebox.showerror("AI错误",rj.get("error",{}).get("message","")))
        except Exception as e: self.root.after(0,lambda:messagebox.showerror("网络异常",str(e)))

    def _sanitize_ai_output(self, text):
        ls = text.split('\n'); r = []; found = False
        for l in ls:
            s = l.strip()
            if not found:
                if not s: continue
                if self.is_header(s): found=True; r.append(l)
            else: r.append(l)
        if not self.var_feature_brief.get():
            final_r = []; in_feat = False
            for line in r:
                clean_line = self.clean_markdown(line)
                if self.is_header(line):
                    in_feat = ("特点" in clean_line or "Features" in clean_line)
                    final_r.append(line)
                elif in_feat and line.strip().startswith(('-','*','•')):
                    if '：' in line: line = line.split('：')[0]
                    elif ':' in line: line = line.split(':')[0]
                    final_r.append(line.rstrip(' -'))
                else: final_r.append(line)
            r = final_r
        return '\n'.join(r) if r else text

    def _apply_ai(self,text,target):
        if target=='cn':
            self.txt_cn.delete("1.0",tk.END); self.txt_cn.insert("1.0",text)
            self.lbl_info.config(text="✅ 完成",fg=self.C_ACCENT_GREEN); self.switch_to_cn()
        else:
            self.txt_en.delete("1.0",tk.END); self.txt_en.insert("1.0",text)
            self.lbl_info.config(text="✅ Done",fg=self.C_ACCENT_GREEN); self.switch_to_en()
        self._highlight_syntax(self.txt_cn); self._highlight_syntax(self.txt_en)
        self._sync_opts_to_text(); self._save_undo_state()

    def switch_to_cn(self):
        self.current_lang='cn'; self.tab_cn_btn.config(bg=self.C_ACCENT_BLUE,fg="white",font=("微软雅黑",9,"bold"))
        self.tab_en_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.tab_prompt_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.txt_en.pack_forget(); self.prompt_container.pack_forget()
        self.txt_cn.pack(fill="both",expand=True); self.refresh_preview()
        
    def switch_to_en(self):
        self.current_lang='en'; self.tab_en_btn.config(bg=self.C_ACCENT_BLUE,fg="white",font=("微软雅黑",9,"bold"))
        self.tab_cn_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.tab_prompt_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.txt_cn.pack_forget(); self.prompt_container.pack_forget()
        self.txt_en.pack(fill="both",expand=True); self.refresh_preview()

    def switch_to_prompt(self):
        self.tab_prompt_btn.config(bg=self.C_ACCENT_ORANGE,fg="white",font=("微软雅黑",9,"bold"))
        self.tab_cn_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.tab_en_btn.config(bg=self.C_BG_HDR,fg=self.C_TEXT,font=("微软雅黑",9))
        self.txt_cn.pack_forget(); self.txt_en.pack_forget()
        self.prompt_container.pack(fill="both",expand=True)

    def load_doc(self):
        p=filedialog.askopenfilename(filetypes=[("文档","*.pdf *.docx")])
        if not p: return
        try:
            content=[]
            if p.endswith('.pdf'):
                with pdfplumber.open(p) as pdf:
                    for pg in pdf.pages:
                        t=pg.extract_text()
                        if t: content.append(t)
                        for tbl in pg.extract_tables():
                            content.append("\n[表格数据开始]")
                            for row in tbl: content.append("| "+" | ".join(str(c).replace('\n',' ') for c in row if c is not None)+" |")
                            content.append("[表格数据结束]\n")
            else:
                doc=Document(p)
                for para in doc.paragraphs: content.append(para.text)
                for tbl in doc.tables:
                    content.append("\n[表格内容]")
                    for row in tbl.rows: content.append("| "+" | ".join(c.text.replace('\n',' ') for c in row.cells)+" |")
            self.raw_text="\n".join(content); self.lbl_info.config(text=f"✅ 解析成功({len(self.raw_text)}字)",fg=self.C_ACCENT_GREEN)
        except Exception as e: messagebox.showerror("读取错误",str(e))
        self._doc_extract_files = [p]
        self._doc_extract_imgs.clear()
        self._doc_show_empty_hint()
        self.doc_extract_files_lbl.config(text=f"已选: {os.path.basename(p)}")
        self._switch_bottom_tab('doc')
        if not self._bottom_panel_expanded: self._toggle_bottom_panel()
        threading.Thread(target=self._doc_do_extract_and_done, daemon=True).start()

    def clean_markdown(self,t):
        if not t: return ""
        t=re.sub(r'\*\*(.+?)\*\*',r'\1',t); t=re.sub(r'__(.+?)__',r'\1',t)
        t=t.replace('**','').replace('__',''); t=re.sub(r'^#+\s*','',t)
        t=re.sub(r'^[Kk]ey\s*:\s*','',t)
        return t.strip()
        
    def pick_bg(self,tp):
        p=filedialog.askopenfilename()
        if p:
            if tp=='cover': self.bg_cover_path=p
            else: self.bg_body_path=p
            self.save_config(); self.refresh_preview()
            
    def _is_opt_key_enabled(self, key):
        if key == 'opt_img': return self.opt_product_image.get()
        if key == 'opt_pkg': return self.opt_product_package.get()
        if key.startswith('cust_'):
            idx = int(key.split('_')[1])
            if idx < len(self.custom_sections): return self.custom_sections[idx][0].get()
        return False

    def _get_opt_section_label(self, key, lang):
        if key == 'opt_img': return "产品图片" if lang == 'cn' else "Product Images"
        if key == 'opt_pkg': return "产品包装" if lang == 'cn' else "Product Packaging"
        if key.startswith('cust_'):
            idx = int(key.split('_')[1])
            if idx < len(self.custom_sections):
                return self.custom_sections[idx][1].get() if lang == 'cn' else self.custom_sections[idx][2].get()
        return ""

    def _is_section_keyword(self,text):
        c=text.strip()
        if c in CN_HEADER_KEYWORDS: return True
        if c.lower() in [k.lower() for k in EN_HEADER_KEYWORDS]: return True
        for en,cn_v,en_v in self.custom_sections:
            if en.get() and (c==cn_v.get().strip() or c.lower()==en_v.get().strip().lower()): return True
        return False
        
    def is_header(self,raw_line):
        raw=raw_line.strip(); clean=self.clean_markdown(raw)
        if re.match(r'^\*\*.+\*\*$',raw) and self._is_section_keyword(clean): return True
        if raw.startswith('#') and self._is_section_keyword(clean): return True
        if self._is_section_keyword(clean): return True
        if re.match(r'^\*\*.+\*\*$', raw) and not raw.startswith(('-','*• ')) and '|' not in raw:
            if len(clean) <= 30: return True
        return False
        
    def _get_bullet_prefix(self,index=0):
        sym=BULLET_STYLES.get(self.var_bullet.get(),"● ")
        if sym=="__NUM__":
            chars="①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"; return chars[index%len(chars)]+" "
        elif sym=="__DOT__": return f"{index+1}. "
        return sym
        
    def _page_top(self,pg): return pg*(self.A4_H+PAGE_GAP)
    
    def _draw_page_bg(self,pg,ox=0):
        top=self._page_top(pg)
        self.canvas.create_rectangle(ox+4, top+4, ox+self.A4_W+4, top+self.A4_H+4, fill="#D1D1D6", outline="")
        self.canvas.create_rectangle(ox,top,ox+self.A4_W,top+self.A4_H,fill="white",outline="")
        
        bg_p=self.bg_cover_path if pg==0 else self.bg_body_path; key=f'bg_{pg}'
        if bg_p and os.path.exists(bg_p) and key not in self.tk_cache:
            self.tk_cache[key]=ImageTk.PhotoImage(Image.open(bg_p).resize((self.A4_W,self.A4_H)))
        if key in self.tk_cache: self.canvas.create_image(ox+self.A4_W//2,top+self.A4_H//2,image=self.tk_cache[key])

    def _get_all_lines_fixed(self, lang=None):
        if lang is None: lang = self.current_lang
        txt = self.txt_cn if lang == 'cn' else self.txt_en
        lines = txt.get("1.0", "end-1c").split('\n')
        tail_frames = []
        while lines and self._parse_img_frame_tag(lines[-1].strip()) is not None: tail_frames.insert(0, lines.pop())
        while lines and not lines[-1].strip(): lines.pop()
        for key in self.opt_sections_order:
            if not self._is_opt_key_enabled(key): continue
            label = self._get_opt_section_label(key, lang)
            if label: lines += ["", f"**{label}**"]
        if tail_frames: lines += [""] + tail_frames
        return lines

    def _get_font_face(self,lang=None):
        if lang is None: lang=self.current_lang
        return self.var_cn_font.get() if lang=='cn' else self.var_en_font.get()

    def _draw_photo_row(self,ml,cy,cw,cols,rh,ips):
        colw=cw/cols
        for j in range(cols):
            x0=ml+j*colw+2; x1=ml+(j+1)*colw-2; y0,y1=cy,cy+rh
            path=ips[j] if j<len(ips) else None
            if path and os.path.exists(path):
                ck=f"photo_{path}_{int(colw)}x{rh}"
                if ck not in self.tk_cache:
                    try:
                        src=Image.open(path).convert("RGB"); gw,gh=int(colw-4),rh-4
                        sw,sh=src.size; sc=min(gw/sw,gh/sh)
                        nw,nh=max(1,int(sw*sc)),max(1,int(sh*sc))
                        src=src.resize((nw,nh),Image.LANCZOS)
                        bg=Image.new("RGB",(gw,gh),"#fff")
                        bg.paste(src,((gw-nw)//2,(gh-nh)//2))
                        self.tk_cache[ck]=ImageTk.PhotoImage(bg)
                    except: self.tk_cache[ck]=None
                ti=self.tk_cache.get(ck)
                if ti:
                    self.canvas.create_rectangle(x0,y0,x1,y1,fill="white",outline="")
                    self.canvas.create_image((x0+x1)/2,(y0+y1)/2,image=ti,anchor="center")
                else: self._draw_empty(x0,y0,x1,y1)
            else: self._draw_empty(x0,y0,x1,y1)
            self.canvas.create_rectangle(x0,y0,x1,y1,outline=self.C_BORDER,width=2,fill="")
            
    def _draw_empty(self,x0,y0,x1,y1):
        self.canvas.create_rectangle(x0,y0,x1,y1,fill=self.C_BG_MAIN,outline="")
        self.canvas.create_line(x0,y0,x1,y1,fill=self.C_BORDER); self.canvas.create_line(x1,y0,x0,y1,fill=self.C_BORDER)
        self.canvas.create_text((x0+x1)/2,(y0+y1)/2,text="＋选图",font=("微软雅黑",8),fill=self.C_TEXT_MUTED)

    def _draw_float_img(self, path, x0, y0, x1, y1):
        gw, gh = int(x1 - x0 - 4), int(y1 - y0 - 4)
        if gw <= 0 or gh <= 0: return
        ck = f"float_{path}_{gw}x{gh}"
        if ck not in self.tk_cache:
            try:
                src = Image.open(path).convert("RGB"); sw, sh = src.size
                sc = min(gw / sw, gh / sh)
                nw, nh = int(sw * sc), int(sh * sc)
                src = src.resize((nw, nh), Image.LANCZOS)
                bg = Image.new("RGB", (gw, gh), "#fff")
                bg.paste(src, ((gw - nw) // 2, (gh - nh) // 2))
                self.tk_cache[ck] = ImageTk.PhotoImage(bg)
            except: self.tk_cache[ck] = None
        ti = self.tk_cache.get(ck)
        if ti:
            self.canvas.create_rectangle(x0, y0, x1, y1, fill="white", outline="")
            self.canvas.create_image((x0 + x1) / 2, (y0 + y1) / 2, image=ti, anchor="center")
        self.canvas.create_rectangle(x0, y0, x1, y1, outline=self.C_BORDER, width=2, fill="")

    def _canvas_offset(self):
        cw=self.canvas.winfo_width()
        if cw<=1: cw=self.canvas.winfo_reqwidth()
        return max(0,(cw-self.A4_W)//2)

    def refresh_preview(self):
        self._dismiss_editor(); self.canvas.delete("all"); self.tk_cache={}; self._edit_items=[]
        ox=self._canvas_offset(); ml=int(self.M_LR*self.PX_PER_MM); mt=int(self.M_TB*self.PX_PER_MM)
        cw=self.A4_W-ml*2; lang=self.current_lang
        al=self._get_all_lines_fixed(lang)
        
        ats, abs_ = self.var_title_size.get(), self.var_body_size.get()
        self._update_status_bar() 
        
        face=self._get_font_face(lang); lg=self.var_line_gap.get(); tg=self.var_title_gap.get()
        tf=(face,ats,"bold"); bf=(face,abs_); bfb=(face,abs_,"bold"); cf=(face,max(ats+14,24),"bold")
        bc=0
        
        self._draw_page_bg(0, ox)
        cover_text = self.var_cover_cn.get().strip() if lang == 'cn' else self.var_cover_en.get().strip()
        if not cover_text: cover_text = "请输入封面产品名称" if lang == 'cn' else "Enter Product Name"
            
        cid = self.canvas.create_text(ox + self.A4_W // 2, self._page_top(0) + self.A4_H // 2, text=cover_text, font=cf, anchor="center", width=cw)
        bb = self.canvas.bbox(cid)
        if bb: self._edit_items.append((cid, -1, bb, cover_text, 'cover', cover_text))
        
        pg=1; self._draw_page_bg(pg,ox); pt=self._page_top(pg); cy=pt+mt; pb=pt+self.A4_H-mt
        cs=""; it=False; fa=None; FG=8
        
        def np():
            nonlocal pg,pt,cy,pb,fa
            pg+=1; self._draw_page_bg(pg,ox); pt=self._page_top(pg); cy=pt+mt; pb=pt+self.A4_H-mt; fa=None

        def get_text_area():
            if fa and cy < fa['y1']:
                if fa['s'] == 'left':
                    tx = fa['x1'] + FG; tw = ox + ml + cw - tx
                else:
                    tx = ox + ml; tw = fa['x0'] - FG - tx
                return tx - ox, max(tw, 80)
            return ml, cw

        def dt(x,y,t,f,a="nw",w=None):
            kw=dict(font=f,anchor=a)
            if w: kw["width"]=w
            return self.canvas.create_text(ox+x,y,text=t,**kw)

        opt_labels_cn = set(); opt_labels_en = set()
        for key in self.opt_sections_order:
            if self._is_opt_key_enabled(key):
                opt_labels_cn.add(self._get_opt_section_label(key, 'cn'))
                opt_labels_en.add(self._get_opt_section_label(key, 'en'))

        for idx,line in enumerate(al):
            if fa and cy >= fa['y1']: fa = None
            if not line.strip():
                cy+=lg;
                if cy>pb: np()
                it=False; bc=0; continue
                
            lt=line.strip()
            
            fi_float = self._parse_float_tag(lt)
            if fi_float:
                side, wp, path = fi_float
                iw = int(cw * wp / 100); ih = int(iw * 0.75)
                if side == 'left': ix0 = ox + ml; ix1 = ix0 + iw
                else: ix1 = ox + ml + cw; ix0 = ix1 - iw

                if path and os.path.exists(path): self._draw_float_img(path, ix0, cy, ix1, cy + ih)
                else: self.canvas.create_rectangle(ix0, cy, ix1, cy + ih, fill=self.C_BG_MAIN, outline=self.C_BORDER, width=2)
                
                fa = {'s': side, 'x0': ix0, 'y0': cy, 'x1': ix1, 'y1': cy + ih}
                self._edit_items.append((None, idx, (ix0, cy, ix1, cy + ih), lt, 'float_img', lt))
                continue
                
            fi=self._parse_img_frame_tag(lt)
            if fi is not None:
                it=False; cols,ips,scale = fi
                if cols >= 4:
                    row_cols = cols // 2; rows2 = 2
                else:
                    row_cols = cols; rows2 = 1
                    
                cy+=14
                if cols == 1:
                    cwp = cw * (scale / 100.0)
                    rhp = int(cwp * 0.72)
                    ml_offset = (cw - cwp) / 2
                else:
                    cwp = cw / row_cols
                    rhp = int(cwp * 0.72)
                    ml_offset = 0
                    
                total_h = rhp * rows2 + 4 * (rows2-1)
                if cy+total_h>pb: np()
                for ri in range(rows2):
                    row_ips = ips[ri*row_cols:(ri+1)*row_cols]
                    self._draw_photo_row(ox+ml+ml_offset, cy+ri*(rhp+4), cwp if cols==1 else cw, row_cols, rhp, row_ips)
                bb=(ox+ml,cy,ox+ml+cw,cy+total_h); self._edit_items.append((None,idx,bb,lt,'img_frame', lt))
                cy+=total_h+8; continue
                
            rel_tx, dynamic_tw = get_text_area()
            
            if self.is_header(lt):
                it=False; bc=0; cs=self.clean_markdown(lt); cy+=6
                if cy+30>pb: 
                    np(); rel_tx, dynamic_tw = get_text_area()
                tid=dt(rel_tx,cy,cs,tf,w=dynamic_tw); bb=self.canvas.bbox(tid); h=bb[3]-bb[1] if bb else 20
                if bb:
                    is_opt = cs in opt_labels_cn or cs in opt_labels_en
                    ltype = 'opt_section' if is_opt else 'header'
                    self._edit_items.append((tid,idx,bb,cs,ltype, lt))
                cy+=h+tg
            elif '|' in lt:
                if '---' in lt: continue
                parts = lt.strip().strip('|').split('|')
                cells = [self.clean_markdown(c.strip()) for c in parts]
                if not parts: continue
                ih=not it; it=True; nc=len(cells); cw2=dynamic_tw/nc; ml2=1
                for c2 in cells:
                    cpl=max(1,int((cw2-8)/8)); ml2=max(ml2,-(-len(c2)//cpl))
                rh2=max(28,ml2*18+8)
                if cy+rh2>pb: 
                    np(); rel_tx, dynamic_tw = get_text_area(); cw2=dynamic_tw/nc
                for j,c2 in enumerate(cells):
                    x0=ox+rel_tx+j*cw2; x1=x0+cw2; fl=self.C_BG_HDR if ih else "white"
                    self.canvas.create_rectangle(x0,cy,x1,cy+rh2,outline=self.C_BORDER,width=1,fill=fl)
                    fu=bfb if ih else bf
                    ci2=self.canvas.create_text(x0+cw2/2,cy+rh2/2,text=c2,font=fu,anchor="center",width=cw2-8)
                    cb2=self.canvas.bbox(ci2)
                    if cb2: self._edit_items.append((ci2,idx,(x0,cy,x1,cy+rh2),c2,'table_cell', lt, j))
                cy+=rh2
            elif lt.startswith(('-','*','•')):
                it=False; dkw="产品描述" if lang=='cn' else "Product Description"
                rt=self.clean_markdown(lt[1:])
                if dkw in cs: ctxt=rt
                else: bp=self._get_bullet_prefix(bc); ctxt=bp+rt; bc+=1
                tid=dt(rel_tx,cy,ctxt,bf,w=dynamic_tw); bb=self.canvas.bbox(tid); h=bb[3]-bb[1] if bb else 16
                if cy+h>pb: 
                    self.canvas.delete(tid); np(); rel_tx, dynamic_tw = get_text_area()
                    tid=dt(rel_tx,cy,ctxt,bf,w=dynamic_tw); bb=self.canvas.bbox(tid); h=bb[3]-bb[1] if bb else 16
                if bb: self._edit_items.append((tid,idx,bb,ctxt,'bullet', lt))
                cy+=h+lg
            else:
                it=False; ctxt=self.clean_markdown(lt)
                tid=dt(rel_tx,cy,ctxt,bf,w=dynamic_tw); bb=self.canvas.bbox(tid); h=bb[3]-bb[1] if bb else 16
                if cy+h>pb: 
                    self.canvas.delete(tid); np(); rel_tx, dynamic_tw = get_text_area()
                    tid=dt(rel_tx,cy,ctxt,bf,w=dynamic_tw); bb=self.canvas.bbox(tid); h=bb[3]-bb[1] if bb else 16
                if bb: self._edit_items.append((tid,idx,bb,ctxt,'text', lt))
                cy+=h+lg+2
                
        self.canvas.config(scrollregion=(0,0,self.A4_W+ox*2,(pg+1)*(self.A4_H+PAGE_GAP)))

    def export_pdf(self):
        lang=self.current_lang; td=tempfile.NamedTemporaryFile(suffix=".docx",delete=False).name
        try: self._gen_word(td,lang, True)
        except Exception as e: messagebox.showerror("失败",str(e)); return
        pp=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")])
        if not pp: os.unlink(td); return
        ok=False
        try:
            from docx2pdf import convert; convert(td,pp); ok=True
        except: pass
        if not ok:
            for lo in [r"C:\Program Files\LibreOffice\program\soffice.exe",r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                       "/usr/bin/libreoffice","/usr/bin/soffice","/Applications/LibreOffice.app/Contents/MacOS/soffice"]:
                if os.path.exists(lo):
                    try:
                        od=os.path.dirname(pp)
                        subprocess.run([lo,"--headless","--convert-to","pdf","--outdir",od,td],check=True,timeout=60,capture_output=True)
                        lp=os.path.join(od,os.path.splitext(os.path.basename(td))[0]+".pdf")
                        if os.path.exists(lp) and lp!=pp: os.replace(lp,pp)
                        ok=True; break
                    except: continue
        if not ok and sys.platform=='win32':
            try:
                import comtypes.client; w=comtypes.client.CreateObject('Word.Application'); w.Visible=False
                d=w.Documents.Open(os.path.abspath(td)); d.SaveAs(os.path.abspath(pp),FileFormat=17); d.Close(); w.Quit(); ok=True
            except: pass
        try: os.unlink(td)
        except: pass
        if ok and os.path.exists(pp): messagebox.showinfo("成功",f"PDF已导出！\n{pp}")
        else: messagebox.showwarning("PDF导出","转换失败，请安装docx2pdf/LibreOffice/MS Word")

    def _set_margins(self,sec):
        sec.page_width=Mm(210);sec.page_height=Mm(297);sec.top_margin=Mm(self.M_TB)
        sec.bottom_margin=Mm(self.M_TB);sec.left_margin=Mm(self.M_LR);sec.right_margin=Mm(self.M_LR)

    def _add_photo_tbl(self,doc,cols,ips,fn,scale=100):
        u=210-self.M_LR*2
        if cols >= 4:
            row_cols = cols // 2; n_rows = 2
        else:
            row_cols = cols; n_rows = 1
            
        if cols == 1:
            u = u * (scale / 100.0)
            
        cwm=u/row_cols; chm=cwm*0.72
        tab=doc.add_table(rows=n_rows,cols=row_cols); tab.style='Table Grid'
        tp=tab._tbl.tblPr; jc=OxmlElement('w:jc'); jc.set(qn('w:val'),'center'); tp.append(jc)
        tw=OxmlElement('w:tblW'); tw.set(qn('w:w'),str(int(u*56.69))); tw.set(qn('w:type'),'dxa'); tp.append(tw)
        for ri in range(n_rows):
            trp=tab.rows[ri]._tr.get_or_add_trPr(); trh=OxmlElement('w:trHeight')
            trh.set(qn('w:val'),str(int(chm*56.69))); trh.set(qn('w:hRule'),'exact'); trp.append(trh)
            for ci2 in range(row_cols):
                img_idx = ri*row_cols+ci2
                cell=tab.cell(ri,ci2); tcp=cell._tc.get_or_add_tcPr()
                tcw=OxmlElement('w:tcW'); tcw.set(qn('w:w'),str(int(cwm*56.69))); tcw.set(qn('w:type'),'dxa'); tcp.append(tcw)
                cell.vertical_alignment=WD_ALIGN_VERTICAL.CENTER; pc=cell.paragraphs[0]; pc.alignment=WD_ALIGN_PARAGRAPH.CENTER
                path=ips[img_idx] if img_idx<len(ips) else None
                if path and os.path.exists(path):
                    try:
                        src=Image.open(path).convert("RGB"); gw,gh=int(cwm*3.7795),int(chm*3.7795)
                        sw,sh=src.size
                        sc=min(gw/sw,gh/sh)
                        nw,nh=max(1,int(sw*sc)),max(1,int(sh*sc))
                        src=src.resize((nw,nh),Image.LANCZOS)
                        bg=Image.new("RGB",(gw,gh),"#fff")
                        bg.paste(src,((gw-nw)//2,(gh-nh)//2))
                        
                        with tempfile.NamedTemporaryFile(suffix=".jpg",delete=False) as tmp:
                            t2=tmp.name; bg.save(t2,"JPEG",quality=92)
                        pc.add_run().add_picture(t2,width=Mm(cwm-1)); os.unlink(t2)
                    except: pc.add_run("").font.name=fn
                else: pc.add_run("").font.name=fn
        doc.add_paragraph()

    def generate_word(self,lang='cn'):
        p=filedialog.asksaveasfilename(defaultextension=".docx",filetypes=[("Word","*.docx")])
        if not p: return
        try: self._gen_word(p,lang, True); messagebox.showinfo("成功",f"{'中文' if lang=='cn' else 'EN'} 规格书已成功导出！")
        except Exception as e:
            import traceback; messagebox.showerror("导出失败",f"{e}\n\n{traceback.format_exc()}")

    def _gen_word(self,p,lang, optimize_spacing=True):
        doc=Document(); fn=self.var_cn_font.get() if lang=='cn' else self.var_en_font.get()
        tp2=self.var_title_size.get(); bp=self.var_body_size.get()
        
        base_lsp = 4 if optimize_spacing else 0
        base_tsp = 6 if optimize_spacing else 0
        lsp = max(0, round(self.var_line_gap.get()/1.33)) + base_lsp
        tsp = max(0, round(self.var_title_gap.get()/1.33)) + base_tsp
        
        cs=doc.sections[0]; self._set_margins(cs); doc.styles['Normal'].font.name=fn
        if lang=='cn': doc.styles['Normal']._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),fn)
        lines=self._get_all_lines_fixed(lang)
        
        if self.bg_cover_path: self._add_bg(cs,self.bg_cover_path)
        cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER
        for _ in range(14): cp.add_run('\n')
        
        cover_text = self.var_cover_cn.get().strip() if lang == 'cn' else self.var_cover_en.get().strip()
        rn=cp.add_run(cover_text if cover_text else ("请输入产品名" if lang == 'cn' else "Enter Product Name"))
        rn.bold=True; rn.font.size=Pt(max(tp2+14,24)); rn.font.name=fn
        if lang=='cn': rn._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),fn)
        
        doc.add_page_break(); bs=doc.add_section(); self._set_margins(bs)
        if self.bg_body_path: self._add_bg(bs,self.bg_body_path)
        tr=[]; it=False; bc=0; csc=""
        
        def arf(pa,text,size=None,bold=False):
            r=pa.add_run(text); r.font.size=Pt(size or bp); r.font.name=fn; r.bold=bold
            if lang=='cn': r._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),fn)
            return r
            
        for idx,line in enumerate(lines):
            if not line.strip(): it=False; bc=0; continue
            lt=line.strip()
            
            fi_float = self._parse_float_tag(lt)
            if fi_float:
                it = False; bc = 0
                side, wp, path = fi_float
                if path and os.path.exists(path):
                    try:
                        pa = doc.add_paragraph()
                        pa.paragraph_format.space_after = Pt(0)
                        avail_w = 210 - self.M_LR * 2
                        img_w = avail_w * (wp / 100.0)
                        pic = pa.add_run().add_picture(path, width=Mm(img_w))
                        
                        inline = pic._inline
                        anc = OxmlElement('wp:anchor')
                        anc.set('distT', "0"); anc.set('distB', "0")
                        anc.set('distL', "114300"); anc.set('distR', "114300")
                        anc.set('simplePos', "0"); anc.set('relativeHeight', "251658240")
                        anc.set('behindDoc', "0"); anc.set('locked', "0")
                        anc.set('layoutInCell', "1"); anc.set('allowOverlap', "1")
                        
                        spos = OxmlElement('wp:simplePos'); spos.set('x', "0"); spos.set('y', "0")
                        anc.append(spos)
                        
                        posH = OxmlElement('wp:positionH'); posH.set('relativeFrom', "margin")
                        align = OxmlElement('wp:align'); align.text = side
                        posH.append(align); anc.append(posH)
                        
                        posV = OxmlElement('wp:positionV'); posV.set('relativeFrom', "paragraph")
                        posOff = OxmlElement('wp:posOffset'); posOff.text = "0"
                        posV.append(posOff); anc.append(posV)
                        
                        for tag_str in ['wp:extent', 'wp:effectExtent']:
                            el = inline.find(qn(tag_str))
                            if el is not None: anc.append(el)
                            
                        if inline.find(qn('wp:effectExtent')) is None:
                            eff = OxmlElement('wp:effectExtent')
                            eff.set('l', "0"); eff.set('t', "0"); eff.set('r', "0"); eff.set('b', "0")
                            anc.append(eff)
                            
                        wrap = OxmlElement('wp:wrapSquare'); wrap.set('wrapText', "largest")
                        anc.append(wrap)
                        
                        for tag_str in ['wp:docPr', 'wp:cNvGraphicFramePr', 'a:graphic']:
                            el = inline.find(qn(tag_str))
                            if el is not None: anc.append(el)
                            
                        inline.getparent().replace(inline, anc)
                    except: pass
                continue

            fi=self._parse_img_frame_tag(lt)
            if fi is not None: 
                it=False; doc.add_paragraph()
                self._add_photo_tbl(doc, fi[0], fi[1], fn, fi[2])
                continue
            
            if '|' in lt:
                if '---' in lt: continue
                parts = lt.strip().strip('|').split('|')
                cells = [self.clean_markdown(c.strip()) for c in parts]
                if not parts: continue
                if not tr: tr.append(cells); it=True
                else: tl=len(tr[0]); tr.append((cells+[""]*tl)[:tl])
                nxt=lines[idx+1].strip() if idx+1<len(lines) else ""
                if not nxt or '|' not in nxt:
                    t=doc.add_table(rows=len(tr),cols=len(tr[0])); t.style='Table Grid'
                    tp3=t._tbl.tblPr; jc3=OxmlElement('w:jc'); jc3.set(qn('w:val'),'center'); tp3.append(jc3)
                    for ri,rd in enumerate(tr):
                        for ci2,val in enumerate(rd):
                            cl=t.cell(ri,ci2); cl.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
                            pc=cl.paragraphs[0]; pc.alignment=WD_ALIGN_PARAGRAPH.CENTER
                            rc=pc.add_run(val); rc.font.size=Pt(bp); rc.font.name=fn
                            if lang=='cn': rc._element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),fn)
                            if ri==0: rc.bold=True 
                    tr=[]; it=False; doc.add_paragraph()
            elif self.is_header(lt):
                it=False; bc=0; csc=self.clean_markdown(lt)
                pa=doc.add_paragraph(); pa.paragraph_format.space_before=Pt(tsp+base_tsp) 
                pa.paragraph_format.space_after=Pt(tsp); arf(pa,csc,tp2,True)
            elif lt.startswith(('-','*','•')):
                it=False; dkw="产品描述" if lang=='cn' else "Product Description"
                rt=self.clean_markdown(lt[1:])
                if dkw in csc: bt=rt
                else: bt=self._get_bullet_prefix(bc)+rt; bc+=1
                pa=doc.add_paragraph(); pa.paragraph_format.space_after=Pt(lsp); arf(pa,bt)
            else:
                it=False; pa=doc.add_paragraph(); pa.paragraph_format.space_after=Pt(lsp+base_lsp) 
                arf(pa,self.clean_markdown(lt))
        doc.save(p)

    def _add_bg(self,section,path):
        hdr=section.header; hdr.is_linked_to_previous=False
        run=hdr.paragraphs[0].add_run(); pic=run.add_picture(path,width=Mm(210),height=Mm(297))
        anc=OxmlElement('wp:anchor'); anc.set(qn('wp:behindDoc'),'1')
        for tag,val in [('wp:positionH','page'),('wp:positionV','page')]:
            pos=OxmlElement(tag); pos.set(qn('wp:relativeFrom'),val)
            off=OxmlElement('wp:posOffset'); off.text='0'; pos.append(off); anc.append(pos)
        for child in pic._inline: anc.append(child)
        pic._inline.getparent().replace(pic._inline,anc)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE,"r", encoding="utf-8") as f: d=json.load(f)
                
                self.layout_config = d.get("layout_config", {"edit_w": 400, "gallery_w": 140, "bottom_h": 160})
                
                self.bg_cover_path=d.get("bg_cover_path",""); self.bg_body_path=d.get("bg_body_path","")
                self.var_title_size.set(d.get("title_size",14)); self.var_body_size.set(d.get("body_size",11))
                self.var_line_gap.set(d.get("line_gap",8)); self.var_title_gap.set(d.get("title_gap",14))
                self.var_cn_font.set(d.get("cn_font","微软雅黑")); self.var_en_font.set(d.get("en_font","Arial"))
                self.var_bullet.set(d.get("bullet","●  实心圆")); self.var_auto_shrink.set(d.get("auto_shrink",True))
                
                self.var_feature_brief.set(d.get("feature_brief", True))
                self.var_cover_cn.set(d.get("cover_cn", "")); self.var_cover_en.set(d.get("cover_en", ""))
                self.var_custom_prompt.set(d.get("custom_prompt", False))
                pcn = d.get("prompt_cn", ""); pen = d.get("prompt_en", "")
                
                self.kimi_api_key = d.get("kimi_api_key", "")
                self.baidu_appid = d.get("baidu_appid", ""); self.baidu_appkey = d.get("baidu_appkey", "")
                
                if pcn: self.tp_cn.config(state=tk.NORMAL); self.tp_cn.delete("1.0", tk.END); self.tp_cn.insert("1.0", pcn)
                if pen: self.tp_en.config(state=tk.NORMAL); self.tp_en.delete("1.0", tk.END); self.tp_en.insert("1.0", pen)
                self._set_prompt_ui(self.var_custom_prompt.get())
                
                for sd in d.get("custom_sections",[]):
                    self.custom_sections.append((tk.BooleanVar(value=sd.get("enabled",True)),
                        tk.StringVar(value=sd.get("cn","自定义章节")),tk.StringVar(value=sd.get("en","Custom Section"))))
            except: pass
            
    def save_config(self):
        cd=[{"enabled":e.get(),"cn":c.get(),"en":n.get()} for e,c,n in self.custom_sections]
        s1 = self.tp_cn.cget("state"); s2 = self.tp_en.cget("state")
        self.tp_cn.config(state=tk.NORMAL); self.tp_en.config(state=tk.NORMAL)
        pcn = self.tp_cn.get("1.0", tk.END).strip(); pen = self.tp_en.get("1.0", tk.END).strip()
        self.tp_cn.config(state=s1); self.tp_en.config(state=s2)
        try:
            with open(CONFIG_FILE,"w", encoding="utf-8") as f:
                json.dump({
                    "layout_config": self.layout_config,
                    "bg_cover_path":self.bg_cover_path,"bg_body_path":self.bg_body_path,
                    "title_size":self.var_title_size.get(),"body_size":self.var_body_size.get(),
                    "line_gap":self.var_line_gap.get(),"title_gap":self.var_title_gap.get(),
                    "cn_font":self.var_cn_font.get(),"en_font":self.var_en_font.get(),
                    "bullet":self.var_bullet.get(),"auto_shrink":self.var_auto_shrink.get(),
                    "feature_brief": self.var_feature_brief.get(),
                    "cover_cn": self.var_cover_cn.get(), "cover_en": self.var_cover_en.get(),
                    "custom_prompt": self.var_custom_prompt.get(),
                    "prompt_cn": pcn, "prompt_en": pen,
                    "kimi_api_key": self.kimi_api_key,
                    "baidu_appid": self.baidu_appid, "baidu_appkey": self.baidu_appkey,
                    "custom_sections":cd
                }, f, ensure_ascii=False)
        except: pass

if __name__=="__main__":
    root=tk.Tk(); app=HuamaiApp(root); root.mainloop()