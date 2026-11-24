"""上傳伺服器模式：開一個 HTTP 網頁，讓別人輸入「動態密碼」後上傳檔案到程式目錄的 downloads 資料夾。

放在 modes 資料夾裡，重新執行 main.py 就會多一個按鈕。
只用 Python 內建模組，不需要額外安裝東西。

上傳流程（網頁與程式配合）：
  1. 網頁先送「上傳請求」（帶密碼、檔名、大小）到 /request
  2. 「一律允許」模式：立刻核准。「手動審核」模式：等程式這邊的使用者按允許／拒絕
  3. 核准後網頁才開始送檔案到 /upload（用核准時給的一次性代碼，不再送密碼）
"""

import hmac
import json
import math
import os
import queue
import re
import secrets
import socket
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from tkinter import messagebox, ttk
from urllib.parse import parse_qs, unquote, urlparse

from core import PAL
from modes.base import Mode

# ===== 設定區 =====
DEFAULT_PORT = 8000  # 預設連接埠
MAX_FILE_SIZE = 2 * 1024 * 1024 * 1024  # 單檔上限 2 GB
PW_SECONDS = 30  # 動態密碼幾秒換一次
PW_GRACE = 5  # 換密碼後，舊密碼再多給幾秒（避免剛好卡在換密碼的瞬間輸入失敗）
MAX_FAILS = 5  # 同一個 IP 連續輸錯幾次就鎖定
LOCK_SECONDS = 60  # 鎖定秒數
APPROVAL_TIMEOUT = 120  # 手動審核等多久沒回應就自動拒絕（秒）
LOG_FILE_NAME = "server_log.log"  # log 存檔名稱（不跟別的模式撞名）
SYSTEM_SOURCE = "系統"  # log 裡「不是來自網路的事件」顯示的來源
# ==================


def _base_dir():
  """程式所在目錄（打包成 exe 時是 exe 所在目錄）。"""
  if getattr(sys, "frozen", False):
    return os.path.dirname(sys.executable)
  return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


BASE_DIR = _base_dir()
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
LOG_PATH = os.path.join(BASE_DIR, LOG_FILE_NAME)


def get_local_ip():
  try:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))  # 不會真的送資料，只是問系統走哪張網卡
    ip = s.getsockname()[0]
    s.close()
    return ip
  except OSError:
    return "127.0.0.1"


def safe_filename(name):
  """把檔名清乾淨，避免跑到別的資料夾或用到 Windows 不能用的字。"""
  name = os.path.basename(name.replace("\\", "/")).strip().strip(".")
  name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
  if not name:
    name = "unnamed"
  return name[:200]


def unique_path(folder, name):
  """檔名重複時自動加 (1)、(2)，不會蓋掉舊檔。"""
  path = os.path.join(folder, name)
  if not os.path.exists(path):
    return path
  stem, ext = os.path.splitext(name)
  i = 1
  while True:
    path = os.path.join(folder, f"{stem} ({i}){ext}")
    if not os.path.exists(path):
      return path
    i += 1


def fmt_size(n):
  n = float(n)
  for unit in ("B", "KB", "MB", "GB"):
    if n < 1024 or unit == "GB":
      return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
    n /= 1024


def fmt_speed(bps):
  return fmt_size(bps) + "/s"


# 網頁樣式仿 sv_ttk（Sun Valley／Windows 11 Fluent 風格），自動跟隨系統淺色／深色
UPLOAD_PAGE = """<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>檔案上傳</title>
<style>
:root{--bg:#f3f3f3;--card:#fff;--fg:#1a1a1a;--sub:#605e5c;--line:#e5e5e5;--ctl:#fdfdfd;
--ctl-line:#d1d1d1;--accent:#005fb8;--accent-fg:#fff;--ok:#0f7b0f;--err:#c42b1c;--track:#e0e0e0}
@media(prefers-color-scheme:dark){:root{--bg:#1c1c1c;--card:#2b2b2b;--fg:#fff;--sub:#c5c5c5;
--line:#3a3a3a;--ctl:#323232;--ctl-line:#454545;--accent:#57c8ff;--accent-fg:#000;--ok:#6ccb5f;
--err:#ff99a4;--track:#454545}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);padding:32px 12px;display:flex;justify-content:center;
font-family:"Segoe UI Variable","Segoe UI","Microsoft JhengHei",system-ui,sans-serif;font-size:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:24px;width:100%;max-width:480px}
h2{margin:0 0 4px;font-size:22px;font-weight:600}
.sub{color:var(--sub);margin:0 0 18px;font-size:13px}
label{display:block;margin:14px 0 6px;font-size:13px}
input[type=text]{width:100%;padding:8px 10px;font-size:16px;letter-spacing:2px;color:var(--fg);
background:var(--ctl);border:1px solid var(--ctl-line);border-bottom-width:1px;border-radius:4px;outline:none;font-family:inherit}
input[type=text]:focus{border-bottom:2px solid var(--accent);padding-bottom:7px}
input[type=file]{width:100%;color:var(--sub);font-size:13px;font-family:inherit}
input[type=file]::file-selector-button{padding:7px 14px;margin-right:10px;border-radius:4px;cursor:pointer;
background:var(--ctl);color:var(--fg);border:1px solid var(--ctl-line);font-family:inherit}
button{width:100%;margin-top:18px;padding:9px;border:none;border-radius:4px;cursor:pointer;font-size:14px;
background:var(--accent);color:var(--accent-fg);font-family:inherit}
button:hover{opacity:.9}
button:disabled{opacity:.45;cursor:default}
.item{margin-top:14px;padding-top:12px;border-top:1px solid var(--line);word-break:break-all}
.name{font-weight:600}
.bar{height:4px;margin:8px 0 6px;background:var(--track);border-radius:2px;overflow:hidden}
.bar div{height:100%;width:0;background:var(--accent);border-radius:2px;transition:width .15s}
.msg{font-size:12px;color:var(--sub)}
.ok{color:var(--ok)}.err{color:var(--err)}
</style></head><body><div class="card">
<h2>檔案上傳</h2>
<p class="sub">請輸入對方程式上顯示的 8 位數密碼（每 30 秒會更換）</p>
<label for="pw">密碼</label>
<input type="text" id="pw" inputmode="numeric" autocomplete="off" maxlength="9" placeholder="0000 0000">
<label for="files">檔案</label>
<input type="file" id="files" multiple>
<button id="btn" onclick="go()">上傳</button>
<div id="list"></div>
</div>
<script>
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function fmt(n){var u=["B","KB","MB","GB"],i=0;while(n>=1024&&i<3){n/=1024;i++;}return (i?n.toFixed(1):n.toFixed(0))+" "+u[i];}
function say(row,text,cls){var m=row.querySelector(".msg");m.className="msg "+(cls||"");m.textContent=text;}
async function sendOne(file,pw,row){
  say(row,"正在送出請求…");
  var r,j={};
  try{
    r=await fetch("/request",{method:"POST",headers:{"X-Password":pw,
      "X-Filename":encodeURIComponent(file.name),"X-Size":String(file.size)}});
    j=await r.json();
  }catch(e){say(row,"失敗：無法連線","err");return 0;}
  if(!r.ok){say(row,"失敗："+(j.error||r.status),"err");return r.status;}
  var id=j.id,st=j.status;
  while(st==="pending"){
    say(row,"等待對方批准…");
    await sleep(1000);
    try{var s=await fetch("/status?id="+encodeURIComponent(id));var sj=await s.json();st=sj.status||"error";}
    catch(e){st="error";}
  }
  if(st!=="approved"){
    var t={denied:"對方拒絕了這個檔案",timeout:"等待批准逾時",closed:"伺服器已關閉"};
    say(row,"未上傳："+(t[st]||"發生錯誤"),"err");return 0;
  }
  return new Promise(function(done){
    var bar=row.querySelector(".bar div"),t0=Date.now();
    var x=new XMLHttpRequest();
    x.open("POST","/upload");
    x.setRequestHeader("X-Id",id);
    x.upload.onprogress=function(e){
      if(e.lengthComputable){
        var sec=(Date.now()-t0)/1000||0.001;
        bar.style.width=(e.loaded/e.total*100)+"%";
        say(row,Math.round(e.loaded/e.total*100)+"%　"+fmt(e.loaded)+" / "+fmt(e.total)+"　"+fmt(e.loaded/sec)+"/s");
      }
    };
    x.onload=function(){
      var rr={};try{rr=JSON.parse(x.responseText);}catch(e){}
      if(x.status==200){var sec=(Date.now()-t0)/1000||0.001;
        say(row,"完成　平均 "+fmt(file.size/sec)+"/s","ok");bar.style.width="100%";}
      else{say(row,"失敗："+(rr.error||x.status),"err");}
      done(x.status);
    };
    x.onerror=function(){say(row,"失敗：連線中斷","err");done(0);};
    x.send(file);
  });
}
async function go(){
  var pw=document.getElementById("pw").value.replace(/\\s/g,"");
  var fs=document.getElementById("files").files;
  if(!/^\\d{8}$/.test(pw)){alert("請輸入 8 位數密碼");return;}
  if(!fs.length){alert("請選擇檔案");return;}
  var btn=document.getElementById("btn");btn.disabled=true;
  var list=document.getElementById("list");list.innerHTML="";
  for(var i=0;i<fs.length;i++){
    var row=document.createElement("div");row.className="item";
    row.innerHTML="<div class='name'></div><div class='bar'><div></div></div><div class='msg'></div>";
    row.querySelector(".name").textContent=fs[i].name;
    list.appendChild(row);
    var st=await sendOne(fs[i],pw,row);
    if(st==403||st==429){break;}
  }
  btn.disabled=false;
}
</script></body></html>
"""


def make_handler(mode):
  """產生處理網頁請求的類別，並讓它能呼叫 mode 的功能。"""

  class Handler(BaseHTTPRequestHandler):
    timeout = 60  # 連線卡住 60 秒就放棄，避免佔住執行緒
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):  # 關掉預設的終端機輸出
      pass

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
      if isinstance(body, (dict, list)):
        body = json.dumps(body, ensure_ascii=False)
      data = body.encode("utf-8")
      self.send_response(code)
      self.send_header("Content-Type", ctype)
      self.send_header("Content-Length", str(len(data)))
      self.send_header("Cache-Control", "no-store")
      if code >= 400:
        self.send_header("Connection", "close")
        self.close_connection = True
      self.end_headers()
      self.wfile.write(data)

    # ---------- GET ----------
    def do_GET(self):
      ip = self.client_address[0]
      url = urlparse(self.path)
      if url.path in ("/", "/index.html"):
        mode.log("開啟上傳頁面", ip)
        self._send(200, UPLOAD_PAGE, "text/html; charset=utf-8")
      elif url.path == "/status":
        rid = parse_qs(url.query).get("id", [""])[0]
        self._send(200, {"status": mode.request_status(rid, ip)})
      else:
        self._send(404, {"error": "找不到頁面"})

    # ---------- POST ----------
    def do_POST(self):
      ip = self.client_address[0]
      if self.path == "/request":
        self._handle_request(ip)
      elif self.path == "/upload":
        self._handle_upload(ip)
      else:
        self._send(404, {"error": "找不到頁面"})

    def _handle_request(self, ip):
      if mode.server is None:
        self._send(503, {"error": "伺服器已關閉"})
        return
      if mode.is_locked(ip):
        mode.log("被鎖定中，拒絕請求", ip)
        self._send(429, {"error": f"輸錯太多次，請等 {LOCK_SECONDS} 秒後再試"})
        return

      pw = unquote(self.headers.get("X-Password", "")).replace(" ", "")
      if not mode.check_password(pw):
        mode.record_fail(ip)
        mode.log("密碼錯誤", ip)
        self._send(403, {"error": "密碼錯誤（密碼每 30 秒會換，請看最新的）"})
        return
      mode.clear_fail(ip)

      try:
        size = int(self.headers.get("X-Size", ""))
      except ValueError:
        self._send(400, {"error": "缺少檔案大小"})
        return
      if size <= 0:
        self._send(400, {"error": "檔案是空的"})
        return
      if size > MAX_FILE_SIZE:
        mode.log(f"檔案過大被拒絕（{fmt_size(size)}）", ip)
        self._send(413, {"error": f"檔案超過上限 {fmt_size(MAX_FILE_SIZE)}"})
        return

      name = safe_filename(unquote(self.headers.get("X-Filename", "")))
      rid, status = mode.add_request(ip, name, size)
      if status == "approved":
        mode.log(f"請求上傳：{name}（{fmt_size(size)}），一律允許模式，自動核准", ip)
      else:
        mode.log(f"請求上傳：{name}（{fmt_size(size)}），等待使用者審核", ip)
      self._send(200, {"ok": True, "id": rid, "status": status})

    def _handle_upload(self, ip):
      req = mode.take_request(self.headers.get("X-Id", ""), ip)
      if req is None:
        mode.log("嘗試上傳但沒有已批准的請求", ip)
        self._send(403, {"error": "沒有已批准的上傳請求"})
        return
      name, size = req["name"], req["size"]

      try:
        length = int(self.headers.get("Content-Length", ""))
      except ValueError:
        self._send(411, {"error": "缺少檔案大小"})
        return
      if length != size:
        mode.log(f"檔案大小與請求不符：{name}", ip)
        self._send(400, {"error": "檔案大小與請求不符"})
        return

      os.makedirs(DOWNLOAD_DIR, exist_ok=True)
      target = unique_path(DOWNLOAD_DIR, name)
      part = target + ".part"
      tid = mode.begin_transfer(ip, os.path.basename(target), size)
      remaining = size
      t0 = time.time()
      try:
        with open(part, "wb") as f:
          while remaining > 0:
            if mode.server is None:
              raise ConnectionAbortedError("伺服器已關閉")
            chunk = self.rfile.read(min(remaining, 256 * 1024))
            if not chunk:
              break
            f.write(chunk)
            remaining -= len(chunk)
            mode.xfer_progress(tid, size - remaining)
        if remaining > 0:
          raise ConnectionError("上傳中途斷線")
        os.replace(part, target)
      except Exception as e:
        try:
          os.remove(part)
        except OSError:
          pass
        mode.log(f"上傳失敗：{name}（{type(e).__name__}: {e}）", ip)
        self.close_connection = True
        return
      finally:
        mode.end_transfer(tid)

      sec = max(time.time() - t0, 0.001)
      mode.log(
          f"上傳成功：{os.path.basename(target)}（{fmt_size(size)}，平均 {fmt_speed(size / sec)}）",
          ip,
      )
      mode.files_dirty = True
      self._send(200, {"ok": True, "name": os.path.basename(target)})

  return Handler


class ServerMode(Mode):
  title = "🌐 上傳伺服器"
  order = 40

  def __init__(self, app):
    super().__init__(app)
    self.server = None
    self.thread = None
    self.log_queue = queue.Queue()
    self.log_lines = []  # 離開頁面再回來時，可以把舊 log 補回去
    self.files_dirty = True
    self.fail_info = {}  # ip -> [失敗次數, 鎖定到什麼時候]
    self.fail_lock = threading.Lock()

    # 動態密碼
    self.pw_lock = threading.Lock()
    self.code = None
    self.prev_code = None
    self.prev_until = 0.0
    self.pw_expire = 0.0

    # 上傳請求（審核用）與傳輸中的檔案
    self.req_lock = threading.Lock()
    self.requests = {}  # id -> {ip, name, size, status, created, used}
    self.xfer_lock = threading.Lock()
    self.transfers = {}  # tid -> {ip, name, size, done, speed, d_last, t_last}
    self.approval_mode = "manual"  # 給執行緒讀的普通變數（tk 變數不能跨執行緒讀）
    self.mode_var = tk.StringVar(master=app, value=self.approval_mode)

    self.port_var = tk.StringVar(master=app, value=str(DEFAULT_PORT))
    self.status_var = tk.StringVar(master=app, value="伺服器：已關閉")
    self.url_var = tk.StringVar(master=app, value="")
    self.code_var = tk.StringVar(master=app, value="－－－－ －－－－")
    self.pw_left_var = tk.StringVar(master=app, value="")
    self.speed_var = tk.StringVar(master=app, value="")

    self.tree = None
    self.text = None
    self.btn_toggle = None
    self.port_entry = None
    self.pw_bar = None
    self.pending_tree = None
    self.xfer_frame = None
    self.idle_label = None
    self.rows = {}  # tid -> 傳輸中那一列的元件
    self._pending_shown = ()
    self._poll_id = None

  # ---------- 給伺服器執行緒用的函式（不能直接碰視窗元件） ----------
  def log(self, msg, source=SYSTEM_SOURCE):
    """寫 log。每一行都有來源 IP；不是來自網路的事件，來源寫「系統」。"""
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{source}] {msg}"
    self.log_queue.put(line)
    try:
      with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    except OSError:
      pass

  def is_locked(self, ip):
    with self.fail_lock:
      info = self.fail_info.get(ip)
      return bool(info and info[1] > time.time())

  def record_fail(self, ip):
    with self.fail_lock:
      info = self.fail_info.setdefault(ip, [0, 0])
      info[0] += 1
      if info[0] >= MAX_FAILS:
        info[0] = 0
        info[1] = time.time() + LOCK_SECONDS

  def clear_fail(self, ip):
    with self.fail_lock:
      self.fail_info.pop(ip, None)

  # ----- 動態密碼 -----
  @staticmethod
  def _new_code():
    return f"{secrets.randbelow(10 ** 8):08d}"

  def _rotate(self):
    now = time.time()
    with self.pw_lock:
      if self.code is not None and now >= self.pw_expire:
        self.prev_code = self.code
        self.prev_until = now + PW_GRACE
        self.code = self._new_code()
        self.pw_expire = now + PW_SECONDS

  def get_code(self):
    """回傳 (目前密碼, 剩幾秒換)。伺服器沒開時密碼是 None。"""
    self._rotate()
    with self.pw_lock:
      return self.code, max(0.0, self.pw_expire - time.time())

  def check_password(self, pw):
    self._rotate()
    with self.pw_lock:
      ok = False
      if self.code is not None and hmac.compare_digest(
          pw.encode("utf-8"), self.code.encode("utf-8")
      ):
        ok = True
      if (
          self.prev_code is not None
          and time.time() < self.prev_until
          and hmac.compare_digest(pw.encode("utf-8"), self.prev_code.encode("utf-8"))
      ):
        ok = True
      return ok

  # ----- 上傳請求 -----
  def add_request(self, ip, name, size):
    rid = secrets.token_urlsafe(12)
    status = "approved" if self.approval_mode == "allow" else "pending"
    with self.req_lock:
      self.requests[rid] = {
          "ip": ip,
          "name": name,
          "size": size,
          "status": status,
          "created": time.time(),
          "used": False,
      }
    return rid, status

  def request_status(self, rid, ip):
    with self.req_lock:
      r = self.requests.get(rid)
      if r is None or r["ip"] != ip:
        return "error"
      return r["status"]

  def take_request(self, rid, ip):
    """上傳開始時呼叫：只有「已批准、還沒用過、同一個 IP」的請求能通過，而且只能用一次。"""
    with self.req_lock:
      r = self.requests.get(rid)
      if r and r["status"] == "approved" and not r["used"] and r["ip"] == ip:
        r["used"] = True
        return dict(r)
    return None

  # ----- 傳輸中的檔案 -----
  def begin_transfer(self, ip, name, size):
    tid = secrets.token_hex(6)
    now = time.time()
    with self.xfer_lock:
      self.transfers[tid] = {
          "ip": ip,
          "name": name,
          "size": size,
          "done": 0,
          "speed": 0.0,
          "d_last": 0,
          "t_last": now,
      }
    return tid

  def xfer_progress(self, tid, done):
    t = self.transfers.get(tid)
    if t is not None:
      t["done"] = done

  def end_transfer(self, tid):
    with self.xfer_lock:
      self.transfers.pop(tid, None)

  # ---------- 畫面 ----------
  def start(self):
    app = self.app
    app.clear_container()
    c = app.container

    ttk.Label(c, text="上傳伺服器", style="H2.TLabel").pack(pady=(6, 2))
    ttk.Label(
        c,
        text=f"收到的檔案會放在：{DOWNLOAD_DIR}",
        style="Hint.TLabel",
    ).pack()

    # 開關列
    top = ttk.Frame(c)
    top.pack(fill="x", pady=(8, 2), padx=8)
    ttk.Label(top, text="連接埠：").pack(side="left")
    self.port_entry = ttk.Entry(top, textvariable=self.port_var, width=7)
    self.port_entry.pack(side="left", padx=(0, 8))
    self.btn_toggle = ttk.Button(
        top, text="開啟伺服器", style="Accent.TButton", command=self.toggle_server
    )
    self.btn_toggle.pack(side="left", padx=4)
    ttk.Button(top, text="返回主選單", command=self.back_home).pack(side="right")

    # 審核模式列
    mrow = ttk.Frame(c)
    mrow.pack(fill="x", padx=8, pady=2)
    ttk.Label(mrow, text="審核模式：").pack(side="left")
    self.mode_var.set(self.approval_mode)
    ttk.Radiobutton(
        mrow, text="一律允許", value="allow", variable=self.mode_var,
        command=self.on_mode_change,
    ).pack(side="left", padx=4)
    ttk.Radiobutton(
        mrow, text="手動審核", value="manual", variable=self.mode_var,
        command=self.on_mode_change,
    ).pack(side="left", padx=4)

    # 動態密碼列
    prow = ttk.Frame(c)
    prow.pack(fill="x", padx=8, pady=2)
    ttk.Label(prow, text="目前上傳密碼：").pack(side="left")
    ttk.Label(
        prow, textvariable=self.code_var, foreground=PAL["accent"],
        font=("Consolas", 18, "bold"),
    ).pack(side="left", padx=6)
    ttk.Button(prow, text="複製", command=self.copy_code).pack(side="left", padx=2)
    ttk.Label(prow, textvariable=self.pw_left_var, style="Hint.TLabel").pack(
        side="left", padx=8
    )
    self.pw_bar = ttk.Progressbar(c, maximum=PW_SECONDS, mode="determinate")
    self.pw_bar.pack(fill="x", padx=8, pady=(0, 2))

    ttk.Label(c, textvariable=self.status_var).pack(anchor="w", padx=10)
    ttk.Label(c, textvariable=self.url_var, style="Hint.TLabel").pack(
        anchor="w", padx=10
    )

    # 等待審核
    ph = ttk.Frame(c)
    ph.pack(fill="x", padx=8, pady=(8, 2))
    ttk.Label(ph, text="等待審核", style="H2.TLabel").pack(side="left")
    ttk.Button(ph, text="全部允許", command=self.approve_all).pack(side="right", padx=2)
    ttk.Button(ph, text="拒絕", command=self.deny_selected).pack(side="right", padx=2)
    ttk.Button(ph, text="允許", style="Accent.TButton", command=self.approve_selected).pack(
        side="right", padx=2
    )
    self.pending_tree = ttk.Treeview(
        c, columns=("name", "size", "ip"), show="headings", height=3,
        selectmode="extended",
    )
    self.pending_tree.heading("name", text="檔名")
    self.pending_tree.heading("size", text="大小")
    self.pending_tree.heading("ip", text="來源 IP")
    self.pending_tree.column("name", width=sw(app, 320))
    self.pending_tree.column("size", width=sw(app, 90), anchor="e")
    self.pending_tree.column("ip", width=sw(app, 130))
    self.pending_tree.pack(fill="x", padx=8)

    # 傳輸中
    xh = ttk.Frame(c)
    xh.pack(fill="x", padx=8, pady=(8, 2))
    ttk.Label(xh, text="傳輸中", style="H2.TLabel").pack(side="left")
    ttk.Label(xh, textvariable=self.speed_var, style="Hint.TLabel").pack(
        side="left", padx=10
    )
    self.xfer_frame = ttk.Frame(c)
    self.xfer_frame.pack(fill="x", padx=8)
    self.rows = {}
    self.idle_label = ttk.Label(
        self.xfer_frame, text="目前沒有傳輸", style="Hint.TLabel"
    )
    self.idle_label.pack(anchor="w")

    # 下方分頁：已接收檔案 / Log
    nb = ttk.Notebook(c)
    nb.pack(fill="both", expand=True, padx=8, pady=(8, 8))

    file_tab = ttk.Frame(nb)
    nb.add(file_tab, text="已接收的檔案")
    fb = ttk.Frame(file_tab)
    fb.pack(fill="x", pady=4)
    ttk.Button(fb, text="開啟資料夾", command=self.open_folder).pack(side="right", padx=2)
    ttk.Button(fb, text="重新整理", command=self.refresh_files).pack(side="right", padx=2)
    tree_frame = ttk.Frame(file_tab)
    tree_frame.pack(fill="both", expand=True)
    self.tree = ttk.Treeview(
        tree_frame, columns=("name", "size", "time"), show="headings", height=6
    )
    self.tree.heading("name", text="檔名")
    self.tree.heading("size", text="大小")
    self.tree.heading("time", text="接收時間")
    self.tree.column("name", width=sw(app, 320))
    self.tree.column("size", width=sw(app, 90), anchor="e")
    self.tree.column("time", width=sw(app, 150))
    sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
    self.tree.configure(yscrollcommand=sb.set)
    self.tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")

    log_tab = ttk.Frame(nb)
    nb.add(log_tab, text="Log")
    lb = ttk.Frame(log_tab)
    lb.pack(fill="x", pady=4)
    ttk.Button(lb, text="清除畫面", command=self.clear_log_view).pack(side="right", padx=2)
    ttk.Button(lb, text="開啟 log 檔", command=self.open_log_file).pack(side="right", padx=2)
    log_frame = ttk.Frame(log_tab)
    log_frame.pack(fill="both", expand=True)
    self.text = tk.Text(
        log_frame,
        height=8,
        wrap="word",
        state="disabled",
        bg=PAL["card"],
        fg=PAL["fg"],
        relief="flat",
    )
    lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.text.yview)
    self.text.configure(yscrollcommand=lsb.set)
    self.text.pack(side="left", fill="both", expand=True)
    lsb.pack(side="right", fill="y")

    # 把之前的 log 補回畫面
    self.text.configure(state="normal")
    self.text.insert("end", "\n".join(self.log_lines) + ("\n" if self.log_lines else ""))
    self.text.configure(state="disabled")
    self.text.see("end")

    self.files_dirty = True
    self._pending_shown = ()
    self.update_toggle_ui()
    self.refresh_files()

    if self._poll_id is not None:
      try:
        self.app.after_cancel(self._poll_id)
      except Exception:
        pass
    self._poll()

  def back_home(self):
    # 伺服器保持運作，只是回到主頁。要停請先按「關閉伺服器」。
    self.app.show_main_home()

  # ---------- 定時檢查（在主執行緒更新畫面） ----------
  def _widgets_alive(self):
    try:
      return self.text is not None and self.text.winfo_exists()
    except tk.TclError:
      return False

  def _poll(self):
    # 1. log
    new_lines = []
    while True:
      try:
        new_lines.append(self.log_queue.get_nowait())
      except queue.Empty:
        break
    if new_lines:
      self.log_lines.extend(new_lines)
      del self.log_lines[:-2000]  # 畫面只保留最近 2000 行（log 檔是完整的）
      if self._widgets_alive():
        self.text.configure(state="normal")
        self.text.insert("end", "\n".join(new_lines) + "\n")
        self.text.configure(state="disabled")
        self.text.see("end")

    # 2、3. 資料與畫面更新。就算某一步出錯，也不能讓整個定時更新停掉
    try:
      # 資料更新（不管有沒有在這個頁面都要做）
      self._expire_requests()
      self._update_speeds()

      # 畫面更新（只有在這個頁面才做）
      if self._widgets_alive():
        self._refresh_code()
        self._refresh_pending()
        self._refresh_transfers()
        if self.files_dirty:
          self.refresh_files()
    except Exception as e:
      err = f"{type(e).__name__}: {e}"
      if err != getattr(self, "_last_ui_err", None):  # 同一個錯誤只記一次，避免洗版
        self._last_ui_err = err
        self.log(f"介面更新發生錯誤：{err}")

    self._poll_id = self.app.after(500, self._poll)

  def _expire_requests(self):
    now = time.time()
    expired = []
    with self.req_lock:
      for rid, r in list(self.requests.items()):
        if r["status"] == "pending" and now - r["created"] > APPROVAL_TIMEOUT:
          r["status"] = "timeout"
          expired.append(r)
        if now - r["created"] > 600 and r["status"] != "pending":
          del self.requests[rid]
    for r in expired:
      self.log(f"審核逾時，自動拒絕：{r['name']}", r["ip"])

  def _update_speeds(self):
    now = time.time()
    with self.xfer_lock:
      for t in self.transfers.values():
        dt = now - t["t_last"]
        if dt >= 0.4:
          inst = (t["done"] - t["d_last"]) / dt
          t["speed"] = inst if t["speed"] == 0 else 0.5 * t["speed"] + 0.5 * inst
          t["d_last"] = t["done"]
          t["t_last"] = now

  def _refresh_code(self):
    code, left = self.get_code()
    if code is None:
      self.code_var.set("－－－－ －－－－")
      self.pw_left_var.set("")
      self.pw_bar["value"] = 0
    else:
      self.code_var.set(f"{code[:4]} {code[4:]}")
      self.pw_left_var.set(f"{math.ceil(left)} 秒後更換")
      self.pw_bar["value"] = left

  def _refresh_pending(self):
    with self.req_lock:
      items = sorted(
          ((rid, dict(r)) for rid, r in self.requests.items() if r["status"] == "pending"),
          key=lambda x: x[1]["created"],
      )
    ids = tuple(rid for rid, _ in items)
    # _pending_shown 是 None 代表「強制重畫」（按完允許／拒絕後會這樣設）
    if self._pending_shown is not None and ids == self._pending_shown:
      return
    new_ids = set(ids) - set(self._pending_shown or ())
    for item in self.pending_tree.get_children():
      self.pending_tree.delete(item)
    for rid, r in items:
      self.pending_tree.insert(
          "", "end", iid=rid, values=(r["name"], fmt_size(r["size"]), r["ip"])
      )
    self._pending_shown = ids
    if new_ids:
      try:
        self.app.bell()  # 有新請求時叫一聲提醒
      except Exception:
        pass

  def _refresh_transfers(self):
    with self.xfer_lock:
      snap = {tid: dict(t) for tid, t in self.transfers.items()}

    for tid in list(self.rows):
      if tid not in snap:
        self.rows.pop(tid)["frame"].destroy()

    total_speed = 0.0
    for tid, t in snap.items():
      row = self.rows.get(tid)
      if row is None:
        frame = ttk.Frame(self.xfer_frame)
        name_lbl = ttk.Label(frame, text=f"{t['name']}　←　{t['ip']}")
        name_lbl.pack(anchor="w")
        bar = ttk.Progressbar(frame, maximum=100, mode="determinate")
        bar.pack(fill="x")
        info = ttk.Label(frame, text="", style="Hint.TLabel")
        info.pack(anchor="w")
        frame.pack(fill="x", pady=2)
        row = {"frame": frame, "bar": bar, "info": info}
        self.rows[tid] = row
      pct = t["done"] / t["size"] * 100 if t["size"] else 0
      row["bar"]["value"] = pct
      row["info"].configure(
          text=f"{pct:.0f}%　{fmt_size(t['done'])} / {fmt_size(t['size'])}　{fmt_speed(t['speed'])}"
      )
      total_speed += t["speed"]

    if self.rows:
      self.idle_label.pack_forget()
      self.speed_var.set(f"總速度：{fmt_speed(total_speed)}　傳輸中：{len(self.rows)} 個")
    else:
      if not self.idle_label.winfo_ismapped():
        self.idle_label.pack(anchor="w")
      self.speed_var.set("")

  # ---------- 按鈕功能 ----------
  def on_mode_change(self):
    self.approval_mode = self.mode_var.get()
    name = "一律允許" if self.approval_mode == "allow" else "手動審核"
    self.log(f"審核模式改為：{name}")

  def copy_code(self):
    code, _ = self.get_code()
    if code is None:
      messagebox.showinfo("沒有密碼", "請先開啟伺服器。", parent=self.app)
      return
    self.app.clipboard_clear()
    self.app.clipboard_append(code)

  def _set_request_status(self, rid, status):
    with self.req_lock:
      r = self.requests.get(rid)
      if r is None or r["status"] != "pending":
        return
      r["status"] = status
      name, ip = r["name"], r["ip"]
    if status == "approved":
      self.log(f"使用者允許上傳：{name}", ip)
    else:
      self.log(f"使用者拒絕上傳：{name}", ip)

  def approve_selected(self):
    for rid in self.pending_tree.selection():
      self._set_request_status(rid, "approved")
    self._pending_shown = None  # 強制下一輪重畫

  def deny_selected(self):
    for rid in self.pending_tree.selection():
      self._set_request_status(rid, "denied")
    self._pending_shown = None

  def approve_all(self):
    for rid in self.pending_tree.get_children():
      self._set_request_status(rid, "approved")
    self._pending_shown = None

  def update_toggle_ui(self):
    running = self.server is not None
    if self.btn_toggle is not None:
      self.btn_toggle.configure(text="關閉伺服器" if running else "開啟伺服器")
    if self.port_entry is not None:
      self.port_entry.configure(state="disabled" if running else "normal")
    if running:
      port = self.server.server_address[1]
      self.status_var.set(f"伺服器：運作中（連接埠 {port}）")
      self.url_var.set(
          f"本機：http://127.0.0.1:{port}    同一個網路：http://{get_local_ip()}:{port}"
      )
    else:
      self.status_var.set("伺服器：已關閉")
      self.url_var.set("")

  def toggle_server(self):
    if self.server is None:
      self.start_server()
    else:
      self.stop_server()

  def start_server(self):
    try:
      port = int(self.port_var.get().strip())
      if not 1 <= port <= 65535:
        raise ValueError
    except ValueError:
      messagebox.showwarning("連接埠錯誤", "請輸入 1～65535 的數字。", parent=self.app)
      return
    try:
      os.makedirs(DOWNLOAD_DIR, exist_ok=True)
      srv = ThreadingHTTPServer(("0.0.0.0", port), make_handler(self))
    except OSError as e:
      messagebox.showerror(
          "無法開啟",
          f"連接埠 {port} 可能已被其他程式使用。\n\n{e}",
          parent=self.app,
      )
      return
    srv.daemon_threads = True
    with self.pw_lock:  # 每次開啟都產生全新的密碼
      self.code = self._new_code()
      self.prev_code = None
      self.pw_expire = time.time() + PW_SECONDS
    self.server = srv
    self.thread = threading.Thread(target=srv.serve_forever, daemon=True)
    self.thread.start()
    self.log(f"伺服器已開啟，連接埠 {port}")
    self.update_toggle_ui()

  def stop_server(self):
    srv = self.server
    if srv is None:
      return
    self.server = None  # 傳輸中的檔案看到這個會自己中止
    with self.pw_lock:
      self.code = None
      self.prev_code = None
    with self.req_lock:  # 還在等審核的請求全部作廢
      for r in self.requests.values():
        if r["status"] == "pending":
          r["status"] = "closed"
    try:
      srv.shutdown()
      srv.server_close()
    except Exception as e:
      self.log(f"關閉時發生錯誤：{type(e).__name__}: {e}")
    self.log("伺服器已關閉")
    self.update_toggle_ui()

  def refresh_files(self):
    self.files_dirty = False
    if self.tree is None:
      return
    try:
      if not self.tree.winfo_exists():
        return
    except tk.TclError:
      return
    for item in self.tree.get_children():
      self.tree.delete(item)
    if not os.path.isdir(DOWNLOAD_DIR):
      return
    rows = []
    for name in os.listdir(DOWNLOAD_DIR):
      if name.endswith(".part"):
        continue
      p = os.path.join(DOWNLOAD_DIR, name)
      if os.path.isfile(p):
        st = os.stat(p)
        rows.append((st.st_mtime, name, st.st_size))
    rows.sort(reverse=True)
    for mtime, name, size in rows:
      self.tree.insert(
          "",
          "end",
          values=(
              name,
              fmt_size(size),
              datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S"),
          ),
      )

  def open_folder(self):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    self._open_path(DOWNLOAD_DIR)

  def open_log_file(self):
    if not os.path.exists(LOG_PATH):
      messagebox.showinfo("沒有 log", "目前還沒有 log 檔。", parent=self.app)
      return
    self._open_path(LOG_PATH)

  def _open_path(self, path):
    try:
      if sys.platform.startswith("win"):
        os.startfile(path)
      elif sys.platform == "darwin":
        os.system(f'open "{path}"')
      else:
        os.system(f'xdg-open "{path}"')
    except Exception as e:
      messagebox.showerror("無法開啟", str(e), parent=self.app)

  def clear_log_view(self):
    self.log_lines.clear()
    if self._widgets_alive():
      self.text.configure(state="normal")
      self.text.delete("1.0", "end")
      self.text.configure(state="disabled")


def sw(app, n):
  """依縮放換算像素；app 沒有 px 時就直接用原值。"""
  try:
    return app.px(n)
  except Exception:
    return n