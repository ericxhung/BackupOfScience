"""AI 分析：Gemini／OpenAI 格式 API、Markdown 渲染、AI 分頁（AIMixin）。"""

import json
import os
import re
import tkinter as tk
import openpyxl
import time
import csv
import io
import queue
import threading
import urllib.error
import urllib.request
from datetime import datetime
from tkinter import messagebox, ttk

from core import PAL, _disp_width, app_dir
from dialogs import CheckList, ProgressWindow


GEMINI_DEFAULT_MODEL = "gemini-3.5-flash-lite"
GEMINI_SETTING_FILE = "gemini_setting.json"
AI_SYSTEM = (
    "你是協助分析實驗／量測 Excel 數據的助手。使用者會附上幾個 Excel 檔轉成的 CSV，"
    "每段以「### 檔案：… ／ 工作表：…」開頭。請用繁體中文回答，"
    "引用數字時要說明來自哪個檔案與工作表；資料裡沒有的內容不要猜，直接說沒有。"
)


def load_gemini_setting():
  try:
    with open(os.path.join(app_dir(), GEMINI_SETTING_FILE), "r", encoding="utf-8") as f:
      d = json.load(f)
    return d if isinstance(d, dict) else {}
  except (OSError, ValueError):
    return {}


def save_gemini_setting(d):
  try:
    with open(os.path.join(app_dir(), GEMINI_SETTING_FILE), "w", encoding="utf-8") as f:
      json.dump(d, f, ensure_ascii=False, indent=2)
    return True
  except OSError:
    return False


def _csv_cell(v):
  if v is None:
    return ""
  if isinstance(v, float):
    return repr(round(v, 10)) if v == v else ""
  if isinstance(v, datetime):
    return v.strftime("%Y-%m-%d %H:%M:%S")
  return str(v)


def build_ai_csv(path, row_limit=None):
  """把一個 Excel 的每個工作表轉成 CSV 文字（用 Excel 存的計算結果）。回傳 (文字, 說明列表)。"""
  notes = []
  out = []
  try:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
  except Exception as e:
    return "", [f"{os.path.basename(path)} 讀取失敗：{e}"]
  try:
    for ws in wb.worksheets:
      buf = io.StringIO()
      w = csv.writer(buf, lineterminator="\n")
      n = 0
      cut = False
      for row in ws.iter_rows(values_only=True):
        vals = [_csv_cell(v) for v in row]
        while vals and vals[-1] == "":
          vals.pop()
        if not vals:
          continue
        if row_limit and n >= row_limit:
          cut = True
          break
        w.writerow(vals)
        n += 1
      if n == 0:
        continue
      head = f"### 檔案：{os.path.basename(path)} ／ 工作表：{ws.title}"
      if cut:
        head += f"（只取前 {row_limit} 列）"
        notes.append(f"{os.path.basename(path)}／{ws.title} 只取前 {row_limit} 列")
      out.append(head + "\n" + buf.getvalue())
  finally:
    wb.close()
  return "\n".join(out), notes


OPENAI_DEFAULT_URL = "https://api.openai.com/v1"
OPENAI_DEFAULT_MODEL = "gpt-4o-mini"
PROVIDERS = ("Gemini", "OpenAI 格式")


def openai_generate(api_key, base_url, model, system_text, history, timeout=180):
  """呼叫 OpenAI 格式（/chat/completions）的 API，也適用其他相容服務。"""
  url = (base_url.strip() or OPENAI_DEFAULT_URL).rstrip("/") + "/chat/completions"
  msgs = [{"role": "system", "content": system_text}]
  for r, t in history:
    msgs.append({"role": "assistant" if r == "model" else "user", "content": t})
  req = urllib.request.Request(
      url,
      data=json.dumps({"model": model, "messages": msgs}).encode("utf-8"),
      headers={
          "Content-Type": "application/json",
          "Authorization": "Bearer " + api_key,
      },
      method="POST",
  )
  try:
    with urllib.request.urlopen(req, timeout=timeout) as resp:
      data = json.loads(resp.read().decode("utf-8"))
  except urllib.error.HTTPError as e:
    try:
      err = json.loads(e.read().decode("utf-8")).get("error", {})
      detail = err.get("message", "") if isinstance(err, dict) else str(err)
    except Exception:
      detail = ""
    raise RuntimeError(f"HTTP {e.code}：{detail or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  choices = data.get("choices") or []
  if not choices:
    raise RuntimeError(f"沒有收到回答：{str(data)[:200]}")
  text = (choices[0].get("message") or {}).get("content") or ""
  if not text:
    text = f"（沒有文字內容，結束原因：{choices[0].get('finish_reason')}）"
  u = data.get("usage") or {}
  return text, {"totalTokenCount": u.get("total_tokens")}


def _sse_events(resp, stop):
  """逐行讀 Server-Sent Events，回傳每個 data: 後面的 JSON 字串。"""
  for raw in resp:
    if stop is not None and stop.is_set():
      return
    line = raw.decode("utf-8", "replace").strip()
    if line.startswith("data:"):
      payload = line[5:].strip()
      if payload and payload != "[DONE]":
        yield payload


def _http_detail(e):
  try:
    err = json.loads(e.read().decode("utf-8")).get("error", {})
    return err.get("message", "") if isinstance(err, dict) else str(err)
  except Exception:
    return ""


def gemini_stream(api_key, model, system_text, history, stop=None, timeout=180):
  """串流版 Gemini。逐段 yield ('text', 文字)，最後可能 yield ('usage', dict)。"""
  url = (
      "https://generativelanguage.googleapis.com/v1beta/models/"
      f"{model}:streamGenerateContent?alt=sse"
  )
  body = {
      "system_instruction": {"parts": [{"text": system_text}]},
      "contents": [{"role": r, "parts": [{"text": t}]} for r, t in history],
  }
  req = urllib.request.Request(
      url,
      data=json.dumps(body).encode("utf-8"),
      headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
      method="POST",
  )
  try:
    resp = urllib.request.urlopen(req, timeout=timeout)
  except urllib.error.HTTPError as e:
    raise RuntimeError(f"HTTP {e.code}：{_http_detail(e) or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  usage = {}
  with resp:
    for payload in _sse_events(resp, stop):
      try:
        data = json.loads(payload)
      except ValueError:
        continue
      if data.get("usageMetadata"):
        usage = data["usageMetadata"]
      cands = data.get("candidates") or []
      if cands:
        parts = (cands[0].get("content") or {}).get("parts") or []
        text = "".join(p.get("text", "") for p in parts)
        if text:
          yield ("text", text)
      elif data.get("promptFeedback"):
        raise RuntimeError(f"沒有收到回答（可能被安全機制擋下）：{data['promptFeedback']}")
  if usage:
    yield ("usage", usage)


def openai_stream(api_key, base_url, model, system_text, history, stop=None, timeout=180):
  """串流版 OpenAI 格式。yield 規則同 gemini_stream。"""
  url = (base_url.strip() or OPENAI_DEFAULT_URL).rstrip("/") + "/chat/completions"
  msgs = [{"role": "system", "content": system_text}]
  for r, t in history:
    msgs.append({"role": "assistant" if r == "model" else "user", "content": t})
  resp = None
  for with_usage in (True, False):  # 有些相容服務不支援 stream_options，失敗就退回不帶
    body = {"model": model, "messages": msgs, "stream": True}
    if with_usage:
      body["stream_options"] = {"include_usage": True}
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + api_key,
        },
        method="POST",
    )
    try:
      resp = urllib.request.urlopen(req, timeout=timeout)
      break
    except urllib.error.HTTPError as e:
      if with_usage and e.code == 400:
        continue
      raise RuntimeError(f"HTTP {e.code}：{_http_detail(e) or e.reason}")
    except urllib.error.URLError as e:
      raise RuntimeError(f"連線失敗：{e.reason}")
  usage = {}
  with resp:
    for payload in _sse_events(resp, stop):
      try:
        data = json.loads(payload)
      except ValueError:
        continue
      u = data.get("usage")
      if u:
        usage = {"totalTokenCount": u.get("total_tokens")}
      for ch in data.get("choices") or []:
        text = (ch.get("delta") or {}).get("content")
        if text:
          yield ("text", text)
  if usage:
    yield ("usage", usage)


def gemini_generate(api_key, model, system_text, history, timeout=180):
  """呼叫 Gemini REST API。history: [(role, text)]，role 為 'user' 或 'model'。回傳 (文字, 用量dict)。"""
  url = (
      "https://generativelanguage.googleapis.com/v1beta/models/"
      f"{model}:generateContent"
  )
  body = {
      "system_instruction": {"parts": [{"text": system_text}]},
      "contents": [{"role": r, "parts": [{"text": t}]} for r, t in history],
  }
  req = urllib.request.Request(
      url,
      data=json.dumps(body).encode("utf-8"),
      headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
      method="POST",
  )
  try:
    with urllib.request.urlopen(req, timeout=timeout) as resp:
      data = json.loads(resp.read().decode("utf-8"))
  except urllib.error.HTTPError as e:
    try:
      detail = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
    except Exception:
      detail = ""
    raise RuntimeError(f"HTTP {e.code}：{detail or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  cands = data.get("candidates") or []
  if not cands:
    fb = data.get("promptFeedback", {})
    raise RuntimeError(f"沒有收到回答（可能被安全機制擋下）：{fb}")
  parts = (cands[0].get("content") or {}).get("parts") or []
  text = "".join(p.get("text", "") for p in parts)
  if not text:
    text = f"（沒有文字內容，結束原因：{cands[0].get('finishReason')}）"
  return text, data.get("usageMetadata", {})



MD_INLINE = re.compile(
    r"\*\*(.+?)\*\*|`([^`]+)`|(?<![\*A-Za-z0-9])\*(?!\s)(.+?)(?<!\s)\*(?![\*A-Za-z0-9])"
    r"|\[([^\]]+)\]\(([^)]+)\)"
)


def md_setup_tags(txt, fam):
  mono = "Consolas"
  txt.tag_configure("md_h1", font=(fam, 15, "bold"), spacing1=8, spacing3=4)
  txt.tag_configure("md_h2", font=(fam, 13, "bold"), spacing1=6, spacing3=3)
  txt.tag_configure("md_h3", font=(fam, 11, "bold"), spacing1=4, spacing3=2)
  txt.tag_configure("md_b", font=(fam, 10, "bold"))
  txt.tag_configure("md_i", font=(fam, 10, "italic"))
  txt.tag_configure("md_code", font=(mono, 10), background=PAL["code_bg"])
  txt.tag_configure(
      "md_block", font=(mono, 10), background=PAL["block_bg"], lmargin1=14, lmargin2=14
  )
  txt.tag_configure("md_table", font=(mono, 10), lmargin1=8, lmargin2=8)
  txt.tag_configure("md_quote", foreground=PAL["muted"], lmargin1=14, lmargin2=14)
  txt.tag_configure("md_link", foreground=PAL["link"], underline=True)
  txt.tag_configure("md_hr", foreground=PAL["faint"])


def _md_inline(txt, text, base=()):
  pos = 0
  for m in MD_INLINE.finditer(text):
    if m.start() > pos:
      txt.insert(tk.END, text[pos:m.start()], base)
    if m.group(1) is not None:
      txt.insert(tk.END, m.group(1), base + ("md_b",))
    elif m.group(2) is not None:
      txt.insert(tk.END, m.group(2), base + ("md_code",))
    elif m.group(3) is not None:
      txt.insert(tk.END, m.group(3), base + ("md_i",))
    else:
      txt.insert(tk.END, m.group(4), base + ("md_link",))
      txt.insert(tk.END, f" ({m.group(5)})", base)
    pos = m.end()
  if pos < len(text):
    txt.insert(tk.END, text[pos:], base)


def _md_table(txt, rows):
  cols = max(len(r) for r in rows)
  rows = [r + [""] * (cols - len(r)) for r in rows]
  widths = [max(_disp_width(r[c]) for r in rows) for c in range(cols)]
  for ri, r in enumerate(rows):
    cells = [r[c] + " " * (widths[c] - _disp_width(r[c])) for c in range(cols)]
    txt.insert(tk.END, " │ ".join(cells) + "\n", ("md_table", "md_b") if ri == 0 else ("md_table",))
    if ri == 0:
      txt.insert(tk.END, "─┼─".join("─" * w for w in widths) + "\n", ("md_table",))


def render_markdown(txt, text):
  """把常見 Markdown（標題、粗斜體、程式碼、清單、引用、表格、連結）畫到 Text 元件。"""
  lines = text.replace("\r\n", "\n").split("\n")
  i = 0
  while i < len(lines):
    line = lines[i]
    st = line.strip()
    if st.startswith("```"):
      i += 1
      code = []
      while i < len(lines) and not lines[i].strip().startswith("```"):
        code.append(lines[i])
        i += 1
      i += 1
      txt.insert(tk.END, "\n".join(code) + "\n", ("md_block",))
      continue
    if st.startswith("|") and st.count("|") >= 2:
      rows = []
      while i < len(lines) and lines[i].strip().startswith("|"):
        cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
        if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
          rows.append(cells)
        i += 1
      if rows:
        _md_table(txt, rows)
      continue
    m = re.match(r"^(#{1,6})\s+(.*)$", st)
    if m:
      lvl = min(len(m.group(1)), 3)
      _md_inline(txt, m.group(2), (f"md_h{lvl}",))
      txt.insert(tk.END, "\n")
    elif re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", st):
      txt.insert(tk.END, "─" * 40 + "\n", ("md_hr",))
    elif re.match(r"^\s*[-*+]\s+", line):
      ind = len(line) - len(line.lstrip())
      body = re.sub(r"^\s*[-*+]\s+", "", line)
      txt.insert(tk.END, "　" * (ind // 2) + "• ")
      _md_inline(txt, body)
      txt.insert(tk.END, "\n")
    elif re.match(r"^\s*\d+[.)]\s+", line):
      ind = len(line) - len(line.lstrip())
      mm = re.match(r"^\s*(\d+[.)])\s+(.*)$", line)
      txt.insert(tk.END, "　" * (ind // 2) + mm.group(1) + " ")
      _md_inline(txt, mm.group(2))
      txt.insert(tk.END, "\n")
    elif st.startswith(">"):
      _md_inline(txt, st.lstrip("> ").strip(), ("md_quote",))
      txt.insert(tk.END, "\n")
    else:
      _md_inline(txt, line)
      txt.insert(tk.END, "\n")
    i += 1


class AIMixin:
  """AI 分析分頁（Gemini／OpenAI 格式）的畫面與邏輯，由 ViewPage 繼承使用。"""

  # ---------- AI 分析（Gemini） ----------
  def _build_ai_tab(self):
    t = self.tab_ai
    gs = load_gemini_setting()
    self.ai_history = []
    self.ai_data_text = ""
    self.ai_files = []
    self.ai_queue = queue.Queue()
    self.ai_busy = False
    self.ai_stop = threading.Event()
    self.stream_text = ""
    self.stream_usage = {}
    self._stream_dirty = False
    self._stream_t = 0.0

    def sub(name, defaults):
      d = gs.get(name)
      d = d if isinstance(d, dict) else {}
      out = dict(defaults)
      for k in defaults:
        if isinstance(d.get(k), str):
          out[k] = d[k]
      return out

    old_key = gs.get("api_key") if isinstance(gs.get("api_key"), str) else ""
    self.ai_cfg = {
        "Gemini": sub("gemini", {
            "api_key": old_key,
            "model": gs.get("model") if isinstance(gs.get("model"), str) and gs.get("model") else GEMINI_DEFAULT_MODEL,
        }),
        "OpenAI 格式": sub("openai", {
            "api_key": "", "model": OPENAI_DEFAULT_MODEL, "base_url": OPENAI_DEFAULT_URL,
        }),
    }
    prov = gs.get("provider") if gs.get("provider") in PROVIDERS else PROVIDERS[0]
    self.ai_prov = prov

    top = ttk.Frame(t)
    top.pack(fill="x")
    ttk.Label(top, text="服務").grid(row=0, column=0, sticky="w")
    self.cmb_prov = ttk.Combobox(top, values=PROVIDERS, state="readonly", width=12)
    self.cmb_prov.set(prov)
    self.cmb_prov.grid(row=0, column=1, padx=(4, 12))
    self.cmb_prov.bind("<<ComboboxSelected>>", self._ai_on_provider)
    ttk.Label(top, text="金鑰").grid(row=0, column=2, sticky="w")
    self.var_key = tk.StringVar()
    ttk.Entry(top, textvariable=self.var_key, show="*", width=30).grid(
        row=0, column=3, padx=(4, 12)
    )
    ttk.Label(top, text="模型").grid(row=0, column=4, sticky="w")
    self.var_model = tk.StringVar()
    ttk.Entry(top, textvariable=self.var_model, width=22).grid(row=0, column=5, padx=(4, 12))
    self.var_remember = tk.BooleanVar(value=bool(gs.get("remember", False)))
    ttk.Checkbutton(
        top, text="記住金鑰（明文存在 gemini_setting.json）", variable=self.var_remember
    ).grid(row=0, column=6)
    ttk.Label(top, text="網址").grid(row=1, column=0, sticky="w", pady=(4, 0))
    self.var_url = tk.StringVar()
    self.ent_url = ttk.Entry(top, textvariable=self.var_url, width=48)
    self.ent_url.grid(row=1, column=1, columnspan=5, sticky="w", padx=(4, 0), pady=(4, 0))
    ttk.Label(
        top, text="（OpenAI 格式才需要；相容服務可改這裡）", style="Hint.TLabel"
    ).grid(row=1, column=6, sticky="w", pady=(4, 0))
    self._ai_load_fields()

    body = ttk.Frame(t)
    body.pack(fill="both", expand=True, pady=(8, 0))
    left = ttk.Frame(body)
    left.pack(side="left", fill="y", padx=(0, 10))
    self.chk_ai = CheckList(left, "要餵給 AI 的 Excel（打勾）", self.app.px(300), self.fam)
    self.chk_ai.pack(fill="x")
    rf = ttk.Frame(left)
    rf.pack(fill="x", pady=(8, 0))
    ttk.Label(rf, text="每個工作表最多取前").pack(side="left")
    self.var_rows = tk.StringVar(value="500")
    ttk.Entry(rf, textvariable=self.var_rows, width=7).pack(side="left", padx=4)
    ttk.Label(rf, text="列（留空＝全部）").pack(side="left")
    ttk.Button(
        left, text="載入選取的 Excel 成 CSV", style="Accent.TButton",
        command=self._ai_load_data,
    ).pack(fill="x", pady=(8, 0))
    self.lbl_ai_data = ttk.Label(
        left, text="尚未載入資料", style="Hint.TLabel",
        wraplength=self.app.px(300), justify="left",
    )
    self.lbl_ai_data.pack(anchor="w", pady=(4, 0))

    right = ttk.Frame(body)
    right.pack(side="left", fill="both", expand=True)
    box = ttk.Frame(right)
    box.pack(fill="both", expand=True)
    self.txt_chat = tk.Text(
        box, wrap="word", state="disabled", font=(self.fam, 10), height=12,
        padx=8, pady=6, bg=PAL["card"], fg=PAL["fg"], insertbackground=PAL["fg"],
        relief="flat", highlightthickness=1, highlightbackground=PAL["card_border"],
    )
    sb = ttk.Scrollbar(box, orient="vertical", command=self.txt_chat.yview)
    self.txt_chat.configure(yscrollcommand=sb.set)
    self.txt_chat.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    md_setup_tags(self.txt_chat, self.fam)
    self.txt_chat.tag_configure("me", foreground=PAL["accent"], font=(self.fam, 10, "bold"))
    self.txt_chat.tag_configure("ai", foreground=PAL["ai"], font=(self.fam, 10, "bold"))
    self.txt_chat.tag_configure("sys", foreground=PAL["faint"])

    inp = ttk.Frame(right)
    inp.pack(fill="x", pady=(6, 0))
    self.txt_ask = tk.Text(
        inp, height=4, wrap="word", font=(self.fam, 10), padx=6, pady=4,
        bg=PAL["card"], fg=PAL["fg"], insertbackground=PAL["fg"], relief="flat",
        highlightthickness=1, highlightbackground=PAL["card_border"],
        highlightcolor=PAL["accent"],
    )
    self.txt_ask.pack(side="left", fill="x", expand=True)
    self.txt_ask.bind("<Return>", self._ai_enter)
    bf = ttk.Frame(inp)
    bf.pack(side="left", padx=(8, 0))
    self.btn_ask = ttk.Button(bf, text="送出", style="Accent.TButton", command=self._ai_send)
    self.btn_ask.pack(fill="x")
    self.btn_stop = ttk.Button(bf, text="停止", command=self._ai_stop, state="disabled")
    self.btn_stop.pack(fill="x", pady=(4, 0))
    ttk.Button(bf, text="清空對話", command=self._ai_clear).pack(fill="x", pady=(4, 0))
    self.lbl_ai = ttk.Label(
        right,
        text="Enter 送出、Shift+Enter 換行。送出時會把已載入的 CSV 與對話一起傳給 Google。",
        style="Hint.TLabel",
    )
    self.lbl_ai.pack(anchor="w", pady=(4, 0))

  def _ai_load_fields(self):
    c = self.ai_cfg[self.ai_prov]
    self.var_key.set(c["api_key"])
    self.var_model.set(c["model"])
    self.var_url.set(c.get("base_url", ""))
    self.ent_url.config(state="normal" if self.ai_prov == "OpenAI 格式" else "disabled")

  def _ai_store_fields(self):
    c = self.ai_cfg[self.ai_prov]
    c["api_key"] = self.var_key.get().strip()
    c["model"] = self.var_model.get().strip()
    if "base_url" in c:
      c["base_url"] = self.var_url.get().strip()

  def _ai_on_provider(self, event=None):
    self._ai_store_fields()
    self.ai_prov = self.cmb_prov.get()
    self._ai_load_fields()

  def _refresh_ai_files(self):
    infos = [i for i in self.infos.values() if not i.get("error")]
    infos.sort(key=lambda i: (self.app.date_text_of(i) or "9999", i["filename"]))
    self.ai_files = infos
    self.chk_ai.set_items(
        [f"{self.app.date_text_of(i) or '無日期'}　{i['filename']}" for i in infos],
        checked=False,
    )

  def _chat_add(self, who, text):
    tag = {"我": "me", "AI": "ai"}.get(who, "sys")
    self.txt_chat.config(state="normal")
    self.txt_chat.insert(tk.END, f"{who}：", tag)
    if who == "AI":
      self.txt_chat.insert(tk.END, "\n")
      render_markdown(self.txt_chat, text)
      self.txt_chat.insert(tk.END, "\n")
    else:
      self.txt_chat.insert(tk.END, text + "\n\n")
    self.txt_chat.config(state="disabled")
    self.txt_chat.see(tk.END)

  def _ai_clear(self):
    if self.ai_busy:
      return
    self.ai_history = []
    self.txt_chat.config(state="normal")
    self.txt_chat.delete("1.0", tk.END)
    self.txt_chat.config(state="disabled")

  def _ai_load_data(self):
    idx = self.chk_ai.checked()
    if not idx:
      messagebox.showwarning("提示", "請先勾選至少一個 Excel！")
      return
    raw = self.var_rows.get().strip()
    limit = None
    if raw:
      try:
        limit = int(raw)
        if limit < 1:
          raise ValueError
      except ValueError:
        messagebox.showerror("格式錯誤", "列數請輸入大於 0 的整數，或留空。")
        return
    prog = ProgressWindow(self.app, self.frame, "轉換成 CSV", len(idx))
    parts, notes = [], []
    try:
      for i in idx:
        info = self.ai_files[i]
        prog.set_text(info["filename"])
        text, nt = build_ai_csv(info["path"], limit)
        if text:
          parts.append(text)
        notes.extend(nt)
        prog.step()
        if prog.cancelled:
          return
    finally:
      prog.close()
    self.ai_data_text = "\n".join(parts)
    n = len(self.ai_data_text)
    msg = f"已載入 {len(parts)} 個檔案，約 {n:,} 個字元（約 {n // 3:,}～{n // 2:,} tokens）。"
    if notes:
      msg += f"\n{len(notes)} 個工作表被截斷。"
    if n > 1_500_000:
      msg += "\n⚠ 資料很大，可能超過模型上限或很花額度，建議減少檔案或列數。"
    self.lbl_ai_data.config(text=msg)
    self._ai_clear()
    self._chat_add("系統", "資料已載入，可以開始提問。（換資料會清空對話）")

  def _ai_enter(self, event):
    if event.state & 0x1:  # Shift
      return None
    self._ai_send()
    return "break"

  def _ai_send(self):
    if self.ai_busy:
      return
    q = self.txt_ask.get("1.0", tk.END).strip()
    if not q:
      return
    self._ai_store_fields()
    prov = self.ai_prov
    cfg = self.ai_cfg[prov]
    env = "GEMINI_API_KEY" if prov == "Gemini" else "OPENAI_API_KEY"
    key = cfg["api_key"] or os.environ.get(env, "").strip()
    model = cfg["model"] or (GEMINI_DEFAULT_MODEL if prov == "Gemini" else OPENAI_DEFAULT_MODEL)
    base_url = cfg.get("base_url", "")
    if not key:
      messagebox.showwarning("提示", f"請先輸入 {prov} 的 API 金鑰（或設定環境變數 {env}）。")
      return
    if not self.ai_data_text:
      if not messagebox.askyesno("尚未載入資料", "還沒載入任何 Excel，AI 看不到你的數據。\n仍要直接提問嗎？"):
        return
    rem = bool(self.var_remember.get())
    save_gemini_setting({
        "provider": prov,
        "remember": rem,
        "gemini": {
            "model": self.ai_cfg["Gemini"]["model"],
            "api_key": self.ai_cfg["Gemini"]["api_key"] if rem else "",
        },
        "openai": {
            "model": self.ai_cfg["OpenAI 格式"]["model"],
            "base_url": self.ai_cfg["OpenAI 格式"]["base_url"],
            "api_key": self.ai_cfg["OpenAI 格式"]["api_key"] if rem else "",
        },
    })
    self.txt_ask.delete("1.0", tk.END)
    self._chat_add("我", q)
    self.ai_history.append(("user", q))
    system = AI_SYSTEM + ("\n\n以下是數據：\n\n" + self.ai_data_text if self.ai_data_text else "")
    hist = list(self.ai_history)
    self.ai_busy = True
    self.ai_stop.clear()
    self.stream_text = ""
    self.stream_usage = {}
    self._stream_dirty = False
    self.btn_ask.config(state="disabled")
    self.btn_stop.config(state="normal")
    self.lbl_ai.config(text="AI 回答中…（可按「停止」）")
    self._stream_begin()
    stop = self.ai_stop

    def work():
      gen = None
      try:
        if prov == "Gemini":
          gen = gemini_stream(key, model, system, hist, stop)
        else:
          gen = openai_stream(key, base_url, model, system, hist, stop)
        for kind, val in gen:
          if stop.is_set():
            break
          self.ai_queue.put(("chunk" if kind == "text" else "usage", val))
        self.ai_queue.put(("stopped" if stop.is_set() else "done", None))
      except Exception as e:
        self.ai_queue.put(("err", str(e)))
      finally:
        if gen is not None:
          try:
            gen.close()
          except Exception:
            pass

    threading.Thread(target=work, daemon=True).start()
    self.frame.after(60, self._ai_poll)

  def _ai_stop(self):
    self.ai_stop.set()
    self.btn_stop.config(state="disabled")

  def _stream_begin(self):
    t = self.txt_chat
    t.config(state="normal")
    t.mark_set("ai_head", tk.END + "-1c")
    t.mark_gravity("ai_head", "left")
    t.insert(tk.END, "AI：", "ai")
    t.insert(tk.END, "\n")
    t.mark_set("ai_body", tk.END + "-1c")
    t.mark_gravity("ai_body", "left")
    t.config(state="disabled")

  def _stream_render(self):
    t = self.txt_chat
    t.config(state="normal")
    t.delete("ai_body", tk.END)
    render_markdown(t, self.stream_text)
    t.config(state="disabled")
    t.see(tk.END)
    self._stream_dirty = False
    self._stream_t = time.time()

  def _ai_poll(self):
    done = None
    try:
      while True:
        kind, payload = self.ai_queue.get_nowait()
        if kind == "chunk":
          self.stream_text += payload
          self._stream_dirty = True
        elif kind == "usage":
          self.stream_usage = payload
        else:
          done = (kind, payload)
          break
    except queue.Empty:
      pass
    try:
      if self._stream_dirty and (done or time.time() - self._stream_t >= 0.2):
        self._stream_render()
      if done is None:
        self.frame.after(60, self._ai_poll)
        return
      self.ai_busy = False
      self.btn_ask.config(state="normal")
      self.btn_stop.config(state="disabled")
      kind, payload = done
      text = self.stream_text
      if kind == "err":
        self.ai_history.pop()  # 失敗的問題不留在對話裡
        if not text:
          self.txt_chat.config(state="normal")
          self.txt_chat.delete("ai_head", tk.END)
          self.txt_chat.config(state="disabled")
        self._chat_add("系統", "發生錯誤：" + payload)
        self.lbl_ai.config(text="發生錯誤，請檢查金鑰、模型名稱或網路。")
        return
      if kind == "stopped" and not text:
        self.ai_history.pop()
        self.txt_chat.config(state="normal")
        self.txt_chat.delete("ai_head", tk.END)
        self.txt_chat.config(state="disabled")
        self.lbl_ai.config(text="已停止。")
        return
      self.ai_history.append(("model", text))
      self.txt_chat.config(state="normal")
      self.txt_chat.insert(tk.END, "\n")
      self.txt_chat.config(state="disabled")
      tok = self.stream_usage.get("totalTokenCount")
      head = "已停止（保留已收到的內容）。" if kind == "stopped" else "完成。"
      self.lbl_ai.config(text=head + (f"本次共用 {tok:,} tokens。" if tok else ""))
    except tk.TclError:
      pass
