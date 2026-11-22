"""共用小視窗與元件：進度視窗、取值範圍設定、打勾清單、開啟檔案。"""

import os
import sys
import tkinter as tk
import subprocess
from tkinter import messagebox, ttk

from core import PAL, get_values, range_key, save_range_settings


def open_file(path):
  """用系統預設程式（Excel）開啟檔案。"""
  try:
    if sys.platform == "win32":
      os.startfile(path)
    elif sys.platform == "darwin":
      subprocess.Popen(["open", path])
    else:
      subprocess.Popen(["xdg-open", path])
    return True
  except Exception as e:
    messagebox.showerror("開啟失敗", f"無法開啟檔案：\n{path}\n\n{e}")
    return False


class ProgressWindow(tk.Toplevel):
  """讀取檔案時的進度視窗（可按取消）。"""

  def __init__(self, app, parent, title, total):
    super().__init__(parent)
    self.title(title)
    self.cancelled = False
    self.resizable(False, False)
    bg = getattr(app, "theme_bg", None)
    if bg:
      self.configure(bg=bg)
    body = ttk.Frame(self, padding=16)
    body.pack(fill="both", expand=True)
    self.lbl = ttk.Label(body, text="準備中…", width=56)
    self.lbl.pack(anchor="w")
    self.bar = ttk.Progressbar(body, maximum=max(total, 1), length=440)
    self.bar.pack(pady=10)
    ttk.Button(body, text="取消", command=self._cancel).pack()
    self.protocol("WM_DELETE_WINDOW", self._cancel)
    self.update_idletasks()
    top = parent.winfo_toplevel()
    w, h = self.winfo_reqwidth(), self.winfo_reqheight()
    x = top.winfo_rootx() + (top.winfo_width() - w) // 2
    y = top.winfo_rooty() + (top.winfo_height() - h) // 2
    self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    self.transient(top)
    self.grab_set()
    self.update()

  def set_text(self, text):
    if len(text) > 56:
      text = text[:53] + "…"
    self.lbl.config(text=text)
    self.update()

  def step(self):
    self.bar["value"] = self.bar["value"] + 1
    self.update()

  def _cancel(self):
    self.cancelled = True

  def close(self):
    try:
      self.grab_release()
      self.destroy()
    except tk.TclError:
      pass


class RangeDialog(tk.Toplevel):
  """每個「檔案＋工作表」各自設定取值範圍；可多選一起編輯，並即時預覽結果。"""

  COLS = (
      ("date", "日期", 90),
      ("file", "檔案", 170),
      ("sheet", "工作表", 100),
      ("range", "目前範圍", 230),
      ("n", "有效筆數", 70),
      ("mean", "平均 A0", 80),
      ("std", "標準差", 65),
      ("drift", "前後半差", 75),
      ("rx", "Rx", 75),
      ("status", "狀態", 220),
  )

  def __init__(self, app, infos):
    super().__init__(app)
    self.app = app
    self.title("取值範圍設定（每個工作表可各自設定）")
    self.saved = False
    self.settings = {k: dict(v) for k, v in app.range_settings.items()}
    self.entries = []
    for info in infos:
      if info.get("error"):
        continue
      for sh in info["sheets"]:
        if sh["kind"] == "data":
          self.entries.append((info, sh, range_key(info["path"], sh["name"])))
    self.entries.sort(
        key=lambda e: (app.date_text_of(e[0]) or "9999", e[0]["filename"])
    )
    app.center_dialog(self, 1180, 660)
    self._build()
    self._fill()
    self.grab_set()

  def _build(self):
    body = ttk.Frame(self, padding=12)
    body.pack(fill="both", expand=True)
    ttk.Label(
        body,
        text=(
            "「沿用公式」＝照 Excel 裡 =AVERAGE(...) 的範圍；「自訂」＝改用第 N~M 筆"
            "有效數據（空白列不算）。\n可按住 Ctrl／Shift 選多列一起套用。"
            "紅字代表範圍有問題，該列數值會留空。「前後半差」是範圍內"
            "後半段平均減前半段平均（A0 單位），越小代表越穩定。"
        ),
        style="Hint.TLabel",
        justify="left",
    ).pack(anchor="w")

    filt = ttk.Frame(body)
    filt.pack(fill="x", pady=6)
    ttk.Label(filt, text="篩選（檔名或工作表名）：").pack(side="left")
    self.var_filter = tk.StringVar()
    ttk.Entry(filt, textvariable=self.var_filter, width=24).pack(
        side="left", padx=6
    )
    ttk.Button(filt, text="全選目前列表", command=self._select_all).pack(
        side="left", padx=6
    )

    box = ttk.Frame(body)
    box.pack(fill="both", expand=True)
    self.tree = ttk.Treeview(
        box,
        columns=[c[0] for c in self.COLS],
        show="headings",
        selectmode="extended",
    )
    for cid, text, width in self.COLS:
      self.tree.heading(cid, text=text)
      self.tree.column(cid, width=self.app.px(width), anchor="w")
    self.tree.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
    sb.pack(side="right", fill="y")
    self.tree.configure(yscrollcommand=sb.set)
    self.tree.tag_configure("warn", foreground=PAL["warn"])
    self.tree.tag_configure("empty", foreground=PAL["faint"])
    self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_select())

    ed = ttk.LabelFrame(body, text="編輯選取的列", padding=10)
    ed.pack(fill="x", pady=8)
    self.var_mode = tk.StringVar(value="formula")
    ttk.Radiobutton(
        ed, text="沿用 Excel 公式的範圍", value="formula", variable=self.var_mode
    ).grid(row=0, column=0, sticky="w", padx=4)
    ttk.Radiobutton(
        ed, text="自訂：第", value="custom", variable=self.var_mode
    ).grid(row=0, column=1, sticky="w", padx=(20, 2))
    self.ent_start = ttk.Entry(ed, width=8)
    self.ent_start.grid(row=0, column=2, padx=2)
    ttk.Label(ed, text="到第").grid(row=0, column=3, padx=2)
    self.ent_end = ttk.Entry(ed, width=8)
    self.ent_end.grid(row=0, column=4, padx=2)
    ttk.Label(ed, text="筆有效數據（空白列不算）").grid(row=0, column=5, padx=4)
    ttk.Button(
        ed, text="套用到選取列", style="Accent.TButton", command=self._apply
    ).grid(row=0, column=6, padx=(20, 4))
    ttk.Button(ed, text="清除選取列的設定", command=self._clear).grid(
        row=0, column=7, padx=4
    )

    bottom = ttk.Frame(body)
    bottom.pack(fill="x")
    ttk.Button(bottom, text="取消", command=self.destroy).pack(
        side="right", padx=4
    )
    ttk.Button(
        bottom, text="儲存並關閉", style="Accent.TButton", command=self._save
    ).pack(side="right", padx=4)

    # 篩選框要等表格建好之後才綁定
    self.var_filter.trace_add("write", lambda *a: self._fill())

  def _fill(self):
    keep = set(self.tree.selection())
    self.tree.delete(*self.tree.get_children())
    q = self.var_filter.get().strip().lower()

    def fmt(x, nd=2):
      return "" if x is None else f"{x:.{nd}f}"

    for i, (info, sh, key) in enumerate(self.entries):
      if q and q not in (info["filename"] + " " + sh["name"]).lower():
        continue
      st = self.settings.get(key)
      v = get_values(sh, st)
      custom = bool(st) and st.get("mode") == "custom"
      rng = v["method"]
      if not custom and v["win"]:
        rng += f"＝第{v['win'][0]}~{v['win'][1]}筆"
      no_value = v["rx"] is None and v["a0"] is None
      status = "；".join(v["warn"]) or v["note"]
      if not status and no_value:
        status = "沒有數值"
      if v["warn"]:
        tags = ("warn",)
      elif no_value:
        tags = ("empty",)
      else:
        tags = ()
      self.tree.insert(
          "",
          "end",
          iid=str(i),
          values=(
              self.app.date_text_of(info),
              info["filename"],
              sh["name"],
              rng,
              v["n"],
              fmt(v["mean"]),
              fmt(v["std"]),
              fmt(v["drift"]),
              fmt(v["rx"], 3),
              status,
          ),
          tags=tags,
      )
    still = [i for i in keep if self.tree.exists(i)]
    if still:
      self.tree.selection_set(still)

  def _select_all(self):
    self.tree.selection_set(self.tree.get_children())

  def _on_select(self):
    sel = self.tree.selection()
    if len(sel) != 1:
      return
    info, sh, key = self.entries[int(sel[0])]
    st = self.settings.get(key)
    custom = bool(st) and st.get("mode") == "custom"
    self.var_mode.set("custom" if custom else "formula")
    if custom:
      a, b = st["start"], st["end"]
    else:
      win = get_values(sh, None)["win"]
      if not win:
        return
      a, b = win
    self.ent_start.delete(0, tk.END)
    self.ent_start.insert(0, str(a))
    self.ent_end.delete(0, tk.END)
    self.ent_end.insert(0, str(b))

  def _apply(self):
    sel = self.tree.selection()
    if not sel:
      messagebox.showwarning("提示", "請先選取要編輯的列！", parent=self)
      return
    mode = self.var_mode.get()
    a = b = None
    if mode == "custom":
      try:
        a = int(self.ent_start.get().strip())
        b = int(self.ent_end.get().strip())
      except ValueError:
        messagebox.showerror(
            "格式錯誤", "起點與終點請輸入整數！", parent=self
        )
        return
      if a < 1 or b < a:
        messagebox.showerror(
            "範圍不合法", "起點要 ≥ 1，而且終點不能小於起點！", parent=self
        )
        return
    warns = []
    for iid in sel:
      info, sh, key = self.entries[int(iid)]
      if mode == "custom":
        self.settings[key] = {"mode": "custom", "start": a, "end": b}
      else:
        self.settings.pop(key, None)
      for w in get_values(sh, self.settings.get(key))["warn"]:
        warns.append(f"{info['filename']}／{sh['name']}：{w}")
    self._fill()
    if warns:
      more = f"\n...以及其他 {len(warns) - 12} 筆" if len(warns) > 12 else ""
      messagebox.showwarning(
          "取值範圍警告",
          "以下工作表的範圍有問題，數值會留空：\n\n" + "\n".join(warns[:12]) + more,
          parent=self,
      )

  def _clear(self):
    sel = self.tree.selection()
    if not sel:
      messagebox.showwarning("提示", "請先選取要清除設定的列！", parent=self)
      return
    for iid in sel:
      self.settings.pop(self.entries[int(iid)][2], None)
    self._fill()

  def _save(self):
    self.app.range_settings = self.settings
    if not save_range_settings(self.settings):
      messagebox.showwarning(
          "儲存失敗",
          "無法寫入 range_setting.json，設定只在這次執行期間有效。",
          parent=self,
      )
    self.saved = True
    self.destroy()


class CheckList(ttk.Frame):
  """打勾方塊清單（可捲動、全選／全不選）。"""

  def __init__(self, parent, title, height_px, fam, app=None):
    super().__init__(parent)
    self.vars = []
    self.fam = fam
    head = ttk.Frame(self)
    head.pack(fill="x")
    ttk.Label(head, text=title).pack(side="left")
    ttk.Button(head, text="全選", width=5, command=lambda: self.set_all(True)).pack(
        side="right"
    )
    ttk.Button(
        head, text="全不選", width=6, command=lambda: self.set_all(False)
    ).pack(side="right", padx=4)
    body = ttk.Frame(self)
    body.pack(fill="both", expand=True, pady=(2, 0))
    self.canvas = tk.Canvas(
        body, height=height_px, highlightthickness=1,
        highlightbackground=PAL["card_border"], bg=PAL["card"],
    )
    sb = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
    self.canvas.configure(yscrollcommand=sb.set)
    self.canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    self.inner = tk.Frame(self.canvas, bg=PAL["card"])
    self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
    self.inner.bind(
        "<Configure>",
        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
    )
    self.canvas.bind(
        "<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width)
    )

  def _on_wheel(self, event):
    self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

  def set_items(self, labels, checked=True):
    for w in self.inner.winfo_children():
      w.destroy()
    self.vars = []
    for text in labels:
      v = tk.BooleanVar(value=checked)
      cb = tk.Checkbutton(
          self.inner, text=text, variable=v, anchor="w", bg=PAL["card"],
          fg=PAL["fg"], activebackground=PAL["hover"], activeforeground=PAL["fg"],
          selectcolor=PAL["card"], font=(self.fam, 10), bd=0,
          highlightthickness=0,
      )
      cb.pack(fill="x", anchor="w")
      self.vars.append(v)
    self.canvas.yview_moveto(0)

  def set_all(self, value):
    for v in self.vars:
      v.set(value)

  def checked(self):
    return [i for i, v in enumerate(self.vars) if v.get()]
