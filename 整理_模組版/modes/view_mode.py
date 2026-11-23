"""模式：檢視（日曆／圖表／概況／趨勢比較／AI 分析）。"""

import json
import os
import re
import tkinter as tk
import calendar
import statistics
import tempfile
import time
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from core import app_dir, pd, EXCEL_EXTS, PAL, _disp_width, build_points, coord_to_index, get_values, listbox_colors, load_item_mapping, median_reading, natural_key, range_key
import charts
from charts import HAS_PLOTLY, LAYOUT_KEYS, LAYOUT_LABELS, PAL, build_plotly_html, draw_chart, natural_key, setup_mpl_fonts
from dialogs import CheckList, PAL, get_values, open_file, range_key
from ai_chat import AIMixin, CheckList, PAL, _disp_width
from modes.base import Mode


# ==================== 檢視模式（日曆／圖表／概況）====================
VIEW_FILE = "view_setting.json"

SRC_LABELS = [
    "Rx 電阻（自動判斷公式）",
    "A0 平均（自動判斷公式）",
    "自訂正則（比對公式文字）",
    "指定儲存格",
]
SRC_KEYS = ["rx", "a0", "regex", "cell"]
SRC_DEFAULT_PARAM = {"rx": "", "a0": "", "regex": r"^=\d+\*", "cell": "I3"}
SRC_HINT = {
    "rx": "依公式意思找 Rx，範圍沿用「⚙ 取值範圍設定」。",
    "a0": "依公式意思找 A0 平均，範圍沿用「⚙ 取值範圍設定」。",
    "regex": "比對公式文字，取第一個符合的格子（A~P 欄、1~30 列）。",
    "cell": "每張工作表都讀這一格的值（例：I3、J2）。",
}
Y_LABELS = {
    "rx": "Rx 電阻 (Ω)",
    "a0": "A0 平均 (ADC)",
    "regex": "數值（正則取值）",
    "cell": "數值（指定儲存格）",
}


def load_view_setting():
  try:
    with open(os.path.join(app_dir(), VIEW_FILE), "r", encoding="utf-8") as f:
      data = json.load(f)
    if isinstance(data, dict):
      return data
  except Exception:
    pass
  return {}


def save_view_setting(data):
  try:
    with open(os.path.join(app_dir(), VIEW_FILE), "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=2)
  except Exception:
    pass


def shorten(text, n):
  return text if len(text) <= n else text[: n - 1] + "…"


class ViewPage(AIMixin):
  """檢視模式：日曆（看哪天有哪些 Excel）、圖表、概況。"""

  def __init__(self, app):
    self.app = app
    self.fam = app.base_family
    self.frame = ttk.Frame(app.container)
    self.frame.pack(fill="both", expand=True)
    self.infos = {}
    self.date_map = {}
    self.undated = []
    self.chart_dates = []
    self.chart_items = []
    self.last_points = []
    self.month_list = []
    self.popup = None
    self.popup_owner = None
    self.hide_job = None
    s = load_view_setting()
    rd = s.get("root_dir", "")
    self.root_dir = rd if isinstance(rd, str) else ""
    ex = s.get("extra_files", [])
    self.extra_files = (
        [p for p in ex if isinstance(p, str)] if isinstance(ex, list) else []
    )
    now = datetime.now()
    self.cal_year, self.cal_month = now.year, now.month
    self._last_wheel = 0.0
    self._build_ui()
    self.frame.bind("<Destroy>", self._on_destroy)
    self.frame.bind_all("<MouseWheel>", self._on_global_wheel, add="+")
    if self.root_dir or self.extra_files:
      self.frame.after(100, self.rescan)
    else:
      self._set_status("請先按「選擇根資料夾」，程式會遞迴找出裡面所有 Excel。")

  # ---------- 共用 ----------
  def _on_destroy(self, event):
    if event.widget is self.frame:
      self._destroy_popup()

  def _on_global_wheel(self, event):
    """滑鼠滾輪：在打勾清單上＝捲動清單；在日曆上＝上滑前一個月、下滑下一個月。"""
    try:
      if not self.frame.winfo_exists():
        return
      w = self.frame.winfo_containing(event.x_root, event.y_root)
      while w is not None:
        if isinstance(w, CheckList):
          w._on_wheel(event)
          return
        if isinstance(w, (tk.Listbox, tk.Text, ttk.Treeview)):
          return
        if w is self.tab_cal:
          now = time.time()
          if now - self._last_wheel >= 0.25:  # 避免一下子滑過好幾個月
            self._last_wheel = now
            self._shift_month(-1 if event.delta > 0 else 1)
          return
        w = w.master
    except (tk.TclError, KeyError):
      pass

  def _set_status(self, text):
    self.lbl_status.config(text=text)

  def _save_setting(self):
    save_view_setting(
        {"root_dir": self.root_dir, "extra_files": self.extra_files}
    )

  def _build_ui(self):
    f = self.frame
    ttk.Label(f, text="【檢視】數據查看與圖表", style="H2.TLabel").pack(
        pady=(4, 6)
    )
    bar = ttk.Frame(f)
    bar.pack(fill="x", padx=12)
    for text, cmd in (
        ("📁 選擇根資料夾", self.choose_root),
        ("＋ 額外加入 Excel", self.add_extra),
        ("管理額外檔案", self.manage_extra),
        ("🔄 重新掃描", self.rescan),
        ("⚙ 取值範圍設定", self.open_range),
    ):
      ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=3)
    ttk.Button(bar, text="返回主選單", command=self.app.show_main_home).pack(
        side="right", padx=3
    )
    self.lbl_status = ttk.Label(f, text="", style="Hint.TLabel")
    self.lbl_status.pack(anchor="w", padx=14, pady=(6, 2))

    self.nb = ttk.Notebook(f)
    self.nb.pack(fill="both", expand=True, padx=10, pady=6)
    self.tab_cal = ttk.Frame(self.nb, padding=6)
    self.tab_chart = ttk.Frame(self.nb, padding=6)
    self.tab_over = ttk.Frame(self.nb, padding=6)
    self.tab_ai = ttk.Frame(self.nb, padding=6)
    self.tab_trend = ttk.Frame(self.nb, padding=6)
    self.nb.add(self.tab_cal, text="📅 日曆")
    self.nb.add(self.tab_chart, text="📈 圖表")
    self.nb.add(self.tab_over, text="📋 概況")
    self.nb.add(self.tab_trend, text="📊 趨勢比較")
    self.nb.add(self.tab_ai, text="🤖 AI 分析")
    self._build_calendar_tab()
    self._build_chart_tab()
    self._build_overview_tab()
    self._build_trend_tab()
    self._build_ai_tab()

  # ---------- 檔案來源 ----------
  def choose_root(self):
    d = filedialog.askdirectory(
        title="選擇存放 Excel 的根資料夾（會遞迴搜尋）",
        initialdir=self.root_dir or None,
    )
    if d:
      self.root_dir = d
      self._save_setting()
      self.rescan()

  def _add_extra_paths(self, files):
    have = {os.path.normcase(os.path.abspath(p)) for p in self.extra_files}
    added = 0
    for p in files:
      key = os.path.normcase(os.path.abspath(p))
      if p.lower().endswith(EXCEL_EXTS) and key not in have:
        self.extra_files.append(p)
        have.add(key)
        added += 1
    if added:
      self._save_setting()
    return added

  def add_extra(self):
    files = filedialog.askopenfilenames(
        title="額外加入 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")],
    )
    if files and self._add_extra_paths(files):
      self.rescan()

  def manage_extra(self):
    top = tk.Toplevel(self.app)
    top.title("管理額外加入的 Excel")
    self.app.center_dialog(top, 680, 400)
    top.grab_set()
    changed = [False]
    body = ttk.Frame(top, padding=12)
    body.pack(fill="both", expand=True)
    ttk.Label(body, text="這些檔案不在根資料夾裡，是額外加入的：").pack(
        anchor="w"
    )
    lb = tk.Listbox(
        body, selectmode=tk.EXTENDED, font=(self.fam, 10), activestyle="none",
        **listbox_colors()
    )
    lb.pack(fill="both", expand=True, pady=6)
    for p in self.extra_files:
      lb.insert(tk.END, p)

    def add():
      files = filedialog.askopenfilenames(
          parent=top,
          title="額外加入 Excel 檔案",
          filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")],
      )
      if files and self._add_extra_paths(files):
        changed[0] = True
        lb.delete(0, tk.END)
        for p in self.extra_files:
          lb.insert(tk.END, p)

    def remove():
      sel = list(lb.curselection())
      if not sel:
        messagebox.showwarning("提示", "請先選取要移除的檔案！", parent=top)
        return
      for i in reversed(sel):
        del self.extra_files[i]
        lb.delete(i)
      self._save_setting()
      changed[0] = True

    def close():
      top.destroy()
      if changed[0]:
        self.rescan()

    bf = ttk.Frame(body)
    bf.pack(fill="x")
    ttk.Button(bf, text="＋ 加入", command=add).pack(side="left", padx=3)
    ttk.Button(bf, text="移除選取", command=remove).pack(side="left", padx=3)
    ttk.Button(bf, text="關閉", style="Accent.TButton", command=close).pack(
        side="right", padx=3
    )
    top.protocol("WM_DELETE_WINDOW", close)

  def _collect_paths(self):
    paths = []
    if self.root_dir and os.path.isdir(self.root_dir):
      for root, _dirs, names in os.walk(self.root_dir):
        for n in names:
          if n.lower().endswith(EXCEL_EXTS) and not n.startswith("~$"):
            paths.append(os.path.join(root, n))
    for p in self.extra_files:
      if os.path.isfile(p):
        paths.append(p)
    seen = set()
    out = []
    for p in paths:
      key = os.path.normcase(os.path.abspath(p))
      if key not in seen:
        seen.add(key)
        out.append(p)
    out.sort(key=lambda p: os.path.basename(p).lower())
    return out

  def rescan(self):
    paths = self._collect_paths()
    if not paths:
      self.infos = {}
      self._refresh_all()
      if self.root_dir and not os.path.isdir(self.root_dir):
        self._set_status(f"找不到根資料夾：{self.root_dir}")
      else:
        self._set_status("沒有找到任何 Excel。請選擇根資料夾或額外加入檔案。")
      return
    infos = self.app.load_infos(paths, title="掃描 Excel 檔案中…")
    if infos is None:
      return
    self.infos = infos
    self._refresh_all()

  def open_range(self):
    paths = [p for p, i in self.infos.items() if not i.get("error")]
    if self.app.open_range_dialog(paths):
      self._refresh_overview()
      self._trend_update()

  def _refresh_all(self):
    self.date_map = {}
    self.undated = []
    for info in self.infos.values():
      d = "" if info.get("error") else self.app.date_text_of(info)
      if d:
        self.date_map.setdefault(d, []).append(info)
      else:
        self.undated.append(info)
    for lst in self.date_map.values():
      lst.sort(key=lambda i: i["filename"])
    self.undated.sort(key=lambda i: i["filename"])

    months = sorted({d[:7] for d in self.date_map}, reverse=True)
    self.month_list = months
    self.cmb_month["values"] = [
        f"{ym[:4]}年{int(ym[5:])}月（{sum(1 for d in self.date_map if d.startswith(ym))} 天）"
        for ym in months
    ]
    cur = f"{self.cal_year:04d}-{self.cal_month:02d}"
    if months and cur not in months:
      self.cal_year, self.cal_month = int(months[0][:4]), int(months[0][5:])

    n_dated = sum(len(v) for v in self.date_map.values())
    n_err = sum(1 for i in self.infos.values() if i.get("error"))
    where = f"根資料夾：{self.root_dir}　" if self.root_dir else ""
    self._set_status(
        f"{where}共 {len(self.infos)} 個 Excel（有日期 {n_dated}、"
        f"無日期 {len(self.undated) - n_err}、讀取失敗 {n_err}）"
    )
    self._render_calendar()
    self._fill_chart_lists()
    self._refresh_overview()
    self._refresh_trend_choices()
    self._refresh_ai_files()

  # ---------- 日曆 ----------
  def _build_calendar_tab(self):
    t = self.tab_cal
    nav = ttk.Frame(t)
    nav.pack(fill="x", pady=(0, 4))
    ttk.Button(nav, text="◀", width=3, command=lambda: self._shift_month(-1)).pack(
        side="left"
    )
    self.lbl_month = ttk.Label(nav, text="", style="H2.TLabel", width=12, anchor="center")
    self.lbl_month.pack(side="left", padx=8)
    ttk.Button(nav, text="▶", width=3, command=lambda: self._shift_month(1)).pack(
        side="left"
    )
    ttk.Button(nav, text="今天", command=self._goto_today).pack(
        side="left", padx=(10, 0)
    )
    ttk.Label(nav, text="跳到有資料的月份：").pack(side="left", padx=(20, 4))
    self.cmb_month = ttk.Combobox(nav, state="readonly", width=18)
    self.cmb_month.pack(side="left")
    self.cmb_month.bind("<<ComboboxSelected>>", self._on_pick_month)
    ttk.Label(
        t,
        text="雙擊 Excel 名稱可用 Excel 開啟；滑鼠移到有資料的日期上會浮出完整列表；在日曆上滾輪往上＝上個月、往下＝下個月。",
        style="Hint.TLabel",
    ).pack(anchor="w", pady=(0, 4))
    self.cal_holder = ttk.Frame(t)
    self.cal_holder.pack(fill="both", expand=True)
    self.box_undated = ttk.LabelFrame(
        t, text="無日期或讀取失敗的檔案（雙擊開啟）", padding=4
    )
    self.lst_undated = tk.Listbox(
        self.box_undated, height=3, font=(self.fam, 9), activestyle="none",
        **listbox_colors()
    )
    self.lst_undated.pack(fill="x")
    self.lst_undated.bind("<Double-Button-1>", self._open_undated)
    self._undated_paths = []

  def _shift_month(self, delta):
    m = self.cal_month + delta
    y = self.cal_year
    if m < 1:
      m, y = 12, y - 1
    elif m > 12:
      m, y = 1, y + 1
    self.cal_year, self.cal_month = y, m
    self._render_calendar()

  def _goto_today(self):
    now = datetime.now()
    self.cal_year, self.cal_month = now.year, now.month
    self._render_calendar()

  def _on_pick_month(self, event=None):
    i = self.cmb_month.current()
    if 0 <= i < len(self.month_list):
      ym = self.month_list[i]
      self.cal_year, self.cal_month = int(ym[:4]), int(ym[5:])
      self._render_calendar()

  def _open_undated(self, event=None):
    sel = self.lst_undated.curselection()
    if sel and sel[0] < len(self._undated_paths):
      open_file(self._undated_paths[sel[0]])

  def _link(self, parent, text, bg, path, size=9, bold=False):
    lbl = tk.Label(
        parent,
        text=text,
        fg=PAL["link"],
        bg=bg,
        cursor="hand2",
        anchor="w",
        font=(self.fam, size, "bold" if bold else "normal"),
    )
    lbl.bind("<Double-Button-1>", lambda e, p=path: open_file(p))
    return lbl

  def _render_calendar(self):
    for w in self.cal_holder.winfo_children():
      w.destroy()
    self._destroy_popup()
    y, m = self.cal_year, self.cal_month
    self.lbl_month.config(text=f"{y} 年 {m} 月")
    ym = f"{y:04d}-{m:02d}"
    if ym in self.month_list:
      self.cmb_month.current(self.month_list.index(ym))
    else:
      self.cmb_month.set("")

    grid = tk.Frame(self.cal_holder, bg=self.app.theme_bg)
    grid.pack(fill="both", expand=True)
    for c, name in enumerate(["日", "一", "二", "三", "四", "五", "六"]):
      tk.Label(
          grid, text=name, bg=self.app.theme_bg, fg=PAL["fg"],
          font=(self.fam, 10, "bold"),
      ).grid(row=0, column=c, sticky="nsew", pady=(0, 2))
      grid.columnconfigure(c, weight=1, uniform="day")
    weeks = calendar.Calendar(firstweekday=6).monthdayscalendar(y, m)
    today = datetime.now().strftime("%Y-%m-%d")
    for r, week in enumerate(weeks):
      grid.rowconfigure(r + 1, weight=1, minsize=self.app.px(90))
      for c, day in enumerate(week):
        if day == 0:
          continue
        date_str = f"{y:04d}-{m:02d}-{day:02d}"
        infos = self.date_map.get(date_str, [])
        has = bool(infos)
        bg = PAL["has"] if has else PAL["card"]
        is_today = date_str == today
        cell = tk.Frame(
            grid,
            bg=bg,
            highlightthickness=2 if is_today else 1,
            highlightbackground=PAL["accent"] if is_today else PAL["card_border"],
        )
        cell.grid(row=r + 1, column=c, sticky="nsew", padx=1, pady=1)
        cell.pack_propagate(False)
        tk.Label(
            cell,
            text=str(day),
            bg=bg,
            fg=PAL["fg"] if has else PAL["muted"],
            font=(self.fam, 10, "bold" if has else "normal"),
            anchor="w",
        ).pack(anchor="w", padx=4, pady=(2, 0))
        for info in infos[:2]:
          self._link(
              cell, "📄 " + shorten(info["filename"], 17), bg, info["path"]
          ).pack(anchor="w", padx=4)
        if len(infos) > 2:
          tk.Label(
              cell,
              text=f"+{len(infos) - 2}",
              bg=bg,
              fg=PAL["muted"],
              font=(self.fam, 9, "bold"),
              anchor="w",
          ).pack(anchor="w", padx=4)
        if has:
          self._bind_hover(cell, infos)

    # 無日期／讀取失敗的檔案
    self._undated_paths = []
    self.lst_undated.delete(0, tk.END)
    for info in self.undated:
      if info.get("error"):
        self.lst_undated.insert(tk.END, f"⚠ 讀取失敗：{info['filename']}（{info['error']}）")
      else:
        self.lst_undated.insert(tk.END, f"（無日期）{info['filename']}")
      self._undated_paths.append(info["path"])
    if self.undated:
      self.box_undated.pack(fill="x", pady=(6, 0))
    else:
      self.box_undated.pack_forget()

  # ---------- 日曆：懸浮列表 ----------
  def _all_widgets(self, w):
    out = [w]
    for c in w.winfo_children():
      out.extend(self._all_widgets(c))
    return out

  def _bind_hover(self, cell, infos):
    for w in self._all_widgets(cell):
      w.bind("<Enter>", lambda e, c=cell, i=infos: self._show_popup(c, i), add="+")
      w.bind("<Leave>", lambda e: self._schedule_hide(), add="+")

  def _cancel_hide(self):
    if self.hide_job is not None:
      try:
        self.frame.after_cancel(self.hide_job)
      except (tk.TclError, ValueError):
        pass
      self.hide_job = None

  def _schedule_hide(self):
    self._cancel_hide()
    try:
      self.hide_job = self.frame.after(250, self._hide_if_outside)
    except tk.TclError:
      self.hide_job = None

  def _hide_if_outside(self):
    self.hide_job = None
    if self.popup is None:
      return
    try:
      px, py = self.frame.winfo_pointerxy()
      w = self.frame.winfo_containing(px, py)
    except (tk.TclError, KeyError):
      w = None
    while w is not None:
      if w is self.popup or w is self.popup_owner:
        return
      w = getattr(w, "master", None)
    self._destroy_popup()

  def _destroy_popup(self):
    self._cancel_hide()
    if self.popup is not None:
      try:
        self.popup.destroy()
      except tk.TclError:
        pass
    self.popup = None
    self.popup_owner = None

  def _show_popup(self, cell, infos):
    self._cancel_hide()
    if self.popup is not None and self.popup_owner is cell:
      return
    self._destroy_popup()
    white = PAL["card"]
    pop = tk.Toplevel(self.frame)
    pop.overrideredirect(True)
    try:
      pop.attributes("-topmost", True)
    except tk.TclError:
      pass
    pop.configure(bg=PAL["popup_border"])
    date_text = self.app.date_text_of(infos[0])
    sw, sh = pop.winfo_screenwidth(), pop.winfo_screenheight()
    limit_h = sh - 90  # 預留工作列

    holder = {"inner": None}

    def build(max_files, max_items):
      if holder["inner"] is not None:
        holder["inner"].destroy()
      inner = tk.Frame(pop, bg=white, padx=10, pady=8)
      inner.pack(padx=1, pady=1)
      holder["inner"] = inner
      tk.Label(
          inner,
          text=f"{date_text}（{len(infos)} 個 Excel）",
          bg=white,
          fg=PAL["fg"],
          font=(self.fam, 10, "bold"),
          anchor="w",
      ).pack(fill="x")
      shown = infos if max_files is None else infos[:max_files]
      for info in shown:
        self._link(
            inner, "📄 " + info["filename"], white, info["path"], 10, True
        ).pack(anchor="w", pady=(8, 0))
        items = [x["name"] for x in info["sheets"] if x["kind"] == "data"]
        if max_items == 0:
          if items:
            tk.Label(
                inner, text=f"　共 {len(items)} 個項目", bg=white, fg=PAL["text2"],
                font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
        else:
          for name in items[:max_items]:
            tk.Label(
                inner, text="　• " + name, bg=white, fg=PAL["text2"],
                font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
          if len(items) > max_items:
            tk.Label(
                inner, text=f"　…另有 {len(items) - max_items} 項", bg=white,
                fg=PAL["muted"], font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
        t_txt, h_txt = median_reading(info)
        if t_txt:
          tk.Label(
              inner, text=f"　🌡 {t_txt}°C　💧 {h_txt}%", bg=white, fg=PAL["muted"],
              font=(self.fam, 9), anchor="w",
          ).pack(anchor="w")
      if max_files is not None and len(infos) > max_files:
        tk.Label(
            inner, text=f"…另有 {len(infos) - max_files} 個 Excel", bg=white,
            fg=PAL["muted"], font=(self.fam, 9), anchor="w",
        ).pack(anchor="w", pady=(6, 0))
      pop.update_idletasks()
      return pop.winfo_reqheight()

    # 內容太高放不下時，逐步縮減顯示的項目數／檔案數
    for mf, mi in ((None, 14), (None, 8), (None, 5), (None, 3), (None, 0),
                   (8, 0), (5, 0), (3, 0), (2, 0)):
      if build(mf, mi) <= limit_h:
        break
    for w in self._all_widgets(pop):
      w.bind("<Enter>", lambda e: self._cancel_hide(), add="+")
      w.bind("<Leave>", lambda e: self._schedule_hide(), add="+")
    pw, ph = pop.winfo_reqwidth(), pop.winfo_reqheight()
    x = cell.winfo_rootx() + cell.winfo_width() - 12
    y = cell.winfo_rooty() + 22
    if x + pw > sw - 10:
      x = max(10, cell.winfo_rootx() - pw + 12)
    if y + ph > sh - 70:  # 下方放不下就往上移
      y = max(10, sh - ph - 70)
    pop.geometry(f"+{x}+{y}")
    self.popup = pop
    self.popup_owner = cell

  # ---------- 圖表 ----------
  def _build_chart_tab(self):
    t = self.tab_chart
    left = ttk.Frame(t)
    left.pack(side="left", fill="y", padx=(0, 10))
    self.chart_left = left
    right = ttk.Frame(t)
    right.pack(side="left", fill="both", expand=True)
    self.btn_collapse = ttk.Button(
        right, text="◀ 隱藏設定（圖表加寬）", command=self._toggle_chart_panel
    )
    self.btn_collapse.pack(side="top", anchor="w", pady=(0, 4))

    bf = ttk.Frame(left)
    bf.pack(fill="x", pady=(0, 8))
    ttk.Button(
        bf, text="📈 畫圖", style="Accent.TButton", command=self.draw
    ).pack(side="left", padx=(0, 6))
    ttk.Button(bf, text="匯出圖上數據…", command=self.export_points).pack(side="left")
    ttk.Button(bf, text="📤 輸出整理報表…", command=self.export_report).pack(
        side="left", padx=(6, 0)
    )

    ttk.Label(left, text="軸組合").pack(anchor="w")
    self.cmb_layout = ttk.Combobox(
        left, values=LAYOUT_LABELS, state="readonly", width=40
    )
    self.cmb_layout.current(0)
    self.cmb_layout.pack(anchor="w", pady=(0, 6))

    ttk.Label(left, text="圖型").pack(anchor="w")
    tf = ttk.Frame(left)
    tf.pack(anchor="w", pady=(0, 4))
    self.var_ctype = tk.StringVar(value="scatter")
    for text, val in (("散佈圖", "scatter"), ("折線圖", "line"), ("長條圖", "bar")):
      ttk.Radiobutton(tf, text=text, value=val, variable=self.var_ctype).pack(
          side="left", padx=(0, 10)
      )
    self.var_connect = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        left, text="折線圖把各點連起來", variable=self.var_connect
    ).pack(anchor="w", pady=(0, 6))

    ttk.Label(left, text="顯示方式").pack(anchor="w")
    vf = ttk.Frame(left)
    vf.pack(anchor="w", pady=(0, 6))
    self.var_view = tk.StringVar(value="browser" if HAS_PLOTLY else "inline")
    ttk.Radiobutton(
        vf, text="瀏覽器互動圖", value="browser", variable=self.var_view
    ).pack(side="left", padx=(0, 10))
    ttk.Radiobutton(
        vf, text="視窗內（舊式）", value="inline", variable=self.var_view
    ).pack(side="left")
    self.plot_dir = None
    self.plot_files = []

    ttk.Label(left, text="y 軸取值方式").pack(anchor="w")
    self.cmb_src = ttk.Combobox(
        left, values=SRC_LABELS, state="readonly", width=40
    )
    self.cmb_src.current(0)
    self.cmb_src.pack(anchor="w")
    self.cmb_src.bind("<<ComboboxSelected>>", self._on_src_change)
    self.var_param = tk.StringVar(value="")
    self.ent_param = ttk.Entry(left, textvariable=self.var_param, width=42)
    self.ent_param.pack(anchor="w", pady=(4, 0))
    self.lbl_param = ttk.Label(
        left, text="", style="Hint.TLabel", wraplength=self.app.px(330), justify="left"
    )
    self.lbl_param.pack(anchor="w")

    self.var_map = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        left,
        text="套用 setting.json 的項目合併／順序",
        variable=self.var_map,
        command=self._fill_chart_lists,
    ).pack(anchor="w", pady=(6, 6))

    self.chk_dates = CheckList(left, "日期（打勾選取）", self.app.px(80), self.fam)
    self.chk_dates.pack(fill="x", pady=(0, 6))
    self.chk_items = CheckList(left, "項目（打勾選取）", self.app.px(130), self.fam)
    self.chk_items.pack(fill="x", pady=(0, 6))

    self.lbl_chart = ttk.Label(
        right, text="", style="Hint.TLabel", wraplength=self.app.px(640), justify="left"
    )
    self.lbl_chart.pack(side="bottom", anchor="w", pady=(4, 0))
    self.fig = None
    self.canvas = None
    if charts.ensure_mpl() and charts.HAS_MPL_TK:
      setup_mpl_fonts()
      self.fig = charts.Figure(figsize=(7, 5), dpi=100, facecolor=PAL["plot_bg"])
      self.canvas = charts.FigureCanvasTkAgg(self.fig, master=right)
      toolbar = charts.NavigationToolbar2Tk(self.canvas, right, pack_toolbar=False)
      toolbar.update()
      toolbar.pack(side="bottom", fill="x")
      self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
    else:
      ttk.Label(
          right,
          text="畫圖需要 matplotlib。請先在命令列執行：pip install matplotlib\n安裝後重新開啟程式即可。",
          justify="left",
      ).pack(expand=True)
    self._on_src_change()

  def _toggle_chart_panel(self):
    if self.chart_left.winfo_ismapped():
      self.chart_left.pack_forget()
      self.btn_collapse.config(text="▶ 顯示設定")
    else:
      self.chart_left.pack(side="left", fill="y", padx=(0, 10), before=self.btn_collapse.master)
      self.btn_collapse.config(text="◀ 隱藏設定（圖表加寬）")

  def _on_src_change(self, event=None):
    key = SRC_KEYS[self.cmb_src.current()]
    self.var_param.set(SRC_DEFAULT_PARAM[key])
    self.ent_param.config(state="normal" if key in ("regex", "cell") else "disabled")
    self.lbl_param.config(text=SRC_HINT[key])

  def _name_func(self):
    raw_to_final = load_item_mapping()[0] if self.var_map.get() else {}
    return lambda n: raw_to_final.get(n, n)

  def _fill_chart_lists(self):
    raw_to_final, order = load_item_mapping() if self.var_map.get() else ({}, [])
    names = []
    seen = set()
    for info in self.infos.values():
      if info.get("error"):
        continue
      for sh in info["sheets"]:
        if sh["kind"] == "data":
          final = raw_to_final.get(sh["name"], sh["name"])
          if final not in seen:
            seen.add(final)
            names.append(final)
    idx = {n: i for i, n in enumerate(order)}
    names.sort(key=lambda n: (idx.get(n, 10**6), natural_key(n)))
    self.chart_items = names
    self.chk_items.set_items(names)
    self.chart_dates = sorted(self.date_map.keys())
    self.chk_dates.set_items(
        [f"{d}（{len(self.date_map[d])} 個檔案）" for d in self.chart_dates]
    )

  def draw(self):
    browser = self.var_view.get() == "browser"
    if browser and not HAS_PLOTLY:
      messagebox.showinfo(
          "提示",
          "互動圖需要 plotly，請先在命令列執行：pip install plotly\n"
          "安裝後重新開啟程式即可。\n\n（也可以把「顯示方式」改成「視窗內」。）",
      )
      return
    if not browser and self.fig is None:
      messagebox.showinfo("提示", "需要先安裝 matplotlib：pip install matplotlib")
      return
    dates = [self.chart_dates[i] for i in self.chk_dates.checked()]
    items = [self.chart_items[i] for i in self.chk_items.checked()]
    if not dates or not items:
      messagebox.showwarning("提示", "請至少選一個日期和一個項目！")
      return
    source = SRC_KEYS[self.cmb_src.current()]
    param = self.var_param.get().strip()
    if source == "regex":
      try:
        re.compile(param)
      except re.error as e:
        messagebox.showerror("正則表達式錯誤", f"無法使用這個正則表達式：\n{e}")
        return
    if source == "cell" and coord_to_index(param) is None:
      messagebox.showerror("儲存格格式錯誤", "請輸入像 I3、J2 這樣的儲存格位置！")
      return
    entries = [(info, d) for d in dates for info in self.date_map[d]]
    points, warns, conflicts = build_points(
        entries, self.app.range_settings, self._name_func(), set(items), source, param
    )
    self.last_points = points
    layout = LAYOUT_KEYS[self.cmb_layout.current()]
    if browser:
      if self.plot_dir is None:
        self.plot_dir = tempfile.mkdtemp(prefix="excel_chart_")
      path = os.path.join(self.plot_dir, f"chart_{int(time.time() * 1000)}.html")
      try:
        st = build_plotly_html(
            points, layout, self.var_ctype.get(), self.var_connect.get(),
            Y_LABELS[source], self.chart_items, path,
        )
      except Exception as e:
        messagebox.showerror("畫圖失敗", f"產生互動圖時發生錯誤：\n{e}")
        return
      if st["points"] - st["missing"] <= 0:
        messagebox.showinfo("提示", "沒有可以畫的數據。")
      else:
        self.plot_files.append(path)
        for old in self.plot_files[:-5]:  # 只留最近 5 張，避免暫存檔越積越多
          try:
            os.remove(old)
          except OSError:
            pass
        self.plot_files = self.plot_files[-5:]
        webbrowser.open("file:///" + path.replace("\\", "/"))
      if self.fig is not None:
        self.fig.clear()
        self.fig.text(
            0.5, 0.5, "圖表已在瀏覽器開啟\n（要在這裡看，請把「顯示方式」改成「視窗內」）",
            ha="center", va="center", fontsize=13, color=PAL["muted"],
        )
        self.fig.set_facecolor(PAL["plot_bg"])
        self.canvas.draw_idle()
    else:
      st = draw_chart(
          self.fig, points, layout, self.var_ctype.get(),
          self.var_connect.get(), Y_LABELS[source], self.chart_items,
      )
      self.canvas.draw_idle()
    msg = f"共 {st['points']} 個點"
    if st["missing"]:
      msg += f"，其中 {st['missing']} 個沒有數值（圖上留空）"
    if st["skipped"]:
      msg += f"；{st['skipped']} 個項目名稱沒有數字（已略過）"
    if conflicts:
      msg += f"；{len(conflicts)} 個項目有重複數據（採用第一個）"
    self.lbl_chart.config(text=msg)
    if warns:
      uniq = list(dict.fromkeys(warns))
      more = f"\n...以及其他 {len(uniq) - 12} 筆" if len(uniq) > 12 else ""
      messagebox.showwarning(
          "取值範圍警告",
          "以下工作表的範圍有問題，數值已留空：\n\n" + "\n".join(uniq[:12]) + more,
      )

  def export_report(self):
    """用目前打勾的日期、項目與取值方式，輸出成 Excel：矩陣表＋長表。"""
    dates = [self.chart_dates[i] for i in self.chk_dates.checked()]
    items = [self.chart_items[i] for i in self.chk_items.checked()]
    if not dates or not items:
      messagebox.showwarning("提示", "請至少選一個日期和一個項目！")
      return
    source = SRC_KEYS[self.cmb_src.current()]
    param = self.var_param.get().strip()
    if source == "regex":
      try:
        re.compile(param)
      except re.error as e:
        messagebox.showerror("正則表達式錯誤", f"無法使用這個正則表達式：\n{e}")
        return
    if source == "cell" and coord_to_index(param) is None:
      messagebox.showerror("儲存格格式錯誤", "請輸入像 I3、J2 這樣的儲存格位置！")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx")],
        title="輸出整理報表",
    )
    if not path:
      return
    entries = [(info, d) for d in dates for info in self.date_map[d]]
    points, warns, conflicts = build_points(
        entries, self.app.range_settings, self._name_func(), set(items), source, param
    )
    order = {n: i for i, n in enumerate(self.chart_items)}
    points.sort(key=lambda p: (order.get(p["item"], 10**6), p["date"], p["file"]))
    label = Y_LABELS[source]
    long_df = pd.DataFrame([
        {"日期": p["date"], "檔案": p["file"], "項目": p["item"],
         "工作表": p["sheet"], label: p["value"]}
        for p in points
    ])
    matrix = {}
    for p in points:
      row = matrix.setdefault(p["item"], {})
      if p["value"] is not None and row.get(p["date"]) is None:
        row[p["date"]] = p["value"]
      else:
        row.setdefault(p["date"], None)
    wide_df = pd.DataFrame(
        [{"項目": it, **{d: matrix[it].get(d) for d in dates}}
         for it in items if it in matrix]
    )
    try:
      with pd.ExcelWriter(path, engine="openpyxl") as xw:
        wide_df.to_excel(xw, sheet_name="矩陣（項目×日期）", index=False)
        long_df.to_excel(xw, sheet_name="長表", index=False)
        for ws in xw.book.worksheets:
          ws.freeze_panes = "B2" if ws.title.startswith("矩陣") else "A2"
          for col in ws.columns:
            w = max((_disp_width(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 45)
      msg = f"已輸出：\n{path}\n\n共 {len(points)} 筆資料。"
      if conflicts:
        msg += f"\n有 {len(conflicts)} 個項目重複數據（矩陣採用第一個，長表全部保留）。"
      if warns:
        msg += f"\n有 {len(set(warns))} 個取值範圍警告，該數值已留空。"
      messagebox.showinfo("成功", msg)
    except Exception as e:
      messagebox.showerror("輸出失敗", str(e))

  def export_points(self):
    if not self.last_points:
      messagebox.showinfo("提示", "請先畫一張圖，再匯出圖上的數據。")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV 檔案", "*.csv"), ("Excel 檔案", "*.xlsx")],
        title="匯出圖上數據",
    )
    if not path:
      return
    df = pd.DataFrame([
        {"日期": p["date"], "檔案": p["file"], "項目": p["item"],
         "工作表": p["sheet"], "數值": p["value"]}
        for p in self.last_points
    ])
    try:
      if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False)
      else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
      messagebox.showinfo("成功", f"已匯出：\n{path}")
    except Exception as e:
      messagebox.showerror("匯出失敗", str(e))

  # ---------- 趨勢比較 ----------
  def _build_trend_tab(self):
    t = self.tab_trend
    bar = ttk.Frame(t)
    bar.pack(fill="x")
    ttk.Label(bar, text="項目").pack(side="left")
    self.cmb_tr_item = ttk.Combobox(bar, state="readonly", width=16)
    self.cmb_tr_item.pack(side="left", padx=(4, 14))
    ttk.Label(bar, text="基準日期").pack(side="left")
    self.cmb_tr_base = ttk.Combobox(bar, state="readonly", width=14)
    self.cmb_tr_base.pack(side="left", padx=(4, 14))
    ttk.Button(bar, text="更新", command=self._trend_update).pack(side="left")
    ttk.Button(bar, text="匯出這張表…", command=self._trend_export).pack(
        side="left", padx=8
    )
    self.cmb_tr_item.bind("<<ComboboxSelected>>", lambda e: self._trend_update())
    self.cmb_tr_base.bind("<<ComboboxSelected>>", lambda e: self._trend_update())
    ttk.Label(
        t,
        text="用目前的取值範圍算出各日期的 Rx，並和「基準日期」比較。變化% ＝（Rx − 基準 Rx）÷ 基準 Rx × 100。",
        style="Hint.TLabel",
    ).pack(anchor="w", pady=(6, 4))
    self.lbl_trend = ttk.Label(t, text="", style="Hint.TLabel")
    self.lbl_trend.pack(side="bottom", anchor="w")
    box = ttk.Frame(t)
    box.pack(fill="both", expand=True)
    self.tree_trend = self._make_tree(box, [
        ("item", "項目", 120), ("date", "日期", 100), ("file", "檔案", 230),
        ("rx", "Rx", 100), ("diff", "與基準差", 110), ("pct", "變化 %", 100),
    ])
    self.tree_trend.tag_configure("base", foreground=PAL["accent"])
    self.tree_trend.tag_configure("none", foreground=PAL["faint"])
    self.trend_rows = []

  def _refresh_trend_choices(self):
    items = ["全部項目"] + list(self.chart_items)
    self.cmb_tr_item.config(values=items)
    if self.cmb_tr_item.get() not in items:
      self.cmb_tr_item.set("全部項目")
    dates = list(self.chart_dates)
    self.cmb_tr_base.config(values=dates)
    if self.cmb_tr_base.get() not in dates:
      self.cmb_tr_base.set(dates[0] if dates else "")
    self._trend_update()

  def _trend_update(self):
    self.tree_trend.delete(*self.tree_trend.get_children())
    self.trend_rows = []
    base = self.cmb_tr_base.get()
    sel = self.cmb_tr_item.get()
    if not self.chart_dates or not base:
      self.lbl_trend.config(text="尚無資料。")
      return
    wanted = set(self.chart_items) if sel in ("", "全部項目") else {sel}
    entries = [(i, d) for d in self.chart_dates for i in self.date_map[d]]
    points, _w, _c = build_points(
        entries, self.app.range_settings, self._name_func(), wanted, "rx", ""
    )
    basev = {}
    for p in points:
      if p["date"] == base and p["value"] is not None:
        basev.setdefault(p["item"], p["value"])
    order = {n: i for i, n in enumerate(self.chart_items)}
    points.sort(key=lambda p: (order.get(p["item"], 10**6), p["date"], p["file"]))
    nobase = set()
    for p in points:
      v = p["value"]
      b = basev.get(p["item"])
      diff = pct = ""
      tags = ()
      if v is None:
        tags = ("none",)
      elif b is None:
        nobase.add(p["item"])
      else:
        diff = f"{v - b:+.3f}"
        pct = f"{(v - b) / b * 100:+.2f}" if b != 0 else ""
        if p["date"] == base:
          tags = ("base",)
      row = (
          p["item"], p["date"], p["file"], "" if v is None else f"{v:.3f}", diff, pct,
      )
      self.trend_rows.append(row)
      self.tree_trend.insert("", "end", values=row, tags=tags)
    msg = f"共 {len(points)} 筆。藍色＝基準日期；灰色＝沒有數值。"
    if nobase:
      msg += f"　{len(nobase)} 個項目在基準日期沒有數值，無法比較。"
    self.lbl_trend.config(text=msg)

  def _trend_export(self):
    if not self.trend_rows:
      messagebox.showinfo("提示", "目前沒有可匯出的資料。")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV 檔案", "*.csv"), ("Excel 檔案", "*.xlsx")],
        title="匯出趨勢比較表",
    )
    if not path:
      return
    df = pd.DataFrame(
        self.trend_rows, columns=["項目", "日期", "檔案", "Rx", "與基準差", "變化%"]
    )
    for c in ("Rx", "與基準差", "變化%"):
      df[c] = pd.to_numeric(df[c], errors="coerce")
    try:
      if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False)
      else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
      messagebox.showinfo("成功", f"已匯出：\n{path}")
    except Exception as e:
      messagebox.showerror("匯出失敗", str(e))

  # ---------- 概況 ----------
  def _make_tree(self, parent, cols):
    tree = ttk.Treeview(parent, columns=[c[0] for c in cols], show="headings")
    for cid, text, width in cols:
      tree.heading(cid, text=text)
      tree.column(cid, width=self.app.px(width), anchor="w")
    tree.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    sb.pack(side="right", fill="y")
    tree.configure(yscrollcommand=sb.set)
    return tree

  def _build_overview_tab(self):
    t = self.tab_over
    self.lbl_over = ttk.Label(t, text="", justify="left", wraplength=self.app.px(1000))
    self.lbl_over.pack(anchor="w")
    pw = ttk.PanedWindow(t, orient="vertical")
    pw.pack(fill="both", expand=True, pady=6)
    f1 = ttk.LabelFrame(pw, text="每個 Excel", padding=4)
    f2 = ttk.LabelFrame(pw, text="每個項目的 Rx 統計（依目前取值範圍設定）", padding=4)
    pw.add(f1, weight=1)
    pw.add(f2, weight=1)
    self.tree_files = self._make_tree(f1, [
        ("date", "日期", 90), ("file", "檔案", 190), ("n", "數據項目", 70),
        ("ok", "有數值", 70), ("bad", "沒有數值的工作表", 260),
        ("t", "溫度中位數(°C)", 110), ("h", "濕度中位數(%)", 110),
    ])
    self.tree_items = self._make_tree(f2, [
        ("item", "項目", 150), ("cnt", "出現次數", 80), ("mean", "平均 Rx", 100),
        ("std", "標準差", 100), ("min", "最小", 90), ("max", "最大", 90),
        ("miss", "沒有數值次數", 110), ("odd", "疑似異常（偏離平均 >3 倍標準差）", 300),
    ])
    self.tree_items.tag_configure("odd", foreground=PAL["warn"])

  def _refresh_overview(self):
    self.tree_files.delete(*self.tree_files.get_children())
    self.tree_items.delete(*self.tree_items.get_children())
    n_files = len(self.infos)
    if not n_files:
      self.lbl_over.config(text="尚無資料。請先選擇根資料夾或額外加入 Excel。")
      return
    settings = self.app.range_settings
    n_sheets = 0
    for d in sorted(self.date_map):
      for info in self.date_map[d]:
        data_sheets = [s for s in info["sheets"] if s["kind"] == "data"]
        n_sheets += len(data_sheets)
        bad = []
        for s in data_sheets:
          v = get_values(s, settings.get(range_key(info["path"], s["name"])))
          if v["rx"] is None and v["a0"] is None:
            bad.append(s["name"] + ("(範圍有問題)" if v["warn"] else ""))
        t_txt, h_txt = median_reading(info)
        self.tree_files.insert("", "end", values=(
            d, info["filename"], len(data_sheets), len(data_sheets) - len(bad),
            "、".join(bad), t_txt, h_txt,
        ))
    entries = [(i, d) for d in sorted(self.date_map) for i in self.date_map[d]]
    name_of = self._name_func()
    names = set()
    for info, _ in entries:
      for s in info["sheets"]:
        if s["kind"] == "data":
          names.add(name_of(s["name"]))
    points, _w, _c = build_points(entries, settings, name_of, names, "rx", "")
    per = {}
    per_pts = {}
    for p in points:
      per.setdefault(p["item"], []).append(p["value"])
      per_pts.setdefault(p["item"], []).append((p["date"], p["value"]))
    for item in sorted(per, key=natural_key):
      vals = [v for v in per[item] if v is not None]
      odd = ""
      if vals:
        mean = sum(vals) / len(vals)
        std = statistics.stdev(vals) if len(vals) > 1 else 0.0
        if len(vals) >= 4 and std > 0:
          bad = [
              f"{d[5:]}（{v:.3f}）" for d, v in per_pts[item]
              if v is not None and abs(v - mean) > 3 * std
          ]
          odd = f"{len(bad)} 筆：" + "、".join(bad) if bad else ""
        elif len(vals) < 4:
          odd = "數據少於 4 筆，不判斷"
        row = (item, len(per[item]), f"{mean:.3f}", f"{std:.3f}",
               f"{min(vals):.3f}", f"{max(vals):.3f}", len(per[item]) - len(vals), odd)
      else:
        row = (item, len(per[item]), "", "", "", "", len(per[item]), "")
      tags = ("odd",) if odd and not odd.startswith("數據少於") else ()
      self.tree_items.insert("", "end", values=row, tags=tags)
    lines = [f"共 {n_files} 個 Excel、{n_sheets} 個數據工作表、{len(per)} 個項目。"]
    if self.date_map:
      days = sorted(self.date_map)
      d0 = datetime.strptime(days[0], "%Y-%m-%d")
      d1 = datetime.strptime(days[-1], "%Y-%m-%d")
      total_days = (d1 - d0).days + 1
      lines.append(
          f"日期範圍：{days[0]} ～ {days[-1]}，其中 {len(days)} 天有資料、"
          f"{total_days - len(days)} 天沒有。"
      )
    if self.undated:
      lines.append(f"另有 {len(self.undated)} 個檔案沒有日期或讀取失敗（見日曆下方清單）。")
    self.lbl_over.config(text="\n".join(lines))


class ViewMode(Mode):
  """檢視：日曆、圖表、概況、趨勢比較、AI 分析（畫面在 ViewPage 類別裡）。"""

  title = "📈 檢視"
  order = 30

  def __init__(self, app):
    super().__init__(app)
    self.view_page = None

  def start(self):
    """主頁按下按鈕時會呼叫這裡。"""
    self.show_view_mode()

  def show_view_mode(self):
    self.app.clear_container()
    w = min(max(self.app.winfo_width(), self.app.px(1150)), self.app.winfo_screenwidth() - 60)
    h = min(max(self.app.winfo_height(), self.app.px(820)), self.app.winfo_screenheight() - 90)
    self.app.geometry(f"{w}x{h}")
    self.view_page = ViewPage(self.app)
