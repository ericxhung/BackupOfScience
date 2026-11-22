"""所有模式共用的服務：開檔進度、日期解析、手動日期視窗、取值範圍視窗、對話框置中等。"""

import os
import re
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

from dialogs import ProgressWindow, RangeDialog
from ai_chat import ProgressWindow


class SharedServices:
  """所有模式共用的服務（主視窗會繼承它，所以模式可以用 self.app.xxx 呼叫）。"""

  def px(self, n):
    return int(n * self.scale)

  def center_dialog(self, top, w, h):
    w, h = self.px(w), self.px(h)
    self.update_idletasks()
    x = self.winfo_rootx() + (self.winfo_width() - w) // 2
    y = self.winfo_rooty() + (self.winfo_height() - h) // 2
    top.geometry(f"{w}x{h}+{max(x, 0)}+{max(y, 0)}")
    top.configure(bg=self.theme_bg)
    top.transient(self)

  def clear_container(self):
    for widget in self.container.winfo_children():
      widget.destroy()

  def load_infos(self, paths, parent=None, title="讀取 Excel 資料中…"):
    """讀取（或從快取取得）多個檔案的資訊；會顯示可取消的進度視窗。取消時回傳 None。"""
    parent = parent or self
    paths = list(paths)
    infos = {}
    pending = [p for p in paths if self.info_cache.peek(p) is None]
    prog = ProgressWindow(self, parent, title, len(pending)) if pending else None
    try:
      for p in paths:
        if prog is not None and self.info_cache.peek(p) is None:
          if prog.cancelled:
            return None
          prog.set_text(os.path.basename(p))
          infos[p] = self.info_cache.get(p)
          prog.step()
        else:
          infos[p] = self.info_cache.get(p)
    finally:
      if prog is not None:
        prog.close()
      self.info_cache.save()
    return infos

  def date_text_of(self, info):
    """從檔名或資料夾名解析日期（YYYY-MM-DD），解析不到回傳空字串。"""
    d = self.parse_date_from_string(info["filename"])
    if not d:
      d = self.parse_date_from_string(
          os.path.basename(os.path.dirname(info["path"]))
      )
    return d or ""

  def open_range_dialog(self, paths):
    paths = list(paths)
    if not paths:
      messagebox.showinfo("提示", "目前沒有可以設定的檔案，請先加入 Excel 檔案！")
      return False
    infos = self.load_infos(paths, title="讀取工作表資訊中…")
    if infos is None:
      return False
    dlg = RangeDialog(self, [infos[p] for p in paths])
    self.wait_window(dlg)
    return dlg.saved

  def parse_date_from_string(self, text):
    match = re.search(r"(\d{4})_(\d{1,2})_(\d{1,2})", text)
    if match:
      y, m, d = match.groups()
      try:
        dt = datetime(int(y), int(m), int(d))
        return dt.strftime("%Y-%m-%d")
      except ValueError:
        pass
    return None

  def request_manual_date(self, filename, folder_name):
    top = tk.Toplevel(self)
    top.title(f"無法辨識日期: {filename}")
    self.center_dialog(top, 460, 220)
    top.grab_set()

    body = ttk.Frame(top, padding=15)
    body.pack(fill="both", expand=True)

    ttk.Label(
        body,
        text=(
            f"無法從檔名 '{filename}' 或資料夾 '{folder_name}' 辨識日期。\n"
            "請手動輸入正確日期（格式：YYYY MM DD），或點擊「不加入」略過："
        ),
        wraplength=self.px(420),
    ).pack(anchor="w")

    entry_date = ttk.Entry(body, width=25)
    entry_date.pack(pady=10)
    entry_date.insert(0, "2026 01 01")

    result_date = [None]

    def on_submit():
      val = entry_date.get().strip()
      match = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", val)
      if match:
        y, m, d = match.groups()
        try:
          dt = datetime(int(y), int(m), int(d))
          result_date[0] = dt.strftime("%Y-%m-%d")
          top.destroy()
          return
        except ValueError:
          pass
      messagebox.showerror(
          "格式錯誤",
          "日期格式不正確，請重新輸入（例如：2026 10 15）",
          parent=top,
      )

    def on_cancel():
      result_date[0] = "SKIP"
      top.destroy()

    btn_frame = ttk.Frame(body)
    btn_frame.pack(pady=10)

    ttk.Button(
        btn_frame, text="確定加入", style="Accent.TButton", command=on_submit
    ).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="不加入", command=on_cancel).pack(
        side="left", padx=5
    )

    self.wait_window(top)
    return result_date[0]
