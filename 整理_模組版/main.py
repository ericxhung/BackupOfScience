"""主程式：視窗外殼、主題、介面大小、主頁。各功能在 modes 資料夾，執行這個檔案即可啟動。"""

import os
import sys
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk

from core import app_dir, InfoCache, PAL, UI_SCALE_CHOICES, enable_high_dpi, load_range_settings, load_theme_setting, load_ui_scale, save_theme_setting, save_ui_scale, set_palette
from services import SharedServices
from modes import LOAD_ERRORS, MODE_CLASSES


# 【UI 美化】嘗試載入 sv-ttk（Windows 11 風格主題）。
# 沒安裝也能執行，只是會退回原本的 tkinter 外觀。
try:
  import sv_ttk

  HAS_SV_TTK = True
except ImportError:
  HAS_SV_TTK = False


class ExcelProcessorApp(SharedServices, tk.Tk):
  """主視窗外殼：主題、字體、介面大小、主頁。各功能在 modes 資料夾。"""

  def __init__(self):
    super().__init__()
    self.title("科展資料處理")

    # 【UI 美化】依螢幕縮放比例調整視窗大小
    # base_tk_scaling＝Windows 顯示縮放換算出來的原始值；ui_scale＝程式內「介面大小」倍率
    self.base_tk_scaling = float(self.tk.call("tk", "scaling"))
    self.ui_scale = load_ui_scale()
    self.tk.call("tk", "scaling", self.base_tk_scaling * self.ui_scale)
    self.scale = max(0.6, self.winfo_fpixels("1i") / 96.0)
    self.geometry(f"{self.px(900)}x{self.px(750)}")
    self.minsize(self.px(820), self.px(650))

    # 【UI 美化】套用主題、統一字型、設定樣式
    self.theme_name = load_theme_setting()
    if self.theme_name == "tech" and not HAS_SV_TTK:
      self.theme_name = "light"  # 科技深色需要 sv-ttk
    set_palette(self.theme_name)
    self.setup_theme()
    self.setup_fonts()
    self.setup_styles()

    # 共用取值：檔案資訊快取、各工作表的取值範圍設定
    self.info_cache = InfoCache()
    self.range_settings = load_range_settings()

    self.container = ttk.Frame(self, padding=10)
    self.container.pack(fill="both", expand=True)

    # 各模式：由 modes 資料夾自動載入（新增模式不用改這裡）
    self.mode_objs = []
    for mode_cls in MODE_CLASSES:
      try:
        self.mode_objs.append(mode_cls(self))
      except Exception as e:
        LOAD_ERRORS.append((mode_cls.__name__, f"{type(e).__name__}: {e}"))

    # 進入點：顯示系統主頁
    self.show_main_home()

  def report_callback_exception(self, exc, val, tb):
    """視窗事件裡發生未處理的錯誤時：寫進 error_log.txt。
    打包成 exe 後沒有黑色命令列視窗，錯誤才不會「靜靜消失」。"""
    import time
    import traceback

    text = "".join(traceback.format_exception(exc, val, tb))
    try:
      with open(os.path.join(app_dir(), "error_log.txt"), "a", encoding="utf-8") as f:
        f.write(time.strftime("[%Y-%m-%d %H:%M:%S]\n") + text + "\n")
    except OSError:
      pass
    traceback.print_exception(exc, val, tb)  # 一般執行時，命令列照常顯示
    # 只有打包版才跳視窗；而且 5 秒內最多跳一次，避免重複錯誤洗版
    if getattr(sys, "frozen", False):
      now = time.time()
      if now - getattr(self, "_last_err_popup", 0) > 5:
        self._last_err_popup = now
        try:
          messagebox.showerror(
              "發生錯誤",
              f"{exc.__name__}: {val}\n\n詳細內容已記錄在 error_log.txt",
              parent=self,
          )
        except Exception:
          pass

  def setup_theme(self):
    if HAS_SV_TTK:
      sv_ttk.set_theme("dark" if self.theme_name == "tech" else "light")
    else:
      try:
        ttk.Style().theme_use("vista")
      except tk.TclError:
        pass

  def setup_fonts(self):
    families = set(tkfont.families())
    self.base_family = None
    for cand in (
        "Microsoft JhengHei UI",
        "Microsoft JhengHei",
        "PingFang TC",
        "Noto Sans CJK TC",
        "Arial",
    ):
      if cand in families:
        self.base_family = cand
        break
    if self.base_family is None:
      self.base_family = tkfont.nametofont("TkDefaultFont").actual("family")

    for name in (
        "TkDefaultFont",
        "TkTextFont",
        "TkMenuFont",
        "TkHeadingFont",
        "TkCaptionFont",
        "SunValleyCaptionFont",
        "SunValleyBodyFont",
        "SunValleyBodyStrongFont",
        "SunValleyBodyLargeFont",
        "SunValleySubtitleFont",
        "SunValleyTitleFont",
        "SunValleyTitleLargeFont",
        "SunValleyDisplayFont",
    ):
      try:
        tkfont.nametofont(name).configure(family=self.base_family)
      except tk.TclError:
        pass

    self.font_title = (self.base_family, 20, "bold")
    self.font_h2 = (self.base_family, 14, "bold")
    self.font_list = (self.base_family, 11)

  FONT_NAMES = (
      "TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
      "TkCaptionFont", "TkSmallCaptionFont", "TkFixedFont", "TkTooltipFont",
      "SunValleyCaptionFont", "SunValleyBodyFont", "SunValleyBodyStrongFont",
      "SunValleyBodyLargeFont", "SunValleySubtitleFont", "SunValleyTitleFont",
      "SunValleyTitleLargeFont", "SunValleyDisplayFont",
  )

  def scale_pixel_fonts(self):
    """sv-ttk 的字體大小是「像素」（負數），不會跟著 Windows 縮放，所以這裡手動乘上縮放倍率。"""
    if not hasattr(self, "_font_base"):
      self._font_base = {}
    for name in self.FONT_NAMES:
      try:
        f = tkfont.nametofont(name)
        if name not in self._font_base:
          size = int(f.cget("size"))
          self._font_base[name] = size if size < 0 else None
        base = self._font_base[name]
        if base is not None:
          f.configure(size=-max(8, int(round(abs(base) * self.scale))))
      except (tk.TclError, ValueError):
        pass

  def setup_styles(self):
    self.scale_pixel_fonts()
    style = ttk.Style()
    style.configure("Treeview", rowheight=self.px(28))
    style.configure("Title.TLabel", font=self.font_title)
    style.configure("H2.TLabel", font=self.font_h2)
    style.configure("Hint.TLabel", foreground=PAL["muted"])
    style.configure("Big.TButton", font=(self.base_family, 12))

    # 【首頁放大】只給主選單頁使用，其他頁面不受影響
    style.configure(
        "HomeTitle.TLabel", font=(self.base_family, 34, "bold")
    )
    style.configure(
        "HomeSub.TLabel", font=(self.base_family, 16), foreground=PAL["muted"]
    )
    style.configure(
        "Home.Accent.TButton", font=(self.base_family, 18, "bold")
    )

    # 給舊式 tk 元件 (Listbox / Canvas / Toplevel) 用的顏色，讓它們跟主題一致
    try:
      bg = style.lookup(".", "background")
      self.theme_bg = bg if bg else ("#1c1c1c" if PAL["dark"] else "#fafafa")
    except tk.TclError:
      self.theme_bg = "#1c1c1c" if PAL["dark"] else "#fafafa"
    self.theme_fg = PAL["fg"]
    self.theme_select_bg = PAL["sel_bg"]
    try:
      self.configure(bg=self.theme_bg)
    except tk.TclError:
      pass

  def apply_ui_scale(self, value):
    """改變介面大小倍率：重設 tk 縮放（字體與元件一起變），並回到主頁重建畫面。"""
    old = self.scale
    self.ui_scale = value
    self.tk.call("tk", "scaling", self.base_tk_scaling * value)
    self.scale = max(0.6, self.winfo_fpixels("1i") / 96.0)
    r = self.scale / old if old else 1.0
    try:
      w = max(self.px(820), int(self.winfo_width() * r))
      h = max(self.px(650), int(self.winfo_height() * r))
      self.minsize(self.px(820), self.px(650))
      self.geometry(f"{w}x{h}")
    except tk.TclError:
      pass
    self.setup_styles()
    save_ui_scale(value)
    self.show_main_home()

  def toggle_theme(self):
    """切換 淺色／科技深色，並重新顯示主頁（其他頁面進入時會用新主題重建）。"""
    new = "light" if self.theme_name == "tech" else "tech"
    if new == "tech" and not HAS_SV_TTK:
      messagebox.showinfo(
          "提示",
          "科技深色主題需要 sv-ttk。\n請先在命令列執行：pip install sv-ttk\n安裝後重新開啟程式即可。",
      )
      return
    self.theme_name = new
    set_palette(new)
    if HAS_SV_TTK:
      sv_ttk.set_theme("dark" if new == "tech" else "light")
    self.setup_styles()
    save_theme_setting(new)
    self.show_main_home()

  def show_main_home(self):
    self.clear_container()

    center = ttk.Frame(self.container)
    center.place(relx=0.5, rely=0.42, anchor="center")

    lbl_title = ttk.Label(
        center, text="科展資料處理", style="HomeTitle.TLabel"
    )
    lbl_title.pack(pady=(0, 10))

    lbl_sub = ttk.Label(
        center, text="請選擇要使用的功能", style="HomeSub.TLabel"
    )
    lbl_sub.pack(pady=(0, 40))

    btn_frame = ttk.Frame(center)
    btn_frame.pack()

    # 主頁按鈕由「已載入的模式」自動產生（依各模式的 order 排序）
    for mode in self.mode_objs:
      ttk.Button(
          btn_frame,
          text=mode.title,
          width=22,
          style="Home.Accent.TButton",
          command=mode.start,
      ).pack(pady=14, ipady=self.px(18))
    for name, err in LOAD_ERRORS:
      ttk.Label(
          btn_frame,
          text=f"⚠ 模式「{name}」載入失敗：{err}",
          foreground=PAL["warn"],
          wraplength=self.px(420),
      ).pack(pady=(6, 0))

    opt = ttk.Frame(center)
    opt.pack(pady=(20, 0))
    ttk.Button(
        opt,
        text="🌓 切換主題（目前：" + ("科技深色" if self.theme_name == "tech" else "淺色") + "）",
        command=self.toggle_theme,
    ).pack(side="left", padx=(0, 16))
    ttk.Label(opt, text="介面大小").pack(side="left", padx=(0, 6))
    cmb_scale = ttk.Combobox(
        opt, state="readonly", width=6,
        values=[f"{int(round(c * 100))}%" for c in UI_SCALE_CHOICES],
    )
    cur = min(UI_SCALE_CHOICES, key=lambda c: abs(c - self.ui_scale))
    cmb_scale.set(f"{int(round(cur * 100))}%")
    cmb_scale.pack(side="left")
    cmb_scale.bind(
        "<<ComboboxSelected>>",
        lambda e: self.apply_ui_scale(UI_SCALE_CHOICES[cmb_scale.current()]),
    )
    ttk.Label(
        center,
        text="介面大小是在 Windows 顯示縮放之外再加的倍率，會記住設定。",
        style="Hint.TLabel",
    ).pack(pady=(6, 0))

    # 右下角的淡淡署名
    ttk.Label(
        self.container, text="By 洪旭辰", foreground=PAL["faint"],
        font=(self.base_family, 11),
    ).place(relx=1.0, rely=1.0, x=-6, y=-4, anchor="se")


if __name__ == "__main__":
  # 【UI 美化】必須在建立視窗之前呼叫，高 DPI 設定才會生效
  enable_high_dpi()
  app = ExcelProcessorApp()
  app.mainloop()
