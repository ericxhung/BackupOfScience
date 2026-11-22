"""模式：數據列印。"""

import os
import re
import tkinter as tk
import openpyxl
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from openpyxl.styles import Border, Font, Side

from core import TH_MODE_FIRST, TH_MODE_MEDIAN, first_reading, get_values, median_reading, range_key, sheet_has_values
from dialogs import get_values, range_key
from modes.base import Mode


class PrintMode(Mode):
  """數據列印：選一個 Excel，確認日期與溫濕度，輸出成可列印的 Excel。"""

  title = "🖨️ 數據列印"
  order = 10

  def __init__(self, app):
    super().__init__(app)
    self.print_info = None

  def start(self):
    """主頁按下按鈕時會呼叫這裡。"""
    self.show_print_mode_select()

  def show_print_mode_select(self):
    file_path = filedialog.askopenfilename(
        title="選取數據列印用的 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xls *.xlsm"), ("所有檔案", "*.*")],
    )
    if not file_path:
      return

    filename = os.path.basename(file_path)
    folder_name = os.path.basename(os.path.dirname(file_path))

    # 解析日期
    date_str = self.app.parse_date_from_string(filename)
    if not date_str:
      date_str = self.app.parse_date_from_string(folder_name)
    if not date_str:
      user_choice = self.app.request_manual_date(filename, folder_name)
      if user_choice == "SKIP" or user_choice is None:
        return
      date_str = user_choice

    self.show_print_layout_page(file_path, date_str)

  def show_print_layout_page(self, file_path, date_str):
    self.app.clear_container()

    lbl_title = ttk.Label(
        self.app.container, text="數據列印設定", style="H2.TLabel"
    )
    lbl_title.pack(pady=10)

    # 1. 上方設定區域 (日期與氣象參數)
    top_frame = ttk.LabelFrame(
        self.app.container, text="基本資訊設定", padding=12
    )
    top_frame.pack(fill="x", padx=20, pady=5)

    # 上方左邊：日期
    ttk.Label(top_frame, text="日期 (YYYY MM DD):").grid(
        row=0, column=0, sticky="w", padx=5, pady=5
    )
    self.entry_print_date = ttk.Entry(top_frame, width=15)
    self.entry_print_date.grid(row=0, column=1, sticky="w", padx=5, pady=5)
    formatted_date_input = date_str.replace("-", " ")
    self.entry_print_date.insert(0, formatted_date_input)

    # 上方右邊：溫度、濕度、天氣狀態
    ttk.Label(top_frame, text="溫度 (°C):").grid(
        row=0, column=2, sticky="w", padx=(20, 5), pady=5
    )
    self.entry_temp = ttk.Entry(top_frame, width=10)
    self.entry_temp.grid(row=0, column=3, sticky="w", padx=5, pady=5)

    ttk.Label(top_frame, text="濕度 (%):").grid(
        row=0, column=4, sticky="w", padx=(10, 5), pady=5
    )
    self.entry_humidity = ttk.Entry(top_frame, width=10)
    self.entry_humidity.grid(row=0, column=5, sticky="w", padx=5, pady=5)

    ttk.Label(top_frame, text="天氣狀態:").grid(
        row=0, column=6, sticky="w", padx=(10, 5), pady=5
    )
    self.entry_weather = ttk.Entry(top_frame, width=12)
    self.entry_weather.grid(row=0, column=7, sticky="w", padx=5, pady=5)
    self.entry_weather.insert(0, "晴")

    # 讀取 Excel（共用取值層）：工作表分類、溫濕度
    self.print_file_path = file_path
    sheet_names = []
    self.print_info = None
    infos = self.app.load_infos([file_path], title="讀取 Excel 檔案中…")
    if infos is not None:
      info = infos[file_path]
      if info.get("error"):
        messagebox.showerror("錯誤", f"讀取 Excel 檔案失敗：\n{info['error']}")
      else:
        self.print_info = info
        sheet_names = [s["name"] for s in info["sheets"]]

    # 溫濕度取法（預設維持原本做法：開頭第一筆）
    ttk.Label(top_frame, text="溫濕度取法:").grid(
        row=1, column=0, sticky="w", padx=5, pady=5
    )
    self.th_mode_var = tk.StringVar(value=TH_MODE_FIRST)
    cmb_th = ttk.Combobox(
        top_frame,
        textvariable=self.th_mode_var,
        values=[TH_MODE_FIRST, TH_MODE_MEDIAN],
        state="readonly",
        width=30,
    )
    cmb_th.grid(row=1, column=1, columnspan=5, sticky="w", padx=5, pady=5)
    cmb_th.bind("<<ComboboxSelected>>", lambda e: self.refresh_print_th())
    self.refresh_print_th()

    # 2. 中央列表：項目勾選清單
    list_frame = ttk.LabelFrame(
        self.app.container, text="選取列印項目 (工作表)", padding=10
    )
    list_frame.pack(fill="both", expand=True, padx=20, pady=10)

    canvas = tk.Canvas(
        list_frame, bg=self.app.theme_bg, highlightthickness=0, borderwidth=0
    )
    scrollbar = ttk.Scrollbar(
        list_frame, orient="vertical", command=canvas.yview
    )
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    self.print_sheet_vars = {}
    sheet_by_name = (
        {s["name"]: s for s in self.print_info["sheets"]}
        if self.print_info
        else {}
    )
    for sname in sheet_names:
      sh = sheet_by_name.get(sname)
      is_check = (
          sname not in ["氣溫", "溫度"]
          and "plxdaq" not in sname.lower()
          and sh is not None
          and sh["kind"] not in ("temp", "plxdaq")
          and not sh["hidden"]
          and not (sh["kind"] == "data" and not sheet_has_values(sh))
      )
      label = sname
      if sh is not None and sh["kind"] == "temp":
        label += "（溫濕度紀錄）"
      elif sh is not None and sh["kind"] == "data" and not sheet_has_values(sh):
        label += "（沒有數值）"
      var = tk.BooleanVar(value=is_check)
      chk = ttk.Checkbutton(scroll_frame, text=label, variable=var)
      chk.pack(anchor="w", pady=2)
      self.print_sheet_vars[sname] = var

    # 3. 下方控制按鈕
    bottom_frame = ttk.Frame(self.app.container)
    bottom_frame.pack(fill="x", padx=20, pady=10)

    btn_back = ttk.Button(
        bottom_frame, text="返回主選單", command=self.app.show_main_home
    )
    btn_back.pack(side="left")

    btn_range = ttk.Button(
        bottom_frame,
        text="⚙ 取值範圍",
        command=lambda: self.app.open_range_dialog([self.print_file_path]),
    )
    btn_range.pack(side="left", padx=8)

    btn_confirm = ttk.Button(
        bottom_frame,
        text="確定 (生成列印用 Excel)",
        style="Accent.TButton",
        command=self.generate_print_excel,
    )
    btn_confirm.pack(side="right")

  def generate_print_excel(self):
    date_val = self.entry_print_date.get().strip()
    match = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", date_val)
    if not match:
      messagebox.showerror(
          "格式錯誤", "日期格式不正確，請輸入 YYYY MM DD（例如：2026 10 15）"
      )
      return

    year, month, day = match.groups()
    temp_val = self.entry_temp.get().strip()
    try:
      temp_val = str(int(float(temp_val)))
    except ValueError:
      pass

    humidity_val = self.entry_humidity.get().strip()
    weather_val = self.entry_weather.get().strip()

    selected_sheets = [
        s for s, var in self.print_sheet_vars.items() if var.get()
    ]
    if not selected_sheets:
      messagebox.showerror("錯誤", "請至少勾選一個項目！")
      return

    try:
      info = self.app.info_cache.get(self.print_file_path)
      if info.get("error"):
        raise RuntimeError(info["error"])
      by_name = {s["name"]: s for s in info["sheets"]}
      export_rows = []
      range_warns = []

      for sname in selected_sheets:
        rx_val = ""
        a0_val = ""
        sh = by_name.get(sname)
        if sh is not None:
          v = get_values(
              sh, self.app.range_settings.get(range_key(self.print_file_path, sname))
          )
          if v["rx"] is not None:
            rx_val = v["rx"]
          if v["a0"] is not None:
            a0_val = v["a0"]
          for w in v["warn"]:
            range_warns.append(f"{sname}：{w}")
        export_rows.append((sname, rx_val, a0_val))

      out_wb = openpyxl.Workbook()
      out_sheet = out_wb.active
      out_sheet.title = "數據列印"

      # 自動計算星期幾（只顯示數字，例如：一）
      dt_obj = datetime(int(year), int(month), int(day))
      weekdays_zh = ["一", "二", "三", "四", "五", "六", "日"]
      weekday_str = weekdays_zh[dt_obj.weekday()]

      # 填入 A1 整合資訊
      out_sheet["A1"] = (
          f"{year}年 {int(month)}月 {int(day)}日 ({weekday_str})"
          f" {temp_val}℃ {humidity_val}% {weather_val}"
      )

      # 填入表格標頭 (從 A3 開始)
      out_sheet["A3"] = "項目"
      out_sheet["B3"] = "Rx"
      out_sheet["C3"] = "A0"

      current_row = 4
      for name, rx, a0 in export_rows:
        out_sheet[f"A{current_row}"] = name
        out_sheet[f"B{current_row}"] = rx
        out_sheet[f"C{current_row}"] = a0
        current_row += 1

      # 字體大小改為 10
      for row in out_sheet.iter_rows():
        for cell in row:
          current_font = cell.font
          bold = current_font.bold if current_font else False
          italic = current_font.italic if current_font else False
          cell.font = Font(name="Arial", size=10, bold=bold, italic=italic)

      # 替項目數據表格加入所有框線 (從第 3 行到最後一列，A 欄到 C 欄)
      thin_side = Side(style="thin", color="000000")
      table_border = Border(
          left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
      )
      for r in range(3, current_row):
        for c in range(1, 4):
          out_sheet.cell(row=r, column=c).border = table_border

      # 將所有縱列寬度一律設定為 19.38
      for col in out_sheet.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        out_sheet.column_dimensions[col_letter].width = 19.38

      save_path = filedialog.asksaveasfilename(
          defaultextension=".xlsx",
          filetypes=[("Excel 檔案", "*.xlsx")],
          title="選擇儲存結果位置與檔名",
      )

      if not save_path:
        return

      out_wb.save(save_path)
      messagebox.showinfo("成功", f"列印用 Excel 檔案已成功生成至：\n{save_path}")
      if range_warns:
        more = (
            f"\n...以及其他 {len(range_warns) - 12} 筆" if len(range_warns) > 12 else ""
        )
        messagebox.showwarning(
            "取值範圍警告",
            "以下項目的範圍有問題，數值已留空：\n\n"
            + "\n".join(range_warns[:12])
            + more,
        )

      # 自動用 Excel 開啟剛剛儲存的檔案 (Windows 專用，視窗保持開啟)
      try:
        os.startfile(save_path)
      except Exception as e:
        print(f"自動開啟檔案失敗: {e}")

    except Exception as e:
      messagebox.showerror("重大錯誤", f"生成列印檔案時發生錯誤：\n{e}")

  def refresh_print_th(self):
    """依「溫濕度取法」重新填入列印頁的溫度與濕度。"""
    info = self.print_info
    temp, hum = "", ""
    if info is not None:
      if self.th_mode_var.get() == TH_MODE_MEDIAN:
        temp, hum = median_reading(info)
      else:
        temp, hum = first_reading(info)
    self.entry_temp.delete(0, tk.END)
    self.entry_humidity.delete(0, tk.END)
    if temp:
      self.entry_temp.insert(0, temp)
    if hum:
      self.entry_humidity.insert(0, hum)
