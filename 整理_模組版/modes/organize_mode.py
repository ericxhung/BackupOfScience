"""模式：數據整理（步驟一～三）。"""

import json
import os
import tkinter as tk
import openpyxl
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from core import app_dir, pd, PAL, get_values, range_key
from dialogs import PAL, get_values, range_key
from modes.base import Mode


class OrganizeMode(Mode):
  """數據整理：多個 Excel 依日期整理（步驟一加入檔案、步驟二管理項目、步驟三輸出）。"""

  title = "📊 數據整理"
  order = 20

  def __init__(self, app):
    super().__init__(app)
    self.file_list = []
    self.item_mapping = {}
    self.mode_var = tk.StringVar(master=app, value="RESISTOR")

  def start(self):
    """主頁按下按鈕時會呼叫這裡。"""
    self.show_step1_home()

  def show_step1_home(self):
    self.app.clear_container()

    lbl_title = ttk.Label(
        self.app.container,
        text="【數據整理】步驟一：選擇模式與管理 Excel 檔案",
        style="H2.TLabel",
    )
    lbl_title.pack(pady=10)

    mode_frame = ttk.LabelFrame(
        self.app.container, text="搜尋模式設定", padding=12
    )
    mode_frame.pack(fill="x", padx=20, pady=5)

    radio_resistor = ttk.Radiobutton(
        mode_frame,
        text="1. 電阻模式 (對應 '=[數字]*[其他]' 格式)",
        value="RESISTOR",
        variable=self.mode_var,
    )
    radio_resistor.pack(anchor="w", pady=3)

    radio_average = ttk.Radiobutton(
        mode_frame,
        text="2. 平均值模式 (對應 '=AVERAGE([其他]' 格式)",
        value="AVERAGE",
        variable=self.mode_var,
    )
    radio_average.pack(anchor="w", pady=3)

    btn_frame = ttk.Frame(self.app.container)
    btn_frame.pack(fill="x", padx=20, pady=8)

    btn_add_files = ttk.Button(
        btn_frame, text="加入 Excel 檔案", command=self.add_excel_files
    )
    btn_add_files.pack(side="left", padx=5)

    btn_add_folder = ttk.Button(
        btn_frame, text="從資料夾遞迴加入", command=self.add_excel_folder
    )
    btn_add_folder.pack(side="left", padx=5)

    list_frame = ttk.Frame(self.app.container)
    list_frame.pack(fill="both", expand=True, padx=20, pady=10)

    self.tree_step1 = ttk.Treeview(
        list_frame, columns=("Filename", "Date", "Path"), show="headings"
    )
    self.tree_step1.heading("Filename", text="檔案名稱")
    self.tree_step1.heading("Date", text="解析日期 (YYYY-MM-DD)")
    self.tree_step1.heading("Path", text="完整路徑")
    self.tree_step1.column("Filename", width=self.app.px(220))
    self.tree_step1.column("Date", width=self.app.px(150))
    self.tree_step1.column("Path", width=self.app.px(380))
    self.tree_step1.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        list_frame, orient="vertical", command=self.tree_step1.yview
    )
    scrollbar.pack(side="right", fill="y")
    self.tree_step1.configure(yscrollcommand=scrollbar.set)

    bottom_frame = ttk.Frame(self.app.container)
    bottom_frame.pack(fill="x", padx=20, pady=10)

    btn_home = ttk.Button(
        bottom_frame, text="返回主選單", command=self.app.show_main_home
    )
    btn_home.pack(side="left", padx=5)

    btn_remove = ttk.Button(
        bottom_frame, text="移除選取檔案 (X)", command=self.remove_selected_file
    )
    btn_remove.pack(side="left", padx=5)

    btn_range = ttk.Button(
        bottom_frame,
        text="⚙ 取值範圍設定",
        command=lambda: self.app.open_range_dialog(
            [f["path"] for f in self.file_list]
        ),
    )
    btn_range.pack(side="left", padx=5)

    btn_next = ttk.Button(
        bottom_frame,
        text="下一步 (管理與排序項目)",
        style="Accent.TButton",
        command=self.go_to_step2,
    )
    btn_next.pack(side="right")

    self.refresh_step1_ui()

  def refresh_step1_ui(self):
    for item in self.tree_step1.get_children():
      self.tree_step1.delete(item)
    for idx, f in enumerate(self.file_list):
      self.tree_step1.insert(
          "", "end", iid=str(idx), values=(f["filename"], f["date"], f["path"])
      )

  def add_files_logic(self, file_paths):
    for path in file_paths:
      if not path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        continue
      filename = os.path.basename(path)
      folder_name = os.path.basename(os.path.dirname(path))

      date_str = self.app.parse_date_from_string(filename)
      if not date_str:
        date_str = self.app.parse_date_from_string(folder_name)
      if not date_str:
        user_choice = self.app.request_manual_date(filename, folder_name)
        if user_choice == "SKIP" or user_choice is None:
          continue
        date_str = user_choice

      self.file_list.append(
          {"path": path, "filename": filename, "date": date_str}
      )
    self.refresh_step1_ui()

  def add_excel_files(self):
    files = filedialog.askopenfilenames(
        title="選擇 Excel 檔案",
        filetypes=[
            ("Excel 檔案", "*.xlsx *.xls *.xlsm"),
            ("所有檔案", "*.*"),
        ],
    )
    if files:
      self.add_files_logic(files)

  def add_excel_folder(self):
    folder = filedialog.askdirectory(title="選擇包含 Excel 的資料夾")
    if folder:
      matched_files = []
      try:
        for root, dirs, filenames in os.walk(folder):
          for file in filenames:
            if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
              matched_files.append(os.path.join(root, file))
      except Exception as e:
        messagebox.showerror("錯誤", f"讀取資料夾時發生錯誤：\n{e}")
        return

      if matched_files:
        self.add_files_logic(matched_files)
        messagebox.showinfo(
            "加入完成", f"已從資料夾中尋找到 {len(matched_files)} 個 Excel 檔案！"
        )
      else:
        messagebox.showwarning(
            "提示", "該資料夾及其子資料夾中找不到任何支援的 Excel 檔案。"
        )

  def remove_selected_file(self):
    selected = self.tree_step1.selection()
    if not selected:
      messagebox.showwarning("提示", "請先選取要移除的檔案！")
      return
    for item in selected:
      idx = int(item)
      self.file_list[idx] = None
    self.file_list = [f for f in self.file_list if f is not None]
    self.refresh_step1_ui()

  def go_to_step2(self):
    if not self.file_list:
      messagebox.showerror("錯誤", "請至少加入一個 Excel 檔案才能進行下一步！")
      return
    self.show_step2_management()

  def show_step2_management(self):
    self.app.clear_container()

    lbl_title = ttk.Label(
        self.app.container,
        text="【數據整理】步驟二：管理項目順序、合併與刪除工作表",
        style="H2.TLabel",
    )
    lbl_title.pack(pady=10)

    all_sheets = set()
    for f in self.file_list:
      try:
        wb = openpyxl.load_workbook(f["path"], read_only=True, data_only=True)
        for sname in wb.sheetnames:
          all_sheets.add(sname)
        wb.close()
      except Exception as e:
        messagebox.showwarning(
            "讀取警告", f"無法讀取檔案 {f['filename']}：\n{e}"
        )

    setting_file_path = os.path.join(app_dir(), "setting.json")
    loaded_from_setting = False

    if os.path.exists(setting_file_path):
      try:
        with open(setting_file_path, "r", encoding="utf-8") as sf:
          saved_mapping = json.load(sf)
          if isinstance(saved_mapping, dict) and len(saved_mapping) > 0:
            self.item_mapping = saved_mapping
            loaded_from_setting = True
      except Exception:
        loaded_from_setting = False

    if not loaded_from_setting or not self.item_mapping:
      self.item_mapping = {sname: [sname] for sname in sorted(list(all_sheets))}

    lbl_info = ttk.Label(
        self.app.container,
        text=(
            "在此可調整項目上下順序、合併項目，或刪除不需要的項目。\n"
            "(提示：設定會自動儲存至 setting.json)"
        ),
        style="Hint.TLabel",
    )
    lbl_info.pack(anchor="w", padx=20, pady=5)

    main_mgmt_frame = ttk.Frame(self.app.container)
    main_mgmt_frame.pack(fill="both", expand=True, padx=20, pady=5)

    # 【UI 美化】Listbox 是舊式 tk 元件，手動設定顏色與邊框讓它跟主題一致
    self.listbox_items = tk.Listbox(
        main_mgmt_frame,
        selectmode=tk.EXTENDED,
        font=self.app.font_list,
        bg=PAL["card"],
        fg=self.app.theme_fg,
        selectbackground=self.app.theme_select_bg,
        selectforeground=PAL["sel_fg"],
        activestyle="none",
        relief="flat",
        borderwidth=0,
        highlightthickness=1,
        highlightbackground=PAL["card_border"],
        highlightcolor=self.app.theme_select_bg,
    )
    self.listbox_items.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        main_mgmt_frame, orient="vertical", command=self.listbox_items.yview
    )
    scrollbar.pack(side="left", fill="y")
    self.listbox_items.configure(yscrollcommand=scrollbar.set)

    self.refresh_listbox()

    btn_ctrl_frame = ttk.Frame(main_mgmt_frame)
    btn_ctrl_frame.pack(side="right", fill="y", padx=10)

    btn_up = ttk.Button(
        btn_ctrl_frame, text="⬆ 往上移動", command=self.move_item_up
    )
    btn_up.pack(fill="x", pady=5)

    btn_down = ttk.Button(
        btn_ctrl_frame, text="⬇ 往下移動", command=self.move_item_down
    )
    btn_down.pack(fill="x", pady=5)

    sep1 = ttk.Separator(btn_ctrl_frame, orient="horizontal")
    sep1.pack(fill="x", pady=10)

    btn_merge = ttk.Button(
        btn_ctrl_frame, text="🔗 合併選取項目", command=self.merge_selected_items
    )
    btn_merge.pack(fill="x", pady=5)

    btn_delete = ttk.Button(
        btn_ctrl_frame, text="❌ 刪除選取項目", command=self.delete_selected_items
    )
    btn_delete.pack(fill="x", pady=5)

    bottom_frame = ttk.Frame(self.app.container)
    bottom_frame.pack(fill="x", padx=20, pady=10)

    btn_back = ttk.Button(
        bottom_frame, text="上一步", command=self.show_step1_home
    )
    btn_back.pack(side="left")

    btn_next = ttk.Button(
        bottom_frame,
        text="下一步 (開始計算並輸出)",
        style="Accent.TButton",
        command=self.go_to_step3,
    )
    btn_next.pack(side="right")

  def refresh_listbox(self):
    self.listbox_items.delete(0, tk.END)
    for key in self.item_mapping.keys():
      self.listbox_items.insert(tk.END, key)

  def move_item_up(self):
    selected_indices = self.listbox_items.curselection()
    if not selected_indices:
      messagebox.showwarning("提示", "請先選取要移動的項目！")
      return
    items = list(self.item_mapping.items())
    for idx in selected_indices:
      if idx > 0:
        items[idx], items[idx - 1] = items[idx - 1], items[idx]
        self.item_mapping = dict(items)
        self.refresh_listbox()
        self.listbox_items.selection_set(idx - 1)
        self.listbox_items.selection_clear(idx)
        break

  def move_item_down(self):
    selected_indices = self.listbox_items.curselection()
    if not selected_indices:
      messagebox.showwarning("提示", "請先選取要移動的項目！")
      return
    items = list(self.item_mapping.items())
    for idx in reversed(selected_indices):
      if idx < len(items) - 1:
        items[idx], items[idx + 1] = items[idx + 1], items[idx]
        self.item_mapping = dict(items)
        self.refresh_listbox()
        self.listbox_items.selection_set(idx + 1)
        self.listbox_items.selection_clear(idx)
        break

  def delete_selected_items(self):
    selected_indices = self.listbox_items.curselection()
    if not selected_indices:
      messagebox.showwarning("提示", "請先選取要刪除的項目！")
      return

    items_keys = list(self.item_mapping.keys())
    selected_names = [items_keys[i] for i in selected_indices]

    if len(selected_names) == len(self.item_mapping):
      if not messagebox.askyesno(
          "警告",
          "你正準備刪除所有的項目！確定要清空嗎？",
          icon="warning",
      ):
        return

    for name in selected_names:
      if name in self.item_mapping:
        del self.item_mapping[name]

    self.refresh_listbox()
    messagebox.showinfo("成功", "已成功刪除選取的項目。")

  def merge_selected_items(self):
    selected_indices = self.listbox_items.curselection()
    if len(selected_indices) < 2:
      messagebox.showwarning("提示", "請至少選取兩個以上的項目進行合併！")
      return

    items_keys = list(self.item_mapping.keys())
    selected_names = [items_keys[i] for i in selected_indices]

    top = tk.Toplevel(self.app)
    top.title("合併項目與選擇代表名稱")
    self.app.center_dialog(top, 460, 260)
    top.grab_set()

    body = ttk.Frame(top, padding=15)
    body.pack(fill="both", expand=True)

    ttk.Label(
        body,
        text=(
            f"準備合併以下項目：\n{', '.join(selected_names)}\n\n"
            "請選擇其中一個作為最終的代表名稱："
        ),
        wraplength=self.app.px(420),
    ).pack(anchor="w")

    combo_var = tk.StringVar(value=selected_names[0])
    combo_names = ttk.Combobox(
        body,
        textvariable=combo_var,
        values=selected_names,
        state="readonly",
        width=30,
    )
    combo_names.pack(pady=10)

    def on_confirm():
      chosen_name = combo_var.get()
      if not chosen_name:
        messagebox.showerror("錯誤", "請選擇一個代表名稱！", parent=top)
        return

      combined_sheets = []
      for name in selected_names:
        combined_sheets.extend(self.item_mapping[name])

      new_mapping = {}
      for k, v in self.item_mapping.items():
        if k not in selected_names:
          new_mapping[k] = v
        elif k == selected_names[0]:
          new_mapping[chosen_name] = combined_sheets

      self.item_mapping = new_mapping
      self.refresh_listbox()
      top.destroy()
      messagebox.showinfo("成功", f"已成功將項目合併，代表名稱為：{chosen_name}")

    ttk.Button(
        body, text="確認合併", style="Accent.TButton", command=on_confirm
    ).pack(pady=10)

  def go_to_step3(self):
    if not self.item_mapping:
      messagebox.showerror("錯誤", "目前沒有任何有效的項目可供輸出！")
      return
    self.show_step3_output()

  def show_step3_output(self):
    self.app.clear_container()

    lbl_title = ttk.Label(
        self.app.container,
        text="【數據整理】步驟三：自動搜尋儲存格、計算與矩陣匯出",
        style="H2.TLabel",
    )
    lbl_title.pack(pady=10)

    mode_text = (
        "電阻模式 (=[數字]*...)"
        if self.mode_var.get() == "RESISTOR"
        else "平均值模式 (=AVERAGE(...)"
    )

    txt_info = ttk.Label(
        self.app.container,
        text=(
            f"目前選擇模式：【{mode_text}】\n"
            "已完成所有設定與項目管理。\n"
            "請選擇下方匯出格式，並點擊按鈕開始運算與匯出。"
        ),
        justify="left",
    )
    txt_info.pack(padx=20, pady=5)

    # 匯出格式選擇選項
    format_frame = ttk.LabelFrame(
        self.app.container, text="選擇匯出檔案格式", padding=12
    )
    format_frame.pack(fill="x", padx=20, pady=10)

    self.export_format_var = tk.StringVar(value="xlsx")
    ttk.Radiobutton(
        format_frame,
        text="Excel 試算表 (*.xlsx)",
        value="xlsx",
        variable=self.export_format_var,
    ).pack(anchor="w", pady=3)
    ttk.Radiobutton(
        format_frame,
        text="CSV 純文字表格 (*.csv)",
        value="csv",
        variable=self.export_format_var,
    ).pack(anchor="w", pady=3)
    ttk.Radiobutton(
        format_frame,
        text="Markdown 表格 (*.md)",
        value="md",
        variable=self.export_format_var,
    ).pack(anchor="w", pady=3)

    btn_run = ttk.Button(
        self.app.container,
        text="開始計算並儲存",
        style="Accent.TButton",
        command=self.process_and_export,
    )
    btn_run.pack(pady=10, ipady=4)

    btn_back_frame = ttk.Frame(self.app.container)
    btn_back = ttk.Button(
        btn_back_frame, text="上一步", command=self.show_step2_management
    )
    btn_back.pack()
    btn_back_frame.pack(pady=5)

  def process_and_export(self):
    try:
      setting_file_path = os.path.join(app_dir(), "setting.json")
      try:
        with open(setting_file_path, "w", encoding="utf-8") as sf:
          json.dump(self.item_mapping, sf, ensure_ascii=False, indent=4)
      except Exception as se:
        print(f"儲存 setting.json 失敗: {se}")

      matrix_data = {}
      source = "rx" if self.mode_var.get() == "RESISTOR" else "a0"
      conflict_reports = []
      range_warns = []

      infos = self.app.load_infos(
          [f["path"] for f in self.file_list], title="讀取 Excel 資料中…"
      )
      if infos is None:
        return

      for f in self.file_list:
        f_path = f["path"]
        dt_obj = datetime.strptime(f["date"], "%Y-%m-%d")
        date_col_name = dt_obj.strftime("%Y/%m/%d")

        if date_col_name not in matrix_data:
          matrix_data[date_col_name] = {}

        info = infos[f_path]
        if info.get("error"):
          messagebox.showerror(
              "錯誤", f"處理檔案 {f['filename']} 時發生錯誤：\n{info['error']}"
          )
          return
        by_name = {s["name"]: s for s in info["sheets"]}

        for final_item_name, raw_sheets in self.item_mapping.items():
          found_values_count = 0
          collected_values = []

          for sname in raw_sheets:
            sh = by_name.get(sname)
            if sh is None:
              continue
            v = get_values(sh, self.app.range_settings.get(range_key(f_path, sname)))
            for w in v["warn"]:
              range_warns.append(f"{f['filename']}／{sname}：{w}")
            if v[source] is not None:
              collected_values.append((sname, v[source]))
              found_values_count += 1

          if found_values_count > 0:
            chosen_val = collected_values[0][1]

            if found_values_count > 1:
              sub_sheets_str = ", ".join([item[0] for item in collected_values])
              conflict_reports.append(
                  f"日期: {date_col_name} | 檔案: {f['filename']} | 項目"
                  f" '{final_item_name}' 偵測到多筆重複數據"
                  f" ({sub_sheets_str})，已依規則採用代表項目數據。"
              )

            matrix_data[date_col_name][final_item_name] = chosen_val
          else:
            matrix_data[date_col_name][final_item_name] = ""

      df = pd.DataFrame(matrix_data)

      final_index_order = list(self.item_mapping.keys())
      for sname in final_index_order:
        if sname not in df.index:
          df.loc[sname] = ""

      df = df.reindex(index=final_index_order)
      df = df.reindex(sorted(df.columns), axis=1)

      df.index.name = "項目"
      df = df.reset_index()

      export_format = (
          self.export_format_var.get()
          if hasattr(self.app, "export_format_var")
          else "xlsx"
      )

      if export_format == "csv":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 檔案", "*.csv")],
            title="選擇儲存 CSV 檔案位置與檔名",
        )
        if not save_path:
          return
        df.to_csv(save_path, index=False, encoding="utf-8-sig")
      elif export_format == "md":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown 檔案", "*.md")],
            title="選擇儲存 Markdown 檔案位置與檔名",
        )
        if not save_path:
          return
        headers = list(df.columns)
        header_str = " | ".join(str(h) for h in headers)
        sep_str = " | ".join("---" for _ in headers)
        md_lines = [f"| {header_str} |", f"| {sep_str} |"]
        for _, row in df.iterrows():
          row_str = " | ".join(
              str(val) if pd.notna(val) and val != "" else "" for val in row
          )
          md_lines.append(f"| {row_str} |")

        with open(save_path, "w", encoding="utf-8") as md_f:
          md_f.write("\n".join(md_lines))
      else:
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 檔案", "*.xlsx")],
            title="選擇儲存 Excel 檔案位置與檔名",
        )
        if not save_path:
          return
        df.to_excel(save_path, index=False)

      if range_warns:
        uniq = list(dict.fromkeys(range_warns))
        more = f"\n...以及其他 {len(uniq) - 12} 筆" if len(uniq) > 12 else ""
        messagebox.showwarning(
            "取值範圍警告",
            "以下工作表的範圍有問題，數值已留空：\n\n"
            + "\n".join(uniq[:12])
            + more,
        )

      if conflict_reports:
        warning_msg = (
            "資料已成功計算與匯出，設定檔已自動儲存。\n"
            "但發現以下項目在寫入時有重複數據碰撞：\n\n"
            + "\n".join(conflict_reports[:10])
        )
        if len(conflict_reports) > 10:
          warning_msg += (
              f"\n...以及其他 {len(conflict_reports) - 10} 筆重複記錄。"
          )
        messagebox.showwarning("資料重複警告提示", warning_msg)
      else:
        messagebox.showinfo(
            "成功",
            "資料已成功排序、合併、輸出 (日期格式 YYYY/MM/DD)，\n"
            f"設定檔已儲存至：{setting_file_path}\n"
            f"儲存路徑：\n{save_path}",
        )

    except Exception as e:
      messagebox.showerror("重大錯誤", f"輸出過程發生未預期的錯誤：\n{e}")
