"""模式範本：複製這個檔案，改成別的檔名（不要用 _ 開頭），就會多一個模式。

步驟：
  1. 複製本檔，改檔名，例如 my_mode.py（檔名只用英文、數字、底線）。
  2. 把下面的類別名稱 HelloMode 改成你的名字（例如 MyMode）。
  3. 改 title（主頁上的按鈕文字）和 order（數字小的排在前面）。
  4. 在 start() 裡畫出你的畫面。
  5. 重新執行 main.py，主頁就會多一個按鈕。

注意：檔名以 _ 開頭的檔案不會被載入，所以這個範本本身不會出現在主頁。
"""

import tkinter as tk
from tkinter import messagebox, ttk

from modes.base import Mode


class HelloMode(Mode):
  title = "👋 打招呼範例"  # 主頁按鈕上的文字
  order = 90  # 排序：10 列印、20 整理、30 檢視，所以 90 會排在最後

  def __init__(self, app):
    super().__init__(app)
    # 這個模式自己的狀態放在這裡。tk.StringVar 要指定 master=app
    self.name_var = tk.StringVar(master=app, value="")

  def start(self):
    """主頁按下按鈕時會來到這裡。"""
    app = self.app
    app.clear_container()  # 先把畫面清空

    ttk.Label(app.container, text="打招呼範例", style="H2.TLabel").pack(pady=(10, 6))
    ttk.Label(
        app.container, text="輸入名字，按「打招呼」。", style="Hint.TLabel"
    ).pack()
    ttk.Entry(app.container, textvariable=self.name_var, width=24).pack(pady=14)

    row = ttk.Frame(app.container)
    row.pack()
    ttk.Button(
        row, text="打招呼", style="Accent.TButton", command=self.say_hello
    ).pack(side="left", padx=6)
    ttk.Button(row, text="返回主選單", command=app.show_main_home).pack(
        side="left", padx=6
    )

  def say_hello(self):
    name = self.name_var.get().strip() or "朋友"
    # 對話框要指定 parent=self.app，才會顯示在主視窗上面
    messagebox.showinfo("你好", f"你好，{name}！", parent=self.app)
