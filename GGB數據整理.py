import tkinter as tk
from tkinter import messagebox

def generate_ggb_list():
    # 取得使用者輸入的名稱與表格內容
    name = entry_name.get().strip()
    raw_data = text_data.get("1.0", tk.END).strip()
    
    if not name:
        messagebox.showerror("錯誤", "請輸入串列名稱！")
        return
        
    if not raw_data:
        messagebox.showerror("錯誤", "請輸入表格數據！")
        return
        
    # 將輸入按行分割，並過濾掉空白行
    lines = [line.strip() for line in raw_data.split("\n") if line.strip()]
    
    if len(lines) == 0:
        messagebox.showerror("錯誤", "沒有偵測到任何數據行！")
        return
        
    try:
        points = []
        # 逐行解析每一筆資料 (x, y, 電壓)
        for i, line in enumerate(lines, start=1):
            parts = [p.strip() for p in line.split(",")]
            
            # 檢查這一行是否剛好有 3 個欄位 (x, y, z)
            if len(parts) != 3:
                messagebox.showerror("錯誤", f"格式錯誤：第 {i} 行「{line}」的數據量不正確！每一行必須剛好有 3 個數值（以逗號分隔）。")
                return
                
            x, y, z = parts
            points.append(f"({x},{y},{z})")
            
        # 組合出標準格式：[name]={(x1,y1,z1),(x2,y2,z2),......,(xn,yn,zn)}
        result_str = f"{name}=" + "{" + ",".join(points) + "}"
        
        # 開啟結果新視窗
        open_result_window(result_str)
        
    except Exception as e:
        messagebox.showerror("錯誤", f"解析資料時發生未預期的錯誤：{e}")

def open_result_window(result_str):
    # 建立新視窗
    win = tk.Toplevel(root)
    win.title("GeoGebra 串列結果")
    win.geometry("550x240")
    win.resizable(False, False)
    
    # 最上方提示文字
    lbl_tip = tk.Label(win, text="複製後貼入GGB按Enter", fg="blue", font=("Microsoft JhengHei", 12, "bold"))
    lbl_tip.pack(pady=10)
    
    # 顯示結果的文字框（唯讀）
    txt_output = tk.Text(win, height=4, width=60, font=("Consolas", 10))
    txt_output.pack(pady=5)
    txt_output.insert(tk.END, result_str)
    txt_output.config(state="disabled")
    
    # 複製按鈕功能
    def copy_to_clipboard():
        root.clipboard_clear()
        root.clipboard_append(result_str)
        messagebox.showinfo("提示", "已成功複製到剪貼簿！")
        
    btn_copy = tk.Button(win, text="複製", command=copy_to_clipboard, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 11, "bold"), width=15)
    btn_copy.pack(pady=10)

# 建立主視窗
root = tk.Tk()
root.title("GGB 座標串列轉換器")
root.geometry("450x380")
root.resizable(False, False)

# 1. 串列名稱輸入區
frame_name = tk.Frame(root)
frame_name.pack(pady=15)
lbl_name = tk.Label(frame_name, text="串列名稱 [name]：", font=("Microsoft JhengHei", 10))
lbl_name.pack(side=tk.LEFT)
entry_name = tk.Entry(frame_name, width=20, font=("Microsoft JhengHei", 10))
entry_name.pack(side=tk.LEFT)
entry_name.insert(0, "list1") # 預設名稱

# 2. 中間表格格式輸入區
lbl_format_tip = tk.Label(root, text="請輸入表格格式 (每行一筆，格式為 x,y,電壓)：", font=("Microsoft JhengHei", 9), fg="gray")
lbl_format_tip.pack(pady=(0, 5))

text_data = tk.Text(root, height=6, width=45, font=("Consolas", 10))
text_data.pack(pady=5)
# 提供範例預填文字（示範多行點資料）
text_data.insert(tk.END, "1,2,3\n4,5,6\n7,8,9")

# 3. 下方完成按鈕
btn_done = tk.Button(root, text="完成", command=generate_ggb_list, bg="#2196F3", fg="white", font=("Microsoft JhengHei", 11, "bold"), width=15)
btn_done.pack(pady=15)

# 啟動主迴圈
root.mainloop()