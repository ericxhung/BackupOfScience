import tkinter as tk
from tkinter import filedialog, simpledialog
import win32com.client as win32
import os

def process_excel_safe():
    # 隱藏 tkinter 主視窗
    root = tk.Tk()
    root.withdraw()
    
    # 1. 彈出檔案選擇視窗
    file_path = filedialog.askopenfilename(
        title="請選擇要處理的 Excel 檔案",
        filetypes=[("Excel files", "*.xlsx;*.xlsm")]
    )
    
    if not file_path:
        print("沒有選擇任何檔案，程式已取消。")
        return

    # 2. 彈出輸入視窗，讓使用者輸入 [number]
    number = simpledialog.askfloat("輸入數值", "請輸入要代入公式的數字 ([number]):")
    if number is None:
        print("沒有輸入數字，程式已取消。")
        return

    # 將路徑轉為 Windows 絕對路徑，避免 COM 元件找不到檔案
    file_path = os.path.abspath(file_path)

    excel = None
    wb = None
    try:
        # 3. 啟動 Excel 應用程式（在背景執行，不顯示畫面）
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False       # 不顯示 Excel 視窗介面
        excel.DisplayAlerts = False # 自動覆蓋或忽略提示
        
        # 4. 開啟 Excel 檔案
        wb = excel.Workbooks.Open(file_path)
        
        # 5. 走訪每一個工作表
        for sheet in wb.Sheets:
            sheet_name = sheet.Name
            
            # 檢查工作表名稱是否為 '溫度' 或 '氣溫'
            if sheet_name in ['溫度', '氣溫']:
                print(f"跳過工作表：{sheet_name}")
                continue
            
            # 在 I2 寫入平均公式
            sheet.Range("I2").Formula = "=AVERAGE(D2:D21)"
            
            # 在 I3 寫入自定義公式，帶入使用者輸入的數字
            sheet.Range("I3").Formula = f"={number}*(1023-I2)/I2"
            
            print(f"已成功處理工作表：{sheet_name}")
            
        # 6. 儲存並關閉檔案
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        
        print(f"所有工作表處理完成！檔案已安全儲存至：{file_path}")
        
    except Exception as e:
        # 如果發生錯誤，確保背景的 Excel 程序有被關掉，避免卡在記憶體裡
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
            except:
                pass
        print(f"執行過程中發生錯誤：{e}")

if __name__ == "__main__":
    process_excel_safe()