import tkinter as tk
from os import system as cmd
from PIL import Image, ImageTk


class FullscreenImageApp:

    def __init__(self, root, image_path):
        self.root = root
        self.image_path = image_path

        # 1. 設定視窗全螢幕並隱藏滑鼠游標（增加沉浸感）
        self.root.attributes("-fullscreen", True)
        self.root.config(cursor="none")

        # 2. 取得螢幕的寬高
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 3. 載入圖片並縮放到符合螢幕大小
        try:
            pil_image = Image.open(self.image_path)
            # 使用高品質濾鏡縮放圖片
            pil_image = pil_image.resize(
                (screen_width, screen_height), Image.Resampling.LANCZOS
            )
            self.img = ImageTk.PhotoImage(pil_image)
        except Exception as e:
            print(f"圖片載入失敗，請檢查路徑或檔名是否正確: {e}")
            self.root.destroy()
            return

        # 4. 建立畫布來放置圖片
        self.canvas = tk.Canvas(
            root, width=screen_width, height=screen_height, highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)
        self.canvas.create_image(0, 0, anchor="nw", image=self.img)

        # 5. 在右上角建立倒數計時文字
        self.time_left = 3
        self.countdown_text = self.canvas.create_text(
            screen_width - 80,
            60,
            text=str(self.time_left),
            fill="red",
            font=("Arial", 50, "bold"),
        )

        # 6. 綁定按鍵：Caps Lock 與 Esc 作為緊急關閉
        self.root.bind("<Caps_Lock>", self.emergency_close)
        self.root.bind("<Key>", self.check_key)

        # 7. 開始倒數計時
        self.update_countdown()

    def update_countdown(self):
        """每秒更新倒數時間，時間到就關閉程式"""
        if self.time_left > 0:
            self.canvas.itemconfig(
                self.countdown_text, text=str(self.time_left)
            )
            self.time_left -= 1
            self.root.after(1000, self.update_countdown)  # 1000 毫秒 = 1 秒
        else:
            self.root.destroy()

    def emergency_close(self, event):
        """Caps Lock 觸發的關閉函式"""
        self.root.destroy()

    def check_key(self, event):
        """備用按鍵檢查（如果按 Esc 也能關閉）"""
        if event.keysym == "Escape":
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()

    # ⚠️ 請把底部的 "your_image.jpg" 改成你要顯示的圖片檔名或完整路徑
    image_file = "丁丁.jpg"

    app = FullscreenImageApp(root, image_file)
    root.mainloop()
    cmd("git add . && git commit -m \"Backup\" && git push -u origin main")
    input("按下 Enter 關閉")