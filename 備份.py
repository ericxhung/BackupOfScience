import glob
import os
import random
import tkinter as tk
from PIL import Image, ImageTk


class SlideshowApp:

    def __init__(self, root):
        self.root = root

        # 1. 設定全螢幕與隱藏滑鼠游標
        self.root.attributes("-fullscreen", True)
        self.root.config(cursor="none")

        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()

        # 2. 自動讀取 ./image 資料夾底下的所有圖片
        image_extensions = ("*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp")
        self.image_paths = []
        for ext in image_extensions:
            self.image_paths.extend(glob.glob(os.path.join("image", ext)))

        # 安全檢查：如果資料夾是空的就直接結束
        if not self.image_paths:
            print("錯誤：在 './image' 資料夾中找不到任何圖片檔案！")
            self.root.destroy()
            return

        # 隨機打亂圖片順序
        random.shuffle(self.image_paths)
        self.current_index = 0

        # 3. 建立畫布
        self.canvas = tk.Canvas(
            root,
            width=self.screen_width,
            height=self.screen_height,
            highlightthickness=0,
        )
        self.canvas.pack(fill="both", expand=True)

        # 4. 設定總倒數時間（5秒）
        self.time_left = 5
        self.countdown_text = self.canvas.create_text(
            self.screen_width - 80,
            60,
            text=str(self.time_left),
            fill="red",
            font=("Arial", 50, "bold"),
        )

        # 5. 綁定緊急關閉按鍵（Caps Lock 與 Esc）
        self.root.bind("<Caps_Lock>", self.emergency_close)
        self.root.bind("<Key>", self.check_key)

        # 6. 開始顯示第一張圖片並啟動 5 秒倒數計時
        self.show_next_image()
        self.update_timer()

    def show_next_image(self):
        """顯示目前的圖片，每秒切換一張，若輪完則循環"""
        if not self.image_paths:
            return

        img_path = self.image_paths[self.current_index]
        try:
            pil_image = Image.open(img_path)
            pil_image = pil_image.resize(
                (self.screen_width, self.screen_height),
                Image.Resampling.LANCZOS,
            )
            self.img = ImageTk.PhotoImage(pil_image)

            # 更新或建立畫布上的圖片
            if hasattr(self, "image_item"):
                self.canvas.itemconfig(self.image_item, image=self.img)
            else:
                self.image_item = self.canvas.create_image(
                    0, 0, anchor="nw", image=self.img
                )
                # 確保右上角倒數文字永遠浮在圖片上方
                self.canvas.tag_raise(self.countdown_text)

        except Exception as e:
            print(f"無法載入圖片 {img_path}: {e}")

        # 移動到下一張圖片索引（若超過清單長度則用 % 讓它從頭循環）
        self.current_index = (self.current_index + 1) % len(self.image_paths)

        # 只要總時間還沒到 5 秒，就排程在 1 秒後換下一張圖
        if self.time_left > 0:
            self.root.after(1000, self.show_next_image)

    def update_timer(self):
        """每秒更新右上角的 5 秒倒數，時間到自動關閉"""
        if self.time_left > 0:
            self.canvas.itemconfig(
                self.countdown_text, text=str(self.time_left)
            )
            self.time_left -= 1
            self.root.after(1000, self.update_timer)
        else:
            self.root.destroy()

    def emergency_close(self, event):
        """Caps Lock 觸發緊急關閉"""
        self.root.destroy()

    def check_key(self, event):
        """Esc 備用關閉"""
        if event.keysym == "Escape":
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = SlideshowApp(root)
    root.mainloop()
    os.system("git add . && git commit -m \"Backup\" && git push -u origin main")
    input("按下 Enter 關閉")
