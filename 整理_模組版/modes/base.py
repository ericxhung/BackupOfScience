"""所有模式的共同父類別。新增模式請看「新增模式手冊.md」。"""


class Mode:
  """一個「模式」＝主頁上的一個按鈕，加上按下去之後的所有畫面。

  子類別要做的事只有三件：
    1. 設定 title（按鈕文字）與 order（排序，數字小的在前面）。
    2. 在 __init__ 裡準備自己的狀態（記得先 super().__init__(app)）。
    3. 實作 start()：主頁按下按鈕時被呼叫，在這裡畫出第一個畫面。

  self.app 是主視窗，可以用它的共用服務，例如：
    self.app.clear_container()      清空畫面
    self.app.show_main_home()       回主頁
    self.app.container              放畫面的容器（ttk.Frame）
    self.app.load_infos(paths)      讀取 Excel 資訊（有進度視窗、有快取）
    self.app.px(n)                  依縮放換算像素
  """

  title = "新模式"
  order = 100

  def __init__(self, app):
    self.app = app

  def start(self):
    raise NotImplementedError("請在子類別實作 start()")
