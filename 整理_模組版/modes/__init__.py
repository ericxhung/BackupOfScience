"""自動載入 modes 資料夾裡的所有模式。

規則：modes 資料夾裡每個 .py 檔（檔名不是 _ 開頭、也不是 base.py），
只要裡面有繼承 Mode 的類別，就會自動出現在主頁。不用改任何其他檔案。
某個模式載入失敗時，主頁會顯示警告，不會讓整個程式打不開。
"""

import importlib
import pkgutil
import sys

from modes.base import Mode

MODE_CLASSES = []
LOAD_ERRORS = []  # [(檔名, 錯誤訊息)]

_names = {m.name for m in pkgutil.iter_modules(__path__)}
if getattr(sys, "frozen", False):
  # 打包成 exe 之後，保底把內建的三個模式也列進去（打包時已用 --collect-submodules modes 收進來）
  _names |= {"print_mode", "organize_mode", "view_mode"}

for _name in sorted(_names):
  if _name.startswith("_") or _name == "base":
    continue
  try:
    _mod = importlib.import_module(f"modes.{_name}")
  except Exception as _e:
    LOAD_ERRORS.append((_name, f"{type(_e).__name__}: {_e}"))
    continue
  for _obj in vars(_mod).values():
    if (
        isinstance(_obj, type)
        and issubclass(_obj, Mode)
        and _obj is not Mode
        and _obj.__module__ == _mod.__name__
    ):
      MODE_CLASSES.append(_obj)

MODE_CLASSES.sort(key=lambda c: (c.order, c.__name__))
