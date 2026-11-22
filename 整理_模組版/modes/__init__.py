"""自動載入 modes 資料夾裡的所有模式。

規則：modes 資料夾裡每個 .py 檔（檔名不是 _ 開頭、也不是 base.py），
只要裡面有繼承 Mode 的類別，就會自動出現在主頁。不用改任何其他檔案。
某個模式載入失敗時，主頁會顯示警告，不會讓整個程式打不開。
"""

import importlib
import pkgutil

from modes.base import Mode

MODE_CLASSES = []
LOAD_ERRORS = []  # [(檔名, 錯誤訊息)]

for _info in sorted(pkgutil.iter_modules(__path__), key=lambda m: m.name):
  if _info.name.startswith("_") or _info.name == "base":
    continue
  try:
    _mod = importlib.import_module(f"modes.{_info.name}")
  except Exception as _e:
    LOAD_ERRORS.append((_info.name, f"{type(_e).__name__}: {_e}"))
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
