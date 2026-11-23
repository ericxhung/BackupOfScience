"""共用取值層（讀 Excel、算 Rx、取值範圍、快取）＋色票／設定檔。不含任何視窗元件。"""

import ctypes
import json
import os
import re
import string
import sys
import openpyxl
import array
import bisect
import math
import pickle
import statistics
import unicodedata



# 【UI 美化】Windows 高 DPI 設定：避免在高解析度螢幕上字體模糊
def enable_high_dpi():
  if sys.platform == "win32":
    try:
      ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
      try:
        ctypes.windll.user32.SetProcessDPIAware()
      except Exception:
        pass


class _LazyPandas:
  """第一次用到 pd.xxx 時才載入 pandas（載入要一兩秒，不要拖慢啟動）。"""

  def __getattr__(self, name):
    import pandas as _pd

    return getattr(_pd, name)


pd = _LazyPandas()


def app_dir():
  """設定檔、快取檔的存放資料夾。
  - 打包成 exe 時：exe 所在的資料夾（不管從哪裡開啟捷徑都一樣）。
  - 一般用 python 執行時：目前工作資料夾（和以前的行為相同）。
  """
  if getattr(sys, "frozen", False):
    return os.path.dirname(os.path.abspath(sys.executable))
  return os.getcwd()


# ==================== 共用取值層（列印／整理／檢視 三個模式共用）====================
# 做法：不再靠「第一個符合的格子」去猜，而是看公式的意思：
#   - 找到 =AVERAGE(某欄範圍) → 看它平均的是 A0 欄還是 R(電阻) 欄
#   - 平均的是 A0 欄 → Rx 取「引用這格的 =k*(1023-X)/X 公式」
#   - 平均的是 R 欄   → 這格本身就是 Rx，A0 另外從 A0 欄用同一段列範圍算
# 找不到可辨識的版型時，退回原本「第一個符合的格子」的做法（行為跟舊版一樣）。

CACHE_VERSION = 1
GRID_ROWS = 30  # 跟原本程式掃描範圍一致：1~30 列
GRID_COLS = 16  # 跟原本程式掃描範圍一致：A~P 欄
COL_LETTERS = list(string.ascii_uppercase[:GRID_COLS])

_RE_RX_LEGACY = re.compile(r"^=\d+\*")
_RE_AVG_LEGACY = re.compile(r"^=AVERAGE\(", re.IGNORECASE)
_RE_AVG_RANGE = re.compile(
    r"^=AVERAGE\(\s*\$?([A-Za-z]+)\$?(\d+)\s*:\s*\$?([A-Za-z]+)\$?(\d+)\s*\)$",
    re.IGNORECASE,
)
_RE_RX_FORMULA = re.compile(
    r"^=(\d+(?:\.\d+)?)\*\(\s*1023\s*-\s*\$?([A-Za-z]+)\$?(\d+)\s*\)"
    r"\s*/\s*\$?([A-Za-z]+)\$?(\d+)$"
)
_RE_TEMP = re.compile(r"(\d+(?:\.\d+)?)\s*(?:C|℃)", re.IGNORECASE)
_RE_HUM_TEXT = re.compile(r"(\d+(?:\.\d+)?)\s*[%％]?")


def is_number(v):
  return (
      isinstance(v, (int, float))
      and not isinstance(v, bool)
      and not (isinstance(v, float) and math.isnan(v))
  )


def coord_to_index(coord):
  """'J2' -> (欄索引 9, 列索引 1)；格式不對回傳 None。索引都從 0 開始。"""
  m = re.match(r"^\$?([A-Za-z]+)\$?(\d+)$", str(coord).strip())
  if not m:
    return None
  ci = 0
  for ch in m.group(1).upper():
    ci = ci * 26 + (ord(ch) - 64)
  return ci - 1, int(m.group(2)) - 1


def grid_value(grid, coord):
  idx = coord_to_index(coord)
  if idx is None:
    return None
  ci, ri = idx
  if 0 <= ri < GRID_ROWS and 0 <= ci < GRID_COLS:
    return grid[ri][ci]
  return None


def number_or_error(v):
  """回傳 (數字或 None, 錯誤文字或 None)。Excel 錯誤值（#DIV/0! 等）會回傳錯誤文字。"""
  if is_number(v):
    return float(v), None
  if isinstance(v, str):
    s = v.strip()
    if s.startswith("#"):
      return None, s
    try:
      return float(s), None
    except ValueError:
      return None, None
  return None, None


def _fmt_plain(x):
  x = float(x)
  return str(int(x)) if x.is_integer() else str(round(x, 1))


def parse_hum_percent(v):
  """把濕度欄的值轉成百分比數字（0~1 視為比例）。無法判斷回傳 None。"""
  pct = None
  if is_number(v):
    pct = float(v)
  elif isinstance(v, str):
    s = v.strip()
    try:
      pct = float(s)
    except ValueError:
      m = _RE_HUM_TEXT.search(s)
      if m:
        pct = float(m.group(1))
  if pct is None:
    return None
  if 0 <= pct <= 1:
    pct = pct * 100
  if not (0 <= pct <= 100):
    return None
  return pct


def _legacy_hum_text(rv):
  """跟舊版一樣的濕度文字轉換（舊版行為原封不動保留）。"""
  if isinstance(rv, (int, float)):
    if 0 <= rv <= 1:
      return _fmt_plain(rv * 100)
    return str(rv)
  s = str(rv).strip()
  try:
    f_val = float(s)
    if 0 <= f_val <= 1:
      return _fmt_plain(f_val * 100)
    return str(f_val)
  except ValueError:
    m = _RE_HUM_TEXT.search(s)
    if m:
      f_val = float(m.group(1))
      if 0 <= f_val <= 1:
        return _fmt_plain(f_val * 100)
      return str(f_val)
  return ""


def legacy_first_reading(grid_d):
  """舊版做法：欄優先掃描，取第一個「xx C」當溫度，它右邊那格當濕度。"""
  auto_temp = ""
  auto_hum = ""
  for ci in range(GRID_COLS):
    for ri in range(GRID_ROWS):
      val = grid_d[ri][ci]
      if val is not None and not auto_temp:
        m = _RE_TEMP.search(str(val).strip())
        if m:
          try:
            auto_temp = str(int(float(m.group(1))))
          except ValueError:
            auto_temp = m.group(1)
          if ci + 1 < GRID_COLS:
            right = grid_d[ri][ci + 1]
            if right is not None:
              auto_hum = _legacy_hum_text(right)
      if auto_temp and auto_hum:
        break
    if auto_temp and auto_hum:
      break
  return auto_temp, auto_hum


TH_MODE_FIRST = "開頭第一筆（原本做法）"
TH_MODE_MEDIAN = "整個檔案所有紀錄的中位數"


def first_reading(info):
  """開頭第一筆：優先用溫度工作表；沒有的話用第一個有溫濕度紀錄的工作表。"""
  for sh in info["sheets"]:
    if sh["kind"] == "temp":
      return legacy_first_reading(sh["grid_d"])
  for sh in info["sheets"]:
    if sh["th"]:
      return legacy_first_reading(sh["grid_d"])
  return "", ""


def median_reading(info):
  """整個檔案所有有效溫濕度紀錄的中位數（會擋掉偶爾出現的怪值）。"""
  ts = []
  hs = []
  for sh in info["sheets"]:
    for t, h in sh["th"]:
      ts.append(t)
      hs.append(h)
  if not ts:
    return "", ""
  return _fmt_plain(statistics.median(ts)), _fmt_plain(statistics.median(hs))


def detect_layout(grid_f, hdr):
  """依公式意思辨識版型，回傳 dict；認不出來回傳 None。"""
  avg = None
  for ci in range(GRID_COLS):
    for ri in range(GRID_ROWS):
      v = grid_f[ri][ci]
      if isinstance(v, str):
        m = _RE_AVG_RANGE.match(v.strip())
        if m and m.group(1).upper() == m.group(3).upper():
          r1, r2 = int(m.group(2)), int(m.group(4))
          if r1 > r2:
            r1, r2 = r2, r1
          avg = {
              "cell": f"{COL_LETTERS[ci]}{ri + 1}",
              "col": m.group(1).upper(),
              "r1": r1,
              "r2": r2,
          }
          break
    if avg:
      break
  if not avg:
    return None

  # 版型一：有一格 =k*(1023-平均格)/平均格  → 平均的是 A0，Rx 就是那格
  for ci in range(GRID_COLS):
    for ri in range(GRID_ROWS):
      v = grid_f[ri][ci]
      if isinstance(v, str):
        m = _RE_RX_FORMULA.match(v.strip())
        if (
            m
            and (m.group(2) + m.group(3)).upper() == avg["cell"]
            and (m.group(4) + m.group(5)).upper() == avg["cell"]
        ):
          return {
              "kind": "A0",
              "avg_cell": avg["cell"],
              "r1": avg["r1"],
              "r2": avg["r2"],
              "a0_col": avg["col"],
              "rx_cell": f"{COL_LETTERS[ci]}{ri + 1}",
              "k": float(m.group(1)),
          }

  # 版型二：平均的那一欄每列都是 =k*(1023-D2)/D2 → 平均的是 R(電阻)，這格就是 Rx
  avg_ci = coord_to_index(avg["col"] + "1")[0]
  if avg_ci < GRID_COLS:
    for ri in range(1, GRID_ROWS):
      v = grid_f[ri][avg_ci]
      if isinstance(v, str):
        m = _RE_RX_FORMULA.match(v.strip())
        if (
            m
            and m.group(2).upper() == m.group(4).upper()
            and m.group(3) == m.group(5)
        ):
          return {
              "kind": "R",
              "avg_cell": avg["cell"],
              "r1": avg["r1"],
              "r2": avg["r2"],
              "a0_col": m.group(2).upper(),
              "rx_cell": avg["cell"],
              "k": float(m.group(1)),
          }
    # 版型三：只有 AVERAGE，欄位標題寫 A0（沒有 Rx 公式）
    if str(hdr.get(avg_ci, "")).upper() == "A0":
      return {
          "kind": "A0",
          "avg_cell": avg["cell"],
          "r1": avg["r1"],
          "r2": avg["r2"],
          "a0_col": avg["col"],
          "rx_cell": None,
          "k": None,
      }
  return None


def find_legacy_cells(grid_f):
  """舊版做法：欄優先掃描，各取第一個 =數字* 與 =AVERAGE( 的格子座標。"""
  rx_cell = None
  avg_cell = None
  for ci in range(GRID_COLS):
    for ri in range(GRID_ROWS):
      v = grid_f[ri][ci]
      if isinstance(v, str):
        if rx_cell is None and _RE_RX_LEGACY.match(v):
          rx_cell = f"{COL_LETTERS[ci]}{ri + 1}"
        if avg_cell is None and _RE_AVG_LEGACY.match(v):
          avg_cell = f"{COL_LETTERS[ci]}{ri + 1}"
  return rx_cell, avg_cell


def _pad_row(row, n=GRID_COLS):
  row = list(row)[:n]
  if len(row) < n:
    row.extend([None] * (n - len(row)))
  return row


def _read_grid(ws):
  grid = []
  for row in ws.iter_rows(
      min_row=1, max_row=GRID_ROWS, max_col=GRID_COLS, values_only=True
  ):
    grid.append(_pad_row(row))
  while len(grid) < GRID_ROWS:
    grid.append([None] * GRID_COLS)
  return grid


def build_sheet_info(title, state, grid_f, ws_d):
  """整理一張工作表的資訊（只讀一遍原始數據，之後所有計算都用這份資料）。"""
  hdr = {}
  for i, v in enumerate(grid_f[0]):
    if v is not None and str(v).strip() != "":
      hdr[i] = str(v).strip()
  t_idx = next((i for i, v in hdr.items() if v == "T"), None)
  h_idx = next((i for i, v in hdr.items() if v == "H"), None)
  has_formula = any(
      isinstance(v, str) and v.startswith("=") for row in grid_f for v in row
  )
  layout = None
  rx_cell = None
  avg_cell = None
  if "plxdaq" in title.lower():
    kind = "plxdaq"
  elif title in ("氣溫", "溫度") or (
      t_idx is not None and h_idx is not None and not has_formula
  ):
    kind = "temp"
  else:
    layout = detect_layout(grid_f, hdr)
    rx_cell, avg_cell = find_legacy_cells(grid_f)
    kind = "data" if (layout or rx_cell or avg_cell) else "other"

  a0_idx = None
  if layout:
    idx = coord_to_index(layout["a0_col"] + "1")
    if idx and idx[0] < GRID_COLS:
      a0_idx = idx[0]

  a0_rows = array.array("l")
  a0_vals = array.array("d")
  th = []
  grid_d = []
  for ridx, row in enumerate(
      ws_d.iter_rows(min_row=1, max_col=GRID_COLS, values_only=True), start=1
  ):
    row = _pad_row(row)
    if ridx <= GRID_ROWS:
      grid_d.append(row)
    if ridx >= 2:
      if a0_idx is not None:
        v = row[a0_idx]
        if is_number(v):
          a0_rows.append(ridx)
          a0_vals.append(float(v))
      if t_idx is not None and h_idx is not None:
        tv = row[t_idx]
        if isinstance(tv, str):
          m = _RE_TEMP.search(tv)
          if m:
            hv = parse_hum_percent(row[h_idx])
            if hv is not None:
              th.append((float(m.group(1)), hv))
  while len(grid_d) < GRID_ROWS:
    grid_d.append([None] * GRID_COLS)

  legacy = {
      "rx_cell": rx_cell,
      "rx_val": grid_value(grid_d, rx_cell) if rx_cell else None,
      "a0_cell": avg_cell,
      "a0_val": grid_value(grid_d, avg_cell) if avg_cell else None,
  }
  return {
      "name": title,
      "hidden": str(state).lower() != "visible",
      "kind": kind,
      "layout": layout,
      "legacy": legacy,
      "grid_f": grid_f,
      "grid_d": grid_d,
      "a0_rows": a0_rows,
      "a0_vals": a0_vals,
      "th": th,
  }


def load_file_info(path):
  """讀一個 Excel 檔，整理成資訊 dict（唯讀模式，比較快）。失敗時 error 欄位有說明。"""
  info = {
      "path": path,
      "filename": os.path.basename(path),
      "mtime": None,
      "size": None,
      "error": None,
      "sheets": [],
  }
  try:
    stt = os.stat(path)
    info["mtime"] = stt.st_mtime
    info["size"] = stt.st_size
    raw = []
    wb_f = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
      for ws in wb_f.worksheets:
        raw.append(
            (ws.title, getattr(ws, "sheet_state", "visible"), _read_grid(ws))
        )
    finally:
      wb_f.close()
    wb_d = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
      for title, state, grid_f in raw:
        info["sheets"].append(
            build_sheet_info(title, state, grid_f, wb_d[title])
        )
    finally:
      wb_d.close()
  except Exception as e:
    info["error"] = f"{type(e).__name__}: {e}"
  return info


class InfoCache:
  """記住每個檔案讀過的結果（檔案沒變就不重讀），存成 view_cache.pkl。"""

  def __init__(self, filename="view_cache.pkl"):
    self.file = os.path.join(app_dir(), filename)
    self.data = {}
    self.loaded = False
    self.dirty = False

  @staticmethod
  def _key(path):
    return os.path.normcase(os.path.abspath(path))

  def _ensure_loaded(self):
    if self.loaded:
      return
    self.loaded = True
    try:
      with open(self.file, "rb") as f:
        obj = pickle.load(f)
      if isinstance(obj, dict) and obj.get("v") == CACHE_VERSION:
        self.data = obj.get("data", {})
    except Exception:
      self.data = {}

  def peek(self, path):
    """快取還有效就回傳，否則回傳 None（不會讀檔）。"""
    self._ensure_loaded()
    try:
      stt = os.stat(path)
    except OSError:
      return None
    info = self.data.get(self._key(path))
    if (
        info
        and not info.get("error")
        and info.get("mtime") == stt.st_mtime
        and info.get("size") == stt.st_size
    ):
      info["path"] = path
      return info
    return None

  def get(self, path):
    info = self.peek(path)
    if info is not None:
      return info
    info = load_file_info(path)
    if not info.get("error"):
      self.data[self._key(path)] = info
      self.dirty = True
    return info

  def save(self):
    if not self.dirty:
      return
    try:
      for key in list(self.data.keys()):
        if not os.path.exists(self.data[key].get("path", key)):
          del self.data[key]
      tmp = self.file + ".tmp"
      with open(tmp, "wb") as f:
        pickle.dump(
            {"v": CACHE_VERSION, "data": self.data},
            f,
            protocol=pickle.HIGHEST_PROTOCOL,
        )
      os.replace(tmp, self.file)
      self.dirty = False
    except Exception:
      pass


# ---------- 取值範圍設定（每個「檔案＋工作表」各自設定，存 range_setting.json）----------
RANGE_FILE = "range_setting.json"


def range_key(path, sheet_name):
  return os.path.normcase(os.path.abspath(path)) + "|" + sheet_name


def load_range_settings():
  try:
    with open(os.path.join(app_dir(), RANGE_FILE), "r", encoding="utf-8") as f:
      data = json.load(f)
    out = {}
    if isinstance(data, dict):
      for k, v in data.items():
        if (
            isinstance(v, dict)
            and v.get("mode") == "custom"
            and isinstance(v.get("start"), int)
            and isinstance(v.get("end"), int)
        ):
          out[k] = {"mode": "custom", "start": v["start"], "end": v["end"]}
    return out
  except Exception:
    return {}


def save_range_settings(settings):
  try:
    with open(os.path.join(app_dir(), RANGE_FILE), "w", encoding="utf-8") as f:
      json.dump(settings, f, ensure_ascii=False, indent=2)
    return True
  except Exception:
    return False


def _window_stats(window, k):
  vals = list(window)
  n = len(vals)
  if n == 0:
    return {"n": 0, "mean": None, "std": None, "drift": None, "drift_pct": None}
  mean = sum(vals) / n
  std = statistics.stdev(vals) if n > 1 else 0.0
  drift = None
  drift_pct = None
  if n >= 4:
    half = n // 2
    m1 = sum(vals[:half]) / half
    m2 = sum(vals[n - half:]) / half
    drift = m2 - m1
    if k and m1 > 0 and m2 > 0:
      r1 = k * (1023 - m1) / m1
      r2 = k * (1023 - m2) / m2
      if r1 != 0:
        drift_pct = (r2 - r1) / r1 * 100
  return {"n": n, "mean": mean, "std": std, "drift": drift, "drift_pct": drift_pct}


def get_values(sheet, setting=None):
  """取得一張工作表的 Rx、A0。

  setting=None 或 mode!='custom'：沿用 Excel 公式的範圍（讀 Excel 存好的值）。
  setting={'mode':'custom','start':a,'end':b}：用「第 a~b 筆有效數據」（跳過空列）重算。
  範圍不合法或超出資料筆數時：數值留空，並在 warn 裡說明。
  """
  res = {
      "rx": None, "a0": None, "n": 0, "mean": None, "std": None,
      "drift": None, "drift_pct": None, "win": None, "method": "",
      "warn": [], "note": "",
  }
  if sheet["kind"] not in ("data", "other"):
    res["note"] = "不是數據工作表"
    return res
  layout = sheet.get("layout")
  custom = bool(setting) and setting.get("mode") == "custom"
  grid_d = sheet["grid_d"]

  if layout is None:
    lg = sheet["legacy"]
    res["rx"], e1 = number_or_error(lg["rx_val"])
    res["a0"], e2 = number_or_error(lg["a0_val"])
    res["method"] = "原公式值（找不到可辨識的版型）"
    res["note"] = e1 or e2 or ""
    if custom:
      res["warn"].append("找不到可辨識的公式版型，無法套用自訂範圍，已改用原公式值")
    return res

  rows = sheet["a0_rows"]
  vals = sheet["a0_vals"]
  k = layout["k"]
  if custom:
    a, b = int(setting["start"]), int(setting["end"])
    if a < 1 or b < a:
      res["warn"].append(f"範圍 {a}~{b} 不合法（起點要 ≥ 1，且終點不能小於起點）")
      res["method"] = f"自訂 第{a}~{b}筆"
      return res
    if b > len(vals):
      res["warn"].append(f"範圍 {a}~{b} 超出有效資料筆數（這張工作表只有 {len(vals)} 筆）")
      res["method"] = f"自訂 第{a}~{b}筆"
      return res
    window = vals[a - 1:b]
    res["win"] = (a, b)
    res["method"] = f"自訂 第{a}~{b}筆"
    st = _window_stats(window, k)
    res.update({x: st[x] for x in ("n", "mean", "std", "drift", "drift_pct")})
    res["a0"] = st["mean"]
    if st["mean"] and st["mean"] > 0:
      if not k:
        res["warn"].append("找不到電阻公式的係數 k，無法算 Rx")
      elif layout["kind"] == "R":
        # R 版型的 Excel 是「每列先算電阻再平均」，自訂範圍也照同樣方式算
        rs = [k * (1023 - x) / x for x in window if x > 0]
        res["rx"] = sum(rs) / len(rs) if rs else None
      else:
        # A0 版型的 Excel 是「先平均 A0 再算電阻」
        res["rx"] = k * (1023 - st["mean"]) / st["mean"]
    return res

  lo = bisect.bisect_left(rows, layout["r1"])
  hi = bisect.bisect_right(rows, layout["r2"])
  window = vals[lo:hi]
  res["win"] = (lo + 1, hi) if hi > lo else None
  res["method"] = f"沿用公式 Excel 第{layout['r1']}~{layout['r2']}列"
  st = _window_stats(window, k)
  res.update({x: st[x] for x in ("n", "mean", "std", "drift", "drift_pct")})
  avg_val, e1 = number_or_error(grid_value(grid_d, layout["avg_cell"]))
  rx_val, e2 = (None, None)
  if layout["rx_cell"]:
    rx_val, e2 = number_or_error(grid_value(grid_d, layout["rx_cell"]))
  res["note"] = e1 or e2 or ""
  if layout["kind"] == "A0":
    res["a0"] = avg_val
    res["rx"] = rx_val
  else:
    res["rx"] = avg_val
    res["a0"] = st["mean"]
  return res


def sheet_has_values(sheet):
  v = get_values(sheet, None)
  return v["rx"] is not None or v["a0"] is not None


# ---------- 項目合併（沿用整理模式存的 setting.json）----------
def load_item_mapping():
  """回傳 ({原工作表名: 代表名稱}, [代表名稱的順序])；沒有設定檔就回傳空的。"""
  try:
    with open(os.path.join(app_dir(), "setting.json"), "r", encoding="utf-8") as f:
      data = json.load(f)
  except Exception:
    return {}, []
  raw_to_final = {}
  order = []
  if isinstance(data, dict):
    for final, raws in data.items():
      if isinstance(raws, list):
        order.append(final)
        for r in raws:
          raw_to_final.setdefault(r, final)
  return raw_to_final, order


def natural_key(text):
  return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", str(text))]


def build_points(entries, range_settings, name_of, wanted, source, param):
  """把「檔案×項目」變成圖表用的點。

  entries: [(info, 'YYYY-MM-DD'), ...]
  source: 'rx' | 'a0' | 'regex' | 'cell'；param: 正則或儲存格座標
  回傳 (points, warns, conflicts)
  """
  points = []
  warns = []
  conflicts = []
  pat = re.compile(param) if source == "regex" else None
  for info, date_str in entries:
    if info.get("error"):
      continue
    by_item = {}
    for sh in info["sheets"]:
      if sh["kind"] != "data":
        continue
      final = name_of(sh["name"])
      if final not in wanted:
        continue
      if source in ("rx", "a0"):
        v = get_values(sh, range_settings.get(range_key(info["path"], sh["name"])))
        val = v[source]
        for w in v["warn"]:
          warns.append(f"{info['filename']}／{sh['name']}：{w}")
      elif source == "regex":
        val = None
        done = False
        for ci in range(GRID_COLS):
          for ri in range(GRID_ROWS):
            f = sh["grid_f"][ri][ci]
            if isinstance(f, str) and pat.match(f):
              val, _ = number_or_error(sh["grid_d"][ri][ci])
              done = True
              break
          if done:
            break
      else:
        val, _ = number_or_error(grid_value(sh["grid_d"], param))
      by_item.setdefault(final, []).append((sh["name"], val))
    for final, lst in by_item.items():
      valid = [(n, v) for n, v in lst if v is not None]
      chosen = valid[0] if valid else (lst[0][0], None)
      if len(valid) > 1:
        conflicts.append(
            f"{date_str}／{info['filename']}／{final}：多個工作表都有數值"
            f"（{', '.join(n for n, _ in valid)}），採用 {chosen[0]}"
        )
      points.append({
          "date": date_str, "file": info["filename"], "path": info["path"],
          "item": final, "sheet": chosen[0], "value": chosen[1],
      })
  return points, warns, conflicts

THEME_FILE = "theme_setting.json"
PALETTES = {
    "light": {
        "dark": False, "fg": "#1c1c1c", "text2": "#444444", "muted": "#666666",
        "faint": "#999999", "card": "#ffffff", "card_border": "#d0d0d0",
        "accent": "#0067c0", "sel_bg": "#0067c0", "sel_fg": "#ffffff",
        "has": "#e8f1fb", "hover": "#eaf2fb", "code_bg": "#eceff1",
        "block_bg": "#f3f4f6", "link": "#0067c0", "warn": "#c0392b",
        "ai": "#26a269", "popup_border": "#b8b8b8", "plot_bg": "#ffffff",
        "grid": "#eceff3", "axis": "#c9ced6", "legend_bg": "rgba(255,255,255,0.8)",
    },
    "tech": {
        "dark": True, "fg": "#e6edf7", "text2": "#b6c4dc", "muted": "#8b9bb4",
        "faint": "#6c7a92", "card": "#10182a", "card_border": "#2b3f66",
        "accent": "#57c8ff", "sel_bg": "#1f4f8f", "sel_fg": "#ffffff",
        "has": "#14284a", "hover": "#1b2c4d", "code_bg": "#1b2740",
        "block_bg": "#0c1424", "link": "#57c8ff", "warn": "#ff6b6b",
        "ai": "#3ddc97", "popup_border": "#57c8ff", "plot_bg": "#0d1424",
        "grid": "#22304d", "axis": "#3a4d73", "legend_bg": "rgba(13,20,36,0.8)",
    },
}
PAL = dict(PALETTES["light"])


def set_palette(name):
  """切換色票（PAL 是共用的字典，原地更新，所有模組都會看到新顏色）。"""
  PAL.clear()
  PAL.update(PALETTES.get(name, PALETTES["light"]))


set_palette("light")


UI_SCALE_CHOICES = (0.8, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0)


def _load_prefs():
  try:
    with open(os.path.join(app_dir(), THEME_FILE), "r", encoding="utf-8") as f:
      d = json.load(f)
    return d if isinstance(d, dict) else {}
  except Exception:
    return {}


def _save_prefs(**kw):
  d = _load_prefs()
  d.update(kw)
  try:
    with open(os.path.join(app_dir(), THEME_FILE), "w", encoding="utf-8") as f:
      json.dump(d, f, ensure_ascii=False)
  except Exception:
    pass


def load_theme_setting():
  t = _load_prefs().get("theme")
  return t if t in PALETTES else "light"


def save_theme_setting(name):
  _save_prefs(theme=name)


def load_ui_scale():
  v = _load_prefs().get("ui_scale")
  return float(v) if isinstance(v, (int, float)) and 0.5 <= v <= 3 else 1.0


def save_ui_scale(v):
  _save_prefs(ui_scale=v)


def listbox_colors():
  return dict(
      bg=PAL["card"], fg=PAL["fg"], selectbackground=PAL["sel_bg"],
      selectforeground=PAL["sel_fg"], relief="flat", borderwidth=0,
      highlightthickness=1, highlightbackground=PAL["card_border"],
      highlightcolor=PAL["accent"],
  )
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


def _disp_width(t):
  return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in t)
