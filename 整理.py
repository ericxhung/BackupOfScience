import ctypes
import json
import os
import re
import string
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont
import openpyxl
from openpyxl.styles import Border, Font, Side
import pandas as pd
import array
import bisect
import calendar
import math
import pickle
import statistics
import subprocess
import tempfile
import time
import unicodedata
import webbrowser
import csv
import io
import queue
import threading
import urllib.error
import urllib.request
from collections import OrderedDict, defaultdict

# 【UI 美化】嘗試載入 sv-ttk（Windows 11 風格主題）。
# 沒安裝也能執行，只是會退回原本的 tkinter 外觀。
try:
  import sv_ttk

  HAS_SV_TTK = True
except ImportError:
  HAS_SV_TTK = False

# 【檢視模式】互動圖表用 plotly（在瀏覽器開啟）；沒安裝就只能用視窗內的 matplotlib 圖。
try:
  import plotly.graph_objects as go

  HAS_PLOTLY = True
except Exception:
  HAS_PLOTLY = False

# 【檢視模式】圖表用 matplotlib；沒安裝也能執行，只是圖表分頁會提示要安裝。
try:
  import matplotlib
  import matplotlib.dates as mdates
  from matplotlib import font_manager as mpl_font_manager
  from matplotlib.figure import Figure

  HAS_MPL = True
except Exception:
  HAS_MPL = False

try:
  from matplotlib.backends.backend_tkagg import (
      FigureCanvasTkAgg,
      NavigationToolbar2Tk,
  )

  HAS_MPL_TK = True
except Exception:
  HAS_MPL_TK = False


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
    self.file = os.path.join(os.getcwd(), filename)
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
    with open(os.path.join(os.getcwd(), RANGE_FILE), "r", encoding="utf-8") as f:
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
    with open(os.path.join(os.getcwd(), RANGE_FILE), "w", encoding="utf-8") as f:
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
    with open(os.path.join(os.getcwd(), "setting.json"), "r", encoding="utf-8") as f:
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


# ---------- 圖表 ----------
LAYOUT_KEYS = ["date_item", "item_date", "num_series"]
LAYOUT_LABELS = [
    "x：日期　y：數值　數列：項目",
    "x：項目　y：數值　數列：日期",
    "x：項目中的數字（如30、40…）　數列：文字部分＋日期",
]
CHART_MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]


def setup_mpl_fonts():
  """matplotlib 預設字型沒有中文，會變成方塊字，所以挑系統上有的中文字型。"""
  if not HAS_MPL:
    return
  try:
    names = {f.name for f in mpl_font_manager.fontManager.ttflist}
    prefer = [
        "Microsoft JhengHei", "Microsoft JhengHei UI", "Microsoft YaHei",
        "PingFang TC", "Noto Sans CJK TC", "Noto Sans TC", "Heiti TC",
        "Arial Unicode MS",
    ]
    chosen = [n for n in prefer if n in names]
    matplotlib.rcParams["font.sans-serif"] = chosen + list(
        matplotlib.rcParams["font.sans-serif"]
    )
    matplotlib.rcParams["axes.unicode_minus"] = False
  except Exception:
    pass


def item_number_parts(name):
  """'升溫30' -> ('升溫', 30.0)；沒有數字回傳 None。"""
  m = re.match(r"^(\D*?)(\d+(?:\.\d+)?)(\D*)$", str(name).strip())
  if not m:
    return None
  text = (m.group(1) + m.group(3)).strip() or "數值"
  return text, float(m.group(2))


def _prepare_series(points, layout, item_order):
  skipped = 0
  files_per_date = defaultdict(set)
  for p in points:
    files_per_date[p["date"]].add(p["file"])

  def date_label(p):
    if len(files_per_date[p["date"]]) > 1:
      return f"{p['date']} {os.path.splitext(p['file'])[0]}"
    return p["date"]

  order_index = {n: i for i, n in enumerate(item_order or [])}

  def item_sort(n):
    return (order_index.get(n, 10**6), natural_key(n))

  series = OrderedDict()
  if layout == "date_item":
    for p in sorted(points, key=lambda p: (item_sort(p["item"]), p["date"])):
      series.setdefault(p["item"], []).append((p["date"], p["value"]))
    cats = sorted({p["date"] for p in points})
  elif layout == "item_date":
    for p in sorted(points, key=lambda p: (p["date"], p["file"])):
      series.setdefault(date_label(p), []).append((p["item"], p["value"]))
    cats = sorted({p["item"] for p in points}, key=item_sort)
  else:
    nums = set()
    for p in sorted(points, key=lambda p: (p["date"], p["file"])):
      parts = item_number_parts(p["item"])
      if parts is None:
        skipped += 1
        continue
      text, num = parts
      series.setdefault(f"{text} {date_label(p)}", []).append((num, p["value"]))
      nums.add(num)
    cats = sorted(nums)
  return series, cats, skipped


def _cat_label(c, layout, cats):
  if layout == "date_item":
    years = {x[:4] for x in cats}
    return c[5:] if len(years) == 1 else c
  if layout == "num_series":
    return f"{c:g}"
  return str(c)


def _mpl_theme(fig, ax):
  fg = PAL["fg"]
  fig.set_facecolor(PAL["plot_bg"])
  ax.set_facecolor(PAL["plot_bg"])
  ax.tick_params(colors=fg)
  ax.xaxis.label.set_color(fg)
  ax.yaxis.label.set_color(fg)
  for sp in ax.spines.values():
    sp.set_color(PAL["axis"])
  ax.grid(True, color=PAL["grid"], alpha=0.8, linewidth=0.8)


def draw_chart(fig, points, layout, ctype, connect=True, y_label="數值", item_order=None):
  """在 fig 上畫圖。ctype: 'scatter' | 'line' | 'bar'。回傳統計 dict。"""
  fig.clear()
  ax = fig.add_subplot(111)
  series, cats, skipped = _prepare_series(points, layout, item_order)
  total = sum(len(v) for v in series.values())
  missing = sum(1 for v in series.values() for _, y in v if y is None)
  stats = {"points": total, "missing": missing, "skipped": skipped}
  if total - missing <= 0:
    ax.text(
        0.5, 0.5, "沒有可以畫的數據", ha="center", va="center",
        transform=ax.transAxes, fontsize=14, color="#888888",
    )
    ax.set_axis_off()
    fig.set_facecolor(PAL["plot_bg"])
    return stats

  labels = list(series.keys())
  S = len(labels)
  try:
    ax.set_prop_cycle(
        color=matplotlib.colormaps["tab10" if S <= 10 else "tab20"].colors
    )
  except Exception:
    pass
  cat_pos = {c: i for i, c in enumerate(cats)}
  cat_text = [_cat_label(c, layout, cats) for c in cats]

  if ctype == "bar":
    w = 0.8 / S
    for si, label in enumerate(labels):
      xs, ys = [], []
      for xk, y in series[label]:
        if y is None:
          continue
        xs.append(cat_pos[xk] + (si - (S - 1) / 2) * w)
        ys.append(y)
      ax.bar(xs, ys, width=w * 0.92, label=label)
    ax.set_xticks(range(len(cats)))
    ax.set_xticklabels(
        cat_text, rotation=30 if len(cats) > 6 else 0,
        ha="right" if len(cats) > 6 else "center",
    )
  else:
    for si, label in enumerate(labels):
      pts = []
      for xk, y in series[label]:
        if layout == "date_item":
          x = mdates.date2num(datetime.strptime(xk, "%Y-%m-%d"))
        elif layout == "item_date":
          x = cat_pos[xk]
        else:
          x = xk
        pts.append((x, float("nan") if y is None else y))
      pts.sort(key=lambda t: t[0])
      marker = CHART_MARKERS[si % len(CHART_MARKERS)]
      if ctype == "line":
        ax.plot(
            [t[0] for t in pts], [t[1] for t in pts], marker=marker,
            linestyle="-" if connect else "none", markersize=6, label=label,
        )
      else:
        good = [t for t in pts if not math.isnan(t[1])]
        ax.scatter(
            [t[0] for t in good], [t[1] for t in good], marker=marker, s=40,
            label=label,
        )
    if layout == "date_item":
      ticks = [mdates.date2num(datetime.strptime(c, "%Y-%m-%d")) for c in cats]
      if len(ticks) <= 14:
        ax.set_xticks(ticks)
        ax.set_xticklabels(cat_text, rotation=30 if len(ticks) > 6 else 0)
      else:
        loc = mdates.AutoDateLocator()
        ax.xaxis.set_major_locator(loc)
        ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(loc))
    elif layout == "item_date":
      ax.set_xticks(range(len(cats)))
      ax.set_xticklabels(
          cat_text, rotation=30 if len(cats) > 6 else 0,
          ha="right" if len(cats) > 6 else "center",
      )
    ax.margins(x=0.06)

  ax.set_xlabel(
      {"date_item": "日期", "item_date": "項目", "num_series": "項目中的數字"}[layout]
  )
  ax.set_ylabel(y_label)
  ax.grid(True, alpha=0.25, linewidth=0.8)
  # 圖例放在圖的下方（橫向排列），圖表本體可以用滿整個寬度
  ncol = max(1, min(S, 5))
  rows = math.ceil(S / ncol)
  fig.legend(
      *ax.get_legend_handles_labels(), loc="lower center", ncol=ncol,
      fontsize=9 if S <= 15 else 8, frameon=False, labelcolor=PAL["fg"],
  )
  for side in ("top", "right"):
    ax.spines[side].set_visible(False)
  ax.set_axisbelow(True)
  fig.subplots_adjust(
      left=0.08, right=0.98, top=0.95, bottom=min(0.5, 0.20 + 0.045 * rows)
  )
  _mpl_theme(fig, ax)
  return stats


PLOTLY_COLORS = [
    "#4C78A8", "#F58518", "#54A24B", "#E45756", "#72B7B2", "#B79A20",
    "#B279A2", "#FF9DA6", "#9D755D", "#7F7F7F", "#A0CBE8", "#FFBE7D",
    "#8CD17D", "#FF9D9A", "#86BCB6", "#F1CE63", "#D4A6C8", "#FABFD2",
    "#D7B5A6", "#BAB0AC",
]
PLOTLY_SYMBOLS = [
    "circle", "square", "diamond", "triangle-up", "triangle-down", "cross",
    "x", "star",
]


def build_plotly_html(points, layout, ctype, connect, y_label, item_order, path):
  """用 plotly 畫互動圖，存成可離線開啟的 HTML。回傳統計 dict。"""
  series, cats, skipped = _prepare_series(points, layout, item_order)
  total = sum(len(v) for v in series.values())
  missing = sum(1 for v in series.values() for _, y in v if y is None)
  stats = {"points": total, "missing": missing, "skipped": skipped}
  if total - missing <= 0:
    return stats
  cat_text = [_cat_label(c, layout, cats) for c in cats]
  text_of = dict(zip(cats, cat_text))
  fig = go.Figure()
  for si, (label, pts) in enumerate(series.items()):
    color = PLOTLY_COLORS[si % len(PLOTLY_COLORS)]
    if ctype == "bar":
      good = [(text_of[xk], y) for xk, y in pts if y is not None]
      fig.add_trace(go.Bar(
          x=[g[0] for g in good], y=[g[1] for g in good], name=label,
          marker=dict(color=color),
          hovertemplate="<b>%{fullData.name}</b><br>%{x}<br>" + y_label + "：%{y:.4f}<extra></extra>",
      ))
      continue
    if layout == "num_series":
      pts = sorted(pts, key=lambda t: t[0])
      xs = [xk for xk, _ in pts]
    elif layout == "date_item":
      xs = [xk for xk, _ in pts]
    else:
      xs = [text_of[xk] for xk, _ in pts]
    ys = [y for _, y in pts]
    mode = "lines+markers" if (ctype == "line" and connect) else "markers"
    fig.add_trace(go.Scatter(
        x=xs, y=ys, name=label, mode=mode, connectgaps=False,
        line=dict(color=color, width=2.5),
        marker=dict(
            color=color, size=9, symbol=PLOTLY_SYMBOLS[si % len(PLOTLY_SYMBOLS)],
            line=dict(color=PAL["plot_bg"], width=1),
        ),
        hovertemplate="<b>%{fullData.name}</b><br>%{x}<br>" + y_label + "：%{y:.4f}<extra></extra>",
    ))
  x_title = {"date_item": "日期", "item_date": "項目", "num_series": "項目中的數字"}[layout]
  fig.update_layout(
      template="plotly_dark" if PAL["dark"] else "plotly_white",
      paper_bgcolor=PAL["plot_bg"],
      font=dict(
          family="Microsoft JhengHei, PingFang TC, Noto Sans CJK TC, sans-serif",
          size=14, color=PAL["fg"],
      ),
      xaxis=dict(title=x_title, gridcolor=PAL["grid"], showline=True, linecolor=PAL["axis"], zeroline=False),
      yaxis=dict(title=y_label, gridcolor=PAL["grid"], showline=True, linecolor=PAL["axis"], zeroline=False),
      hovermode="closest",
      barmode="group",
      legend=dict(title=dict(text="點圖例開關（雙擊只看單一項）"), bgcolor=PAL["legend_bg"]),
      margin=dict(l=70, r=30, t=40, b=70),
      plot_bgcolor=PAL["plot_bg"],
  )
  if ctype == "bar":
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=cat_text)
  elif layout == "date_item":
    fig.update_xaxes(type="date")
    if len(cats) <= 14:
      fig.update_xaxes(tickmode="array", tickvals=cats, ticktext=cat_text)
  elif layout == "item_date":
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=cat_text)
  fig.write_html(
      path, include_plotlyjs=True, full_html=True,
      default_width="100%", default_height="97vh",
      config={"responsive": True, "displaylogo": False,
              "toImageButtonOptions": {"format": "png", "scale": 2}},
  )
  return stats


def open_file(path):
  """用系統預設程式（Excel）開啟檔案。"""
  try:
    if sys.platform == "win32":
      os.startfile(path)
    elif sys.platform == "darwin":
      subprocess.Popen(["open", path])
    else:
      subprocess.Popen(["xdg-open", path])
    return True
  except Exception as e:
    messagebox.showerror("開啟失敗", f"無法開啟檔案：\n{path}\n\n{e}")
    return False


class ProgressWindow(tk.Toplevel):
  """讀取檔案時的進度視窗（可按取消）。"""

  def __init__(self, app, parent, title, total):
    super().__init__(parent)
    self.title(title)
    self.cancelled = False
    self.resizable(False, False)
    bg = getattr(app, "theme_bg", None)
    if bg:
      self.configure(bg=bg)
    body = ttk.Frame(self, padding=16)
    body.pack(fill="both", expand=True)
    self.lbl = ttk.Label(body, text="準備中…", width=56)
    self.lbl.pack(anchor="w")
    self.bar = ttk.Progressbar(body, maximum=max(total, 1), length=440)
    self.bar.pack(pady=10)
    ttk.Button(body, text="取消", command=self._cancel).pack()
    self.protocol("WM_DELETE_WINDOW", self._cancel)
    self.update_idletasks()
    top = parent.winfo_toplevel()
    w, h = self.winfo_reqwidth(), self.winfo_reqheight()
    x = top.winfo_rootx() + (top.winfo_width() - w) // 2
    y = top.winfo_rooty() + (top.winfo_height() - h) // 2
    self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    self.transient(top)
    self.grab_set()
    self.update()

  def set_text(self, text):
    if len(text) > 56:
      text = text[:53] + "…"
    self.lbl.config(text=text)
    self.update()

  def step(self):
    self.bar["value"] = self.bar["value"] + 1
    self.update()

  def _cancel(self):
    self.cancelled = True

  def close(self):
    try:
      self.grab_release()
      self.destroy()
    except tk.TclError:
      pass


class RangeDialog(tk.Toplevel):
  """每個「檔案＋工作表」各自設定取值範圍；可多選一起編輯，並即時預覽結果。"""

  COLS = (
      ("date", "日期", 90),
      ("file", "檔案", 170),
      ("sheet", "工作表", 100),
      ("range", "目前範圍", 230),
      ("n", "有效筆數", 70),
      ("mean", "平均 A0", 80),
      ("std", "標準差", 65),
      ("drift", "前後半差", 75),
      ("rx", "Rx", 75),
      ("status", "狀態", 220),
  )

  def __init__(self, app, infos):
    super().__init__(app)
    self.app = app
    self.title("取值範圍設定（每個工作表可各自設定）")
    self.saved = False
    self.settings = {k: dict(v) for k, v in app.range_settings.items()}
    self.entries = []
    for info in infos:
      if info.get("error"):
        continue
      for sh in info["sheets"]:
        if sh["kind"] == "data":
          self.entries.append((info, sh, range_key(info["path"], sh["name"])))
    self.entries.sort(
        key=lambda e: (app.date_text_of(e[0]) or "9999", e[0]["filename"])
    )
    app.center_dialog(self, 1180, 660)
    self._build()
    self._fill()
    self.grab_set()

  def _build(self):
    body = ttk.Frame(self, padding=12)
    body.pack(fill="both", expand=True)
    ttk.Label(
        body,
        text=(
            "「沿用公式」＝照 Excel 裡 =AVERAGE(...) 的範圍；「自訂」＝改用第 N~M 筆"
            "有效數據（空白列不算）。\n可按住 Ctrl／Shift 選多列一起套用。"
            "紅字代表範圍有問題，該列數值會留空。「前後半差」是範圍內"
            "後半段平均減前半段平均（A0 單位），越小代表越穩定。"
        ),
        style="Hint.TLabel",
        justify="left",
    ).pack(anchor="w")

    filt = ttk.Frame(body)
    filt.pack(fill="x", pady=6)
    ttk.Label(filt, text="篩選（檔名或工作表名）：").pack(side="left")
    self.var_filter = tk.StringVar()
    ttk.Entry(filt, textvariable=self.var_filter, width=24).pack(
        side="left", padx=6
    )
    ttk.Button(filt, text="全選目前列表", command=self._select_all).pack(
        side="left", padx=6
    )

    box = ttk.Frame(body)
    box.pack(fill="both", expand=True)
    self.tree = ttk.Treeview(
        box,
        columns=[c[0] for c in self.COLS],
        show="headings",
        selectmode="extended",
    )
    for cid, text, width in self.COLS:
      self.tree.heading(cid, text=text)
      self.tree.column(cid, width=self.app.px(width), anchor="w")
    self.tree.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
    sb.pack(side="right", fill="y")
    self.tree.configure(yscrollcommand=sb.set)
    self.tree.tag_configure("warn", foreground=PAL["warn"])
    self.tree.tag_configure("empty", foreground=PAL["faint"])
    self.tree.bind("<<TreeviewSelect>>", lambda e: self._on_select())

    ed = ttk.LabelFrame(body, text="編輯選取的列", padding=10)
    ed.pack(fill="x", pady=8)
    self.var_mode = tk.StringVar(value="formula")
    ttk.Radiobutton(
        ed, text="沿用 Excel 公式的範圍", value="formula", variable=self.var_mode
    ).grid(row=0, column=0, sticky="w", padx=4)
    ttk.Radiobutton(
        ed, text="自訂：第", value="custom", variable=self.var_mode
    ).grid(row=0, column=1, sticky="w", padx=(20, 2))
    self.ent_start = ttk.Entry(ed, width=8)
    self.ent_start.grid(row=0, column=2, padx=2)
    ttk.Label(ed, text="到第").grid(row=0, column=3, padx=2)
    self.ent_end = ttk.Entry(ed, width=8)
    self.ent_end.grid(row=0, column=4, padx=2)
    ttk.Label(ed, text="筆有效數據（空白列不算）").grid(row=0, column=5, padx=4)
    ttk.Button(
        ed, text="套用到選取列", style="Accent.TButton", command=self._apply
    ).grid(row=0, column=6, padx=(20, 4))
    ttk.Button(ed, text="清除選取列的設定", command=self._clear).grid(
        row=0, column=7, padx=4
    )

    bottom = ttk.Frame(body)
    bottom.pack(fill="x")
    ttk.Button(bottom, text="取消", command=self.destroy).pack(
        side="right", padx=4
    )
    ttk.Button(
        bottom, text="儲存並關閉", style="Accent.TButton", command=self._save
    ).pack(side="right", padx=4)

    # 篩選框要等表格建好之後才綁定
    self.var_filter.trace_add("write", lambda *a: self._fill())

  def _fill(self):
    keep = set(self.tree.selection())
    self.tree.delete(*self.tree.get_children())
    q = self.var_filter.get().strip().lower()

    def fmt(x, nd=2):
      return "" if x is None else f"{x:.{nd}f}"

    for i, (info, sh, key) in enumerate(self.entries):
      if q and q not in (info["filename"] + " " + sh["name"]).lower():
        continue
      st = self.settings.get(key)
      v = get_values(sh, st)
      custom = bool(st) and st.get("mode") == "custom"
      rng = v["method"]
      if not custom and v["win"]:
        rng += f"＝第{v['win'][0]}~{v['win'][1]}筆"
      no_value = v["rx"] is None and v["a0"] is None
      status = "；".join(v["warn"]) or v["note"]
      if not status and no_value:
        status = "沒有數值"
      if v["warn"]:
        tags = ("warn",)
      elif no_value:
        tags = ("empty",)
      else:
        tags = ()
      self.tree.insert(
          "",
          "end",
          iid=str(i),
          values=(
              self.app.date_text_of(info),
              info["filename"],
              sh["name"],
              rng,
              v["n"],
              fmt(v["mean"]),
              fmt(v["std"]),
              fmt(v["drift"]),
              fmt(v["rx"], 3),
              status,
          ),
          tags=tags,
      )
    still = [i for i in keep if self.tree.exists(i)]
    if still:
      self.tree.selection_set(still)

  def _select_all(self):
    self.tree.selection_set(self.tree.get_children())

  def _on_select(self):
    sel = self.tree.selection()
    if len(sel) != 1:
      return
    info, sh, key = self.entries[int(sel[0])]
    st = self.settings.get(key)
    custom = bool(st) and st.get("mode") == "custom"
    self.var_mode.set("custom" if custom else "formula")
    if custom:
      a, b = st["start"], st["end"]
    else:
      win = get_values(sh, None)["win"]
      if not win:
        return
      a, b = win
    self.ent_start.delete(0, tk.END)
    self.ent_start.insert(0, str(a))
    self.ent_end.delete(0, tk.END)
    self.ent_end.insert(0, str(b))

  def _apply(self):
    sel = self.tree.selection()
    if not sel:
      messagebox.showwarning("提示", "請先選取要編輯的列！", parent=self)
      return
    mode = self.var_mode.get()
    a = b = None
    if mode == "custom":
      try:
        a = int(self.ent_start.get().strip())
        b = int(self.ent_end.get().strip())
      except ValueError:
        messagebox.showerror(
            "格式錯誤", "起點與終點請輸入整數！", parent=self
        )
        return
      if a < 1 or b < a:
        messagebox.showerror(
            "範圍不合法", "起點要 ≥ 1，而且終點不能小於起點！", parent=self
        )
        return
    warns = []
    for iid in sel:
      info, sh, key = self.entries[int(iid)]
      if mode == "custom":
        self.settings[key] = {"mode": "custom", "start": a, "end": b}
      else:
        self.settings.pop(key, None)
      for w in get_values(sh, self.settings.get(key))["warn"]:
        warns.append(f"{info['filename']}／{sh['name']}：{w}")
    self._fill()
    if warns:
      more = f"\n...以及其他 {len(warns) - 12} 筆" if len(warns) > 12 else ""
      messagebox.showwarning(
          "取值範圍警告",
          "以下工作表的範圍有問題，數值會留空：\n\n" + "\n".join(warns[:12]) + more,
          parent=self,
      )

  def _clear(self):
    sel = self.tree.selection()
    if not sel:
      messagebox.showwarning("提示", "請先選取要清除設定的列！", parent=self)
      return
    for iid in sel:
      self.settings.pop(self.entries[int(iid)][2], None)
    self._fill()

  def _save(self):
    self.app.range_settings = self.settings
    if not save_range_settings(self.settings):
      messagebox.showwarning(
          "儲存失敗",
          "無法寫入 range_setting.json，設定只在這次執行期間有效。",
          parent=self,
      )
    self.saved = True
    self.destroy()


# ==================== 檢視模式（日曆／圖表／概況）====================
VIEW_FILE = "view_setting.json"

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
CAL_BORDER = CAL_HAS = CAL_TODAY = CAL_LINK = CAL_MUTED = ""


def set_palette(name):
  """切換色票（同時更新日曆用的舊式常數）。"""
  global CAL_BORDER, CAL_HAS, CAL_TODAY, CAL_LINK, CAL_MUTED
  PAL.clear()
  PAL.update(PALETTES.get(name, PALETTES["light"]))
  CAL_BORDER = PAL["card_border"]
  CAL_HAS = PAL["has"]
  CAL_TODAY = PAL["accent"]
  CAL_LINK = PAL["link"]
  CAL_MUTED = PAL["muted"]


set_palette("light")


UI_SCALE_CHOICES = (0.8, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0)


def _load_prefs():
  try:
    with open(os.path.join(os.getcwd(), THEME_FILE), "r", encoding="utf-8") as f:
      d = json.load(f)
    return d if isinstance(d, dict) else {}
  except Exception:
    return {}


def _save_prefs(**kw):
  d = _load_prefs()
  d.update(kw)
  try:
    with open(os.path.join(os.getcwd(), THEME_FILE), "w", encoding="utf-8") as f:
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

SRC_LABELS = [
    "Rx 電阻（自動判斷公式）",
    "A0 平均（自動判斷公式）",
    "自訂正則（比對公式文字）",
    "指定儲存格",
]
SRC_KEYS = ["rx", "a0", "regex", "cell"]
SRC_DEFAULT_PARAM = {"rx": "", "a0": "", "regex": r"^=\d+\*", "cell": "I3"}
SRC_HINT = {
    "rx": "依公式意思找 Rx，範圍沿用「⚙ 取值範圍設定」。",
    "a0": "依公式意思找 A0 平均，範圍沿用「⚙ 取值範圍設定」。",
    "regex": "比對公式文字，取第一個符合的格子（A~P 欄、1~30 列）。",
    "cell": "每張工作表都讀這一格的值（例：I3、J2）。",
}
Y_LABELS = {
    "rx": "Rx 電阻 (Ω)",
    "a0": "A0 平均 (ADC)",
    "regex": "數值（正則取值）",
    "cell": "數值（指定儲存格）",
}
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


def load_view_setting():
  try:
    with open(os.path.join(os.getcwd(), VIEW_FILE), "r", encoding="utf-8") as f:
      data = json.load(f)
    if isinstance(data, dict):
      return data
  except Exception:
    pass
  return {}


def save_view_setting(data):
  try:
    with open(os.path.join(os.getcwd(), VIEW_FILE), "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=2)
  except Exception:
    pass


def shorten(text, n):
  return text if len(text) <= n else text[: n - 1] + "…"


class CheckList(ttk.Frame):
  """打勾方塊清單（可捲動、全選／全不選）。"""

  def __init__(self, parent, title, height_px, fam, app=None):
    super().__init__(parent)
    self.vars = []
    self.fam = fam
    head = ttk.Frame(self)
    head.pack(fill="x")
    ttk.Label(head, text=title).pack(side="left")
    ttk.Button(head, text="全選", width=5, command=lambda: self.set_all(True)).pack(
        side="right"
    )
    ttk.Button(
        head, text="全不選", width=6, command=lambda: self.set_all(False)
    ).pack(side="right", padx=4)
    body = ttk.Frame(self)
    body.pack(fill="both", expand=True, pady=(2, 0))
    self.canvas = tk.Canvas(
        body, height=height_px, highlightthickness=1,
        highlightbackground=PAL["card_border"], bg=PAL["card"],
    )
    sb = ttk.Scrollbar(body, orient="vertical", command=self.canvas.yview)
    self.canvas.configure(yscrollcommand=sb.set)
    self.canvas.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    self.inner = tk.Frame(self.canvas, bg=PAL["card"])
    self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
    self.inner.bind(
        "<Configure>",
        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
    )
    self.canvas.bind(
        "<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width)
    )

  def _on_wheel(self, event):
    self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

  def set_items(self, labels, checked=True):
    for w in self.inner.winfo_children():
      w.destroy()
    self.vars = []
    for text in labels:
      v = tk.BooleanVar(value=checked)
      cb = tk.Checkbutton(
          self.inner, text=text, variable=v, anchor="w", bg=PAL["card"],
          fg=PAL["fg"], activebackground=PAL["hover"], activeforeground=PAL["fg"],
          selectcolor=PAL["card"], font=(self.fam, 10), bd=0,
          highlightthickness=0,
      )
      cb.pack(fill="x", anchor="w")
      self.vars.append(v)
    self.canvas.yview_moveto(0)

  def set_all(self, value):
    for v in self.vars:
      v.set(value)

  def checked(self):
    return [i for i, v in enumerate(self.vars) if v.get()]


GEMINI_DEFAULT_MODEL = "gemini-3.5-flash-lite"
GEMINI_SETTING_FILE = "gemini_setting.json"
AI_SYSTEM = (
    "你是協助分析實驗／量測 Excel 數據的助手。使用者會附上幾個 Excel 檔轉成的 CSV，"
    "每段以「### 檔案：… ／ 工作表：…」開頭。請用繁體中文回答，"
    "引用數字時要說明來自哪個檔案與工作表；資料裡沒有的內容不要猜，直接說沒有。"
)


def load_gemini_setting():
  try:
    with open(GEMINI_SETTING_FILE, "r", encoding="utf-8") as f:
      d = json.load(f)
    return d if isinstance(d, dict) else {}
  except (OSError, ValueError):
    return {}


def save_gemini_setting(d):
  try:
    with open(GEMINI_SETTING_FILE, "w", encoding="utf-8") as f:
      json.dump(d, f, ensure_ascii=False, indent=2)
    return True
  except OSError:
    return False


def _csv_cell(v):
  if v is None:
    return ""
  if isinstance(v, float):
    return repr(round(v, 10)) if v == v else ""
  if isinstance(v, datetime):
    return v.strftime("%Y-%m-%d %H:%M:%S")
  return str(v)


def build_ai_csv(path, row_limit=None):
  """把一個 Excel 的每個工作表轉成 CSV 文字（用 Excel 存的計算結果）。回傳 (文字, 說明列表)。"""
  notes = []
  out = []
  try:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
  except Exception as e:
    return "", [f"{os.path.basename(path)} 讀取失敗：{e}"]
  try:
    for ws in wb.worksheets:
      buf = io.StringIO()
      w = csv.writer(buf, lineterminator="\n")
      n = 0
      cut = False
      for row in ws.iter_rows(values_only=True):
        vals = [_csv_cell(v) for v in row]
        while vals and vals[-1] == "":
          vals.pop()
        if not vals:
          continue
        if row_limit and n >= row_limit:
          cut = True
          break
        w.writerow(vals)
        n += 1
      if n == 0:
        continue
      head = f"### 檔案：{os.path.basename(path)} ／ 工作表：{ws.title}"
      if cut:
        head += f"（只取前 {row_limit} 列）"
        notes.append(f"{os.path.basename(path)}／{ws.title} 只取前 {row_limit} 列")
      out.append(head + "\n" + buf.getvalue())
  finally:
    wb.close()
  return "\n".join(out), notes


OPENAI_DEFAULT_URL = "https://api.openai.com/v1"
OPENAI_DEFAULT_MODEL = "gpt-4o-mini"
PROVIDERS = ("Gemini", "OpenAI 格式")


def openai_generate(api_key, base_url, model, system_text, history, timeout=180):
  """呼叫 OpenAI 格式（/chat/completions）的 API，也適用其他相容服務。"""
  url = (base_url.strip() or OPENAI_DEFAULT_URL).rstrip("/") + "/chat/completions"
  msgs = [{"role": "system", "content": system_text}]
  for r, t in history:
    msgs.append({"role": "assistant" if r == "model" else "user", "content": t})
  req = urllib.request.Request(
      url,
      data=json.dumps({"model": model, "messages": msgs}).encode("utf-8"),
      headers={
          "Content-Type": "application/json",
          "Authorization": "Bearer " + api_key,
      },
      method="POST",
  )
  try:
    with urllib.request.urlopen(req, timeout=timeout) as resp:
      data = json.loads(resp.read().decode("utf-8"))
  except urllib.error.HTTPError as e:
    try:
      err = json.loads(e.read().decode("utf-8")).get("error", {})
      detail = err.get("message", "") if isinstance(err, dict) else str(err)
    except Exception:
      detail = ""
    raise RuntimeError(f"HTTP {e.code}：{detail or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  choices = data.get("choices") or []
  if not choices:
    raise RuntimeError(f"沒有收到回答：{str(data)[:200]}")
  text = (choices[0].get("message") or {}).get("content") or ""
  if not text:
    text = f"（沒有文字內容，結束原因：{choices[0].get('finish_reason')}）"
  u = data.get("usage") or {}
  return text, {"totalTokenCount": u.get("total_tokens")}


def _sse_events(resp, stop):
  """逐行讀 Server-Sent Events，回傳每個 data: 後面的 JSON 字串。"""
  for raw in resp:
    if stop is not None and stop.is_set():
      return
    line = raw.decode("utf-8", "replace").strip()
    if line.startswith("data:"):
      payload = line[5:].strip()
      if payload and payload != "[DONE]":
        yield payload


def _http_detail(e):
  try:
    err = json.loads(e.read().decode("utf-8")).get("error", {})
    return err.get("message", "") if isinstance(err, dict) else str(err)
  except Exception:
    return ""


def gemini_stream(api_key, model, system_text, history, stop=None, timeout=180):
  """串流版 Gemini。逐段 yield ('text', 文字)，最後可能 yield ('usage', dict)。"""
  url = (
      "https://generativelanguage.googleapis.com/v1beta/models/"
      f"{model}:streamGenerateContent?alt=sse"
  )
  body = {
      "system_instruction": {"parts": [{"text": system_text}]},
      "contents": [{"role": r, "parts": [{"text": t}]} for r, t in history],
  }
  req = urllib.request.Request(
      url,
      data=json.dumps(body).encode("utf-8"),
      headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
      method="POST",
  )
  try:
    resp = urllib.request.urlopen(req, timeout=timeout)
  except urllib.error.HTTPError as e:
    raise RuntimeError(f"HTTP {e.code}：{_http_detail(e) or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  usage = {}
  with resp:
    for payload in _sse_events(resp, stop):
      try:
        data = json.loads(payload)
      except ValueError:
        continue
      if data.get("usageMetadata"):
        usage = data["usageMetadata"]
      cands = data.get("candidates") or []
      if cands:
        parts = (cands[0].get("content") or {}).get("parts") or []
        text = "".join(p.get("text", "") for p in parts)
        if text:
          yield ("text", text)
      elif data.get("promptFeedback"):
        raise RuntimeError(f"沒有收到回答（可能被安全機制擋下）：{data['promptFeedback']}")
  if usage:
    yield ("usage", usage)


def openai_stream(api_key, base_url, model, system_text, history, stop=None, timeout=180):
  """串流版 OpenAI 格式。yield 規則同 gemini_stream。"""
  url = (base_url.strip() or OPENAI_DEFAULT_URL).rstrip("/") + "/chat/completions"
  msgs = [{"role": "system", "content": system_text}]
  for r, t in history:
    msgs.append({"role": "assistant" if r == "model" else "user", "content": t})
  resp = None
  for with_usage in (True, False):  # 有些相容服務不支援 stream_options，失敗就退回不帶
    body = {"model": model, "messages": msgs, "stream": True}
    if with_usage:
      body["stream_options"] = {"include_usage": True}
    req = urllib.request.Request(
        url,
        data=json.dumps(body).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": "Bearer " + api_key,
        },
        method="POST",
    )
    try:
      resp = urllib.request.urlopen(req, timeout=timeout)
      break
    except urllib.error.HTTPError as e:
      if with_usage and e.code == 400:
        continue
      raise RuntimeError(f"HTTP {e.code}：{_http_detail(e) or e.reason}")
    except urllib.error.URLError as e:
      raise RuntimeError(f"連線失敗：{e.reason}")
  usage = {}
  with resp:
    for payload in _sse_events(resp, stop):
      try:
        data = json.loads(payload)
      except ValueError:
        continue
      u = data.get("usage")
      if u:
        usage = {"totalTokenCount": u.get("total_tokens")}
      for ch in data.get("choices") or []:
        text = (ch.get("delta") or {}).get("content")
        if text:
          yield ("text", text)
  if usage:
    yield ("usage", usage)


def gemini_generate(api_key, model, system_text, history, timeout=180):
  """呼叫 Gemini REST API。history: [(role, text)]，role 為 'user' 或 'model'。回傳 (文字, 用量dict)。"""
  url = (
      "https://generativelanguage.googleapis.com/v1beta/models/"
      f"{model}:generateContent"
  )
  body = {
      "system_instruction": {"parts": [{"text": system_text}]},
      "contents": [{"role": r, "parts": [{"text": t}]} for r, t in history],
  }
  req = urllib.request.Request(
      url,
      data=json.dumps(body).encode("utf-8"),
      headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
      method="POST",
  )
  try:
    with urllib.request.urlopen(req, timeout=timeout) as resp:
      data = json.loads(resp.read().decode("utf-8"))
  except urllib.error.HTTPError as e:
    try:
      detail = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
    except Exception:
      detail = ""
    raise RuntimeError(f"HTTP {e.code}：{detail or e.reason}")
  except urllib.error.URLError as e:
    raise RuntimeError(f"連線失敗：{e.reason}")
  cands = data.get("candidates") or []
  if not cands:
    fb = data.get("promptFeedback", {})
    raise RuntimeError(f"沒有收到回答（可能被安全機制擋下）：{fb}")
  parts = (cands[0].get("content") or {}).get("parts") or []
  text = "".join(p.get("text", "") for p in parts)
  if not text:
    text = f"（沒有文字內容，結束原因：{cands[0].get('finishReason')}）"
  return text, data.get("usageMetadata", {})



MD_INLINE = re.compile(
    r"\*\*(.+?)\*\*|`([^`]+)`|(?<![\*A-Za-z0-9])\*(?!\s)(.+?)(?<!\s)\*(?![\*A-Za-z0-9])"
    r"|\[([^\]]+)\]\(([^)]+)\)"
)


def md_setup_tags(txt, fam):
  mono = "Consolas"
  txt.tag_configure("md_h1", font=(fam, 15, "bold"), spacing1=8, spacing3=4)
  txt.tag_configure("md_h2", font=(fam, 13, "bold"), spacing1=6, spacing3=3)
  txt.tag_configure("md_h3", font=(fam, 11, "bold"), spacing1=4, spacing3=2)
  txt.tag_configure("md_b", font=(fam, 10, "bold"))
  txt.tag_configure("md_i", font=(fam, 10, "italic"))
  txt.tag_configure("md_code", font=(mono, 10), background=PAL["code_bg"])
  txt.tag_configure(
      "md_block", font=(mono, 10), background=PAL["block_bg"], lmargin1=14, lmargin2=14
  )
  txt.tag_configure("md_table", font=(mono, 10), lmargin1=8, lmargin2=8)
  txt.tag_configure("md_quote", foreground=PAL["muted"], lmargin1=14, lmargin2=14)
  txt.tag_configure("md_link", foreground=PAL["link"], underline=True)
  txt.tag_configure("md_hr", foreground=PAL["faint"])


def _disp_width(t):
  return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in t)


def _md_inline(txt, text, base=()):
  pos = 0
  for m in MD_INLINE.finditer(text):
    if m.start() > pos:
      txt.insert(tk.END, text[pos:m.start()], base)
    if m.group(1) is not None:
      txt.insert(tk.END, m.group(1), base + ("md_b",))
    elif m.group(2) is not None:
      txt.insert(tk.END, m.group(2), base + ("md_code",))
    elif m.group(3) is not None:
      txt.insert(tk.END, m.group(3), base + ("md_i",))
    else:
      txt.insert(tk.END, m.group(4), base + ("md_link",))
      txt.insert(tk.END, f" ({m.group(5)})", base)
    pos = m.end()
  if pos < len(text):
    txt.insert(tk.END, text[pos:], base)


def _md_table(txt, rows):
  cols = max(len(r) for r in rows)
  rows = [r + [""] * (cols - len(r)) for r in rows]
  widths = [max(_disp_width(r[c]) for r in rows) for c in range(cols)]
  for ri, r in enumerate(rows):
    cells = [r[c] + " " * (widths[c] - _disp_width(r[c])) for c in range(cols)]
    txt.insert(tk.END, " │ ".join(cells) + "\n", ("md_table", "md_b") if ri == 0 else ("md_table",))
    if ri == 0:
      txt.insert(tk.END, "─┼─".join("─" * w for w in widths) + "\n", ("md_table",))


def render_markdown(txt, text):
  """把常見 Markdown（標題、粗斜體、程式碼、清單、引用、表格、連結）畫到 Text 元件。"""
  lines = text.replace("\r\n", "\n").split("\n")
  i = 0
  while i < len(lines):
    line = lines[i]
    st = line.strip()
    if st.startswith("```"):
      i += 1
      code = []
      while i < len(lines) and not lines[i].strip().startswith("```"):
        code.append(lines[i])
        i += 1
      i += 1
      txt.insert(tk.END, "\n".join(code) + "\n", ("md_block",))
      continue
    if st.startswith("|") and st.count("|") >= 2:
      rows = []
      while i < len(lines) and lines[i].strip().startswith("|"):
        cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
        if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
          rows.append(cells)
        i += 1
      if rows:
        _md_table(txt, rows)
      continue
    m = re.match(r"^(#{1,6})\s+(.*)$", st)
    if m:
      lvl = min(len(m.group(1)), 3)
      _md_inline(txt, m.group(2), (f"md_h{lvl}",))
      txt.insert(tk.END, "\n")
    elif re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", st):
      txt.insert(tk.END, "─" * 40 + "\n", ("md_hr",))
    elif re.match(r"^\s*[-*+]\s+", line):
      ind = len(line) - len(line.lstrip())
      body = re.sub(r"^\s*[-*+]\s+", "", line)
      txt.insert(tk.END, "　" * (ind // 2) + "• ")
      _md_inline(txt, body)
      txt.insert(tk.END, "\n")
    elif re.match(r"^\s*\d+[.)]\s+", line):
      ind = len(line) - len(line.lstrip())
      mm = re.match(r"^\s*(\d+[.)])\s+(.*)$", line)
      txt.insert(tk.END, "　" * (ind // 2) + mm.group(1) + " ")
      _md_inline(txt, mm.group(2))
      txt.insert(tk.END, "\n")
    elif st.startswith(">"):
      _md_inline(txt, st.lstrip("> ").strip(), ("md_quote",))
      txt.insert(tk.END, "\n")
    else:
      _md_inline(txt, line)
      txt.insert(tk.END, "\n")
    i += 1


class ViewPage:
  """檢視模式：日曆（看哪天有哪些 Excel）、圖表、概況。"""

  def __init__(self, app):
    self.app = app
    self.fam = app.base_family
    self.frame = ttk.Frame(app.container)
    self.frame.pack(fill="both", expand=True)
    self.infos = {}
    self.date_map = {}
    self.undated = []
    self.chart_dates = []
    self.chart_items = []
    self.last_points = []
    self.month_list = []
    self.popup = None
    self.popup_owner = None
    self.hide_job = None
    s = load_view_setting()
    rd = s.get("root_dir", "")
    self.root_dir = rd if isinstance(rd, str) else ""
    ex = s.get("extra_files", [])
    self.extra_files = (
        [p for p in ex if isinstance(p, str)] if isinstance(ex, list) else []
    )
    now = datetime.now()
    self.cal_year, self.cal_month = now.year, now.month
    self._last_wheel = 0.0
    self._build_ui()
    self.frame.bind("<Destroy>", self._on_destroy)
    self.frame.bind_all("<MouseWheel>", self._on_global_wheel, add="+")
    if self.root_dir or self.extra_files:
      self.frame.after(100, self.rescan)
    else:
      self._set_status("請先按「選擇根資料夾」，程式會遞迴找出裡面所有 Excel。")

  # ---------- 共用 ----------
  def _on_destroy(self, event):
    if event.widget is self.frame:
      self._destroy_popup()

  def _on_global_wheel(self, event):
    """滑鼠滾輪：在打勾清單上＝捲動清單；在日曆上＝上滑前一個月、下滑下一個月。"""
    try:
      if not self.frame.winfo_exists():
        return
      w = self.frame.winfo_containing(event.x_root, event.y_root)
      while w is not None:
        if isinstance(w, CheckList):
          w._on_wheel(event)
          return
        if isinstance(w, (tk.Listbox, tk.Text, ttk.Treeview)):
          return
        if w is self.tab_cal:
          now = time.time()
          if now - self._last_wheel >= 0.25:  # 避免一下子滑過好幾個月
            self._last_wheel = now
            self._shift_month(-1 if event.delta > 0 else 1)
          return
        w = w.master
    except (tk.TclError, KeyError):
      pass

  def _set_status(self, text):
    self.lbl_status.config(text=text)

  def _save_setting(self):
    save_view_setting(
        {"root_dir": self.root_dir, "extra_files": self.extra_files}
    )

  def _build_ui(self):
    f = self.frame
    ttk.Label(f, text="【檢視】數據查看與圖表", style="H2.TLabel").pack(
        pady=(4, 6)
    )
    bar = ttk.Frame(f)
    bar.pack(fill="x", padx=12)
    for text, cmd in (
        ("📁 選擇根資料夾", self.choose_root),
        ("＋ 額外加入 Excel", self.add_extra),
        ("管理額外檔案", self.manage_extra),
        ("🔄 重新掃描", self.rescan),
        ("⚙ 取值範圍設定", self.open_range),
    ):
      ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=3)
    ttk.Button(bar, text="返回主選單", command=self.app.show_main_home).pack(
        side="right", padx=3
    )
    self.lbl_status = ttk.Label(f, text="", style="Hint.TLabel")
    self.lbl_status.pack(anchor="w", padx=14, pady=(6, 2))

    self.nb = ttk.Notebook(f)
    self.nb.pack(fill="both", expand=True, padx=10, pady=6)
    self.tab_cal = ttk.Frame(self.nb, padding=6)
    self.tab_chart = ttk.Frame(self.nb, padding=6)
    self.tab_over = ttk.Frame(self.nb, padding=6)
    self.tab_ai = ttk.Frame(self.nb, padding=6)
    self.tab_trend = ttk.Frame(self.nb, padding=6)
    self.nb.add(self.tab_cal, text="📅 日曆")
    self.nb.add(self.tab_chart, text="📈 圖表")
    self.nb.add(self.tab_over, text="📋 概況")
    self.nb.add(self.tab_trend, text="📊 趨勢比較")
    self.nb.add(self.tab_ai, text="🤖 AI 分析")
    self._build_calendar_tab()
    self._build_chart_tab()
    self._build_overview_tab()
    self._build_trend_tab()
    self._build_ai_tab()

  # ---------- 檔案來源 ----------
  def choose_root(self):
    d = filedialog.askdirectory(
        title="選擇存放 Excel 的根資料夾（會遞迴搜尋）",
        initialdir=self.root_dir or None,
    )
    if d:
      self.root_dir = d
      self._save_setting()
      self.rescan()

  def _add_extra_paths(self, files):
    have = {os.path.normcase(os.path.abspath(p)) for p in self.extra_files}
    added = 0
    for p in files:
      key = os.path.normcase(os.path.abspath(p))
      if p.lower().endswith(EXCEL_EXTS) and key not in have:
        self.extra_files.append(p)
        have.add(key)
        added += 1
    if added:
      self._save_setting()
    return added

  def add_extra(self):
    files = filedialog.askopenfilenames(
        title="額外加入 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")],
    )
    if files and self._add_extra_paths(files):
      self.rescan()

  def manage_extra(self):
    top = tk.Toplevel(self.app)
    top.title("管理額外加入的 Excel")
    self.app.center_dialog(top, 680, 400)
    top.grab_set()
    changed = [False]
    body = ttk.Frame(top, padding=12)
    body.pack(fill="both", expand=True)
    ttk.Label(body, text="這些檔案不在根資料夾裡，是額外加入的：").pack(
        anchor="w"
    )
    lb = tk.Listbox(
        body, selectmode=tk.EXTENDED, font=(self.fam, 10), activestyle="none",
        **listbox_colors()
    )
    lb.pack(fill="both", expand=True, pady=6)
    for p in self.extra_files:
      lb.insert(tk.END, p)

    def add():
      files = filedialog.askopenfilenames(
          parent=top,
          title="額外加入 Excel 檔案",
          filetypes=[("Excel 檔案", "*.xlsx *.xlsm *.xls"), ("所有檔案", "*.*")],
      )
      if files and self._add_extra_paths(files):
        changed[0] = True
        lb.delete(0, tk.END)
        for p in self.extra_files:
          lb.insert(tk.END, p)

    def remove():
      sel = list(lb.curselection())
      if not sel:
        messagebox.showwarning("提示", "請先選取要移除的檔案！", parent=top)
        return
      for i in reversed(sel):
        del self.extra_files[i]
        lb.delete(i)
      self._save_setting()
      changed[0] = True

    def close():
      top.destroy()
      if changed[0]:
        self.rescan()

    bf = ttk.Frame(body)
    bf.pack(fill="x")
    ttk.Button(bf, text="＋ 加入", command=add).pack(side="left", padx=3)
    ttk.Button(bf, text="移除選取", command=remove).pack(side="left", padx=3)
    ttk.Button(bf, text="關閉", style="Accent.TButton", command=close).pack(
        side="right", padx=3
    )
    top.protocol("WM_DELETE_WINDOW", close)

  def _collect_paths(self):
    paths = []
    if self.root_dir and os.path.isdir(self.root_dir):
      for root, _dirs, names in os.walk(self.root_dir):
        for n in names:
          if n.lower().endswith(EXCEL_EXTS) and not n.startswith("~$"):
            paths.append(os.path.join(root, n))
    for p in self.extra_files:
      if os.path.isfile(p):
        paths.append(p)
    seen = set()
    out = []
    for p in paths:
      key = os.path.normcase(os.path.abspath(p))
      if key not in seen:
        seen.add(key)
        out.append(p)
    out.sort(key=lambda p: os.path.basename(p).lower())
    return out

  def rescan(self):
    paths = self._collect_paths()
    if not paths:
      self.infos = {}
      self._refresh_all()
      if self.root_dir and not os.path.isdir(self.root_dir):
        self._set_status(f"找不到根資料夾：{self.root_dir}")
      else:
        self._set_status("沒有找到任何 Excel。請選擇根資料夾或額外加入檔案。")
      return
    infos = self.app.load_infos(paths, title="掃描 Excel 檔案中…")
    if infos is None:
      return
    self.infos = infos
    self._refresh_all()

  def open_range(self):
    paths = [p for p, i in self.infos.items() if not i.get("error")]
    if self.app.open_range_dialog(paths):
      self._refresh_overview()
      self._trend_update()

  def _refresh_all(self):
    self.date_map = {}
    self.undated = []
    for info in self.infos.values():
      d = "" if info.get("error") else self.app.date_text_of(info)
      if d:
        self.date_map.setdefault(d, []).append(info)
      else:
        self.undated.append(info)
    for lst in self.date_map.values():
      lst.sort(key=lambda i: i["filename"])
    self.undated.sort(key=lambda i: i["filename"])

    months = sorted({d[:7] for d in self.date_map}, reverse=True)
    self.month_list = months
    self.cmb_month["values"] = [
        f"{ym[:4]}年{int(ym[5:])}月（{sum(1 for d in self.date_map if d.startswith(ym))} 天）"
        for ym in months
    ]
    cur = f"{self.cal_year:04d}-{self.cal_month:02d}"
    if months and cur not in months:
      self.cal_year, self.cal_month = int(months[0][:4]), int(months[0][5:])

    n_dated = sum(len(v) for v in self.date_map.values())
    n_err = sum(1 for i in self.infos.values() if i.get("error"))
    where = f"根資料夾：{self.root_dir}　" if self.root_dir else ""
    self._set_status(
        f"{where}共 {len(self.infos)} 個 Excel（有日期 {n_dated}、"
        f"無日期 {len(self.undated) - n_err}、讀取失敗 {n_err}）"
    )
    self._render_calendar()
    self._fill_chart_lists()
    self._refresh_overview()
    self._refresh_trend_choices()
    self._refresh_ai_files()

  # ---------- 日曆 ----------
  def _build_calendar_tab(self):
    t = self.tab_cal
    nav = ttk.Frame(t)
    nav.pack(fill="x", pady=(0, 4))
    ttk.Button(nav, text="◀", width=3, command=lambda: self._shift_month(-1)).pack(
        side="left"
    )
    self.lbl_month = ttk.Label(nav, text="", style="H2.TLabel", width=12, anchor="center")
    self.lbl_month.pack(side="left", padx=8)
    ttk.Button(nav, text="▶", width=3, command=lambda: self._shift_month(1)).pack(
        side="left"
    )
    ttk.Button(nav, text="今天", command=self._goto_today).pack(
        side="left", padx=(10, 0)
    )
    ttk.Label(nav, text="跳到有資料的月份：").pack(side="left", padx=(20, 4))
    self.cmb_month = ttk.Combobox(nav, state="readonly", width=18)
    self.cmb_month.pack(side="left")
    self.cmb_month.bind("<<ComboboxSelected>>", self._on_pick_month)
    ttk.Label(
        t,
        text="雙擊 Excel 名稱可用 Excel 開啟；滑鼠移到有資料的日期上會浮出完整列表；在日曆上滾輪往上＝上個月、往下＝下個月。",
        style="Hint.TLabel",
    ).pack(anchor="w", pady=(0, 4))
    self.cal_holder = ttk.Frame(t)
    self.cal_holder.pack(fill="both", expand=True)
    self.box_undated = ttk.LabelFrame(
        t, text="無日期或讀取失敗的檔案（雙擊開啟）", padding=4
    )
    self.lst_undated = tk.Listbox(
        self.box_undated, height=3, font=(self.fam, 9), activestyle="none",
        **listbox_colors()
    )
    self.lst_undated.pack(fill="x")
    self.lst_undated.bind("<Double-Button-1>", self._open_undated)
    self._undated_paths = []

  def _shift_month(self, delta):
    m = self.cal_month + delta
    y = self.cal_year
    if m < 1:
      m, y = 12, y - 1
    elif m > 12:
      m, y = 1, y + 1
    self.cal_year, self.cal_month = y, m
    self._render_calendar()

  def _goto_today(self):
    now = datetime.now()
    self.cal_year, self.cal_month = now.year, now.month
    self._render_calendar()

  def _on_pick_month(self, event=None):
    i = self.cmb_month.current()
    if 0 <= i < len(self.month_list):
      ym = self.month_list[i]
      self.cal_year, self.cal_month = int(ym[:4]), int(ym[5:])
      self._render_calendar()

  def _open_undated(self, event=None):
    sel = self.lst_undated.curselection()
    if sel and sel[0] < len(self._undated_paths):
      open_file(self._undated_paths[sel[0]])

  def _link(self, parent, text, bg, path, size=9, bold=False):
    lbl = tk.Label(
        parent,
        text=text,
        fg=CAL_LINK,
        bg=bg,
        cursor="hand2",
        anchor="w",
        font=(self.fam, size, "bold" if bold else "normal"),
    )
    lbl.bind("<Double-Button-1>", lambda e, p=path: open_file(p))
    return lbl

  def _render_calendar(self):
    for w in self.cal_holder.winfo_children():
      w.destroy()
    self._destroy_popup()
    y, m = self.cal_year, self.cal_month
    self.lbl_month.config(text=f"{y} 年 {m} 月")
    ym = f"{y:04d}-{m:02d}"
    if ym in self.month_list:
      self.cmb_month.current(self.month_list.index(ym))
    else:
      self.cmb_month.set("")

    grid = tk.Frame(self.cal_holder, bg=self.app.theme_bg)
    grid.pack(fill="both", expand=True)
    for c, name in enumerate(["日", "一", "二", "三", "四", "五", "六"]):
      tk.Label(
          grid, text=name, bg=self.app.theme_bg, fg=PAL["fg"],
          font=(self.fam, 10, "bold"),
      ).grid(row=0, column=c, sticky="nsew", pady=(0, 2))
      grid.columnconfigure(c, weight=1, uniform="day")
    weeks = calendar.Calendar(firstweekday=6).monthdayscalendar(y, m)
    today = datetime.now().strftime("%Y-%m-%d")
    for r, week in enumerate(weeks):
      grid.rowconfigure(r + 1, weight=1, minsize=self.app.px(90))
      for c, day in enumerate(week):
        if day == 0:
          continue
        date_str = f"{y:04d}-{m:02d}-{day:02d}"
        infos = self.date_map.get(date_str, [])
        has = bool(infos)
        bg = CAL_HAS if has else PAL["card"]
        is_today = date_str == today
        cell = tk.Frame(
            grid,
            bg=bg,
            highlightthickness=2 if is_today else 1,
            highlightbackground=CAL_TODAY if is_today else CAL_BORDER,
        )
        cell.grid(row=r + 1, column=c, sticky="nsew", padx=1, pady=1)
        cell.pack_propagate(False)
        tk.Label(
            cell,
            text=str(day),
            bg=bg,
            fg=PAL["fg"] if has else CAL_MUTED,
            font=(self.fam, 10, "bold" if has else "normal"),
            anchor="w",
        ).pack(anchor="w", padx=4, pady=(2, 0))
        for info in infos[:2]:
          self._link(
              cell, "📄 " + shorten(info["filename"], 17), bg, info["path"]
          ).pack(anchor="w", padx=4)
        if len(infos) > 2:
          tk.Label(
              cell,
              text=f"+{len(infos) - 2}",
              bg=bg,
              fg=CAL_MUTED,
              font=(self.fam, 9, "bold"),
              anchor="w",
          ).pack(anchor="w", padx=4)
        if has:
          self._bind_hover(cell, infos)

    # 無日期／讀取失敗的檔案
    self._undated_paths = []
    self.lst_undated.delete(0, tk.END)
    for info in self.undated:
      if info.get("error"):
        self.lst_undated.insert(tk.END, f"⚠ 讀取失敗：{info['filename']}（{info['error']}）")
      else:
        self.lst_undated.insert(tk.END, f"（無日期）{info['filename']}")
      self._undated_paths.append(info["path"])
    if self.undated:
      self.box_undated.pack(fill="x", pady=(6, 0))
    else:
      self.box_undated.pack_forget()

  # ---------- 日曆：懸浮列表 ----------
  def _all_widgets(self, w):
    out = [w]
    for c in w.winfo_children():
      out.extend(self._all_widgets(c))
    return out

  def _bind_hover(self, cell, infos):
    for w in self._all_widgets(cell):
      w.bind("<Enter>", lambda e, c=cell, i=infos: self._show_popup(c, i), add="+")
      w.bind("<Leave>", lambda e: self._schedule_hide(), add="+")

  def _cancel_hide(self):
    if self.hide_job is not None:
      try:
        self.frame.after_cancel(self.hide_job)
      except (tk.TclError, ValueError):
        pass
      self.hide_job = None

  def _schedule_hide(self):
    self._cancel_hide()
    try:
      self.hide_job = self.frame.after(250, self._hide_if_outside)
    except tk.TclError:
      self.hide_job = None

  def _hide_if_outside(self):
    self.hide_job = None
    if self.popup is None:
      return
    try:
      px, py = self.frame.winfo_pointerxy()
      w = self.frame.winfo_containing(px, py)
    except (tk.TclError, KeyError):
      w = None
    while w is not None:
      if w is self.popup or w is self.popup_owner:
        return
      w = getattr(w, "master", None)
    self._destroy_popup()

  def _destroy_popup(self):
    self._cancel_hide()
    if self.popup is not None:
      try:
        self.popup.destroy()
      except tk.TclError:
        pass
    self.popup = None
    self.popup_owner = None

  def _show_popup(self, cell, infos):
    self._cancel_hide()
    if self.popup is not None and self.popup_owner is cell:
      return
    self._destroy_popup()
    white = PAL["card"]
    pop = tk.Toplevel(self.frame)
    pop.overrideredirect(True)
    try:
      pop.attributes("-topmost", True)
    except tk.TclError:
      pass
    pop.configure(bg=PAL["popup_border"])
    date_text = self.app.date_text_of(infos[0])
    sw, sh = pop.winfo_screenwidth(), pop.winfo_screenheight()
    limit_h = sh - 90  # 預留工作列

    holder = {"inner": None}

    def build(max_files, max_items):
      if holder["inner"] is not None:
        holder["inner"].destroy()
      inner = tk.Frame(pop, bg=white, padx=10, pady=8)
      inner.pack(padx=1, pady=1)
      holder["inner"] = inner
      tk.Label(
          inner,
          text=f"{date_text}（{len(infos)} 個 Excel）",
          bg=white,
          fg=PAL["fg"],
          font=(self.fam, 10, "bold"),
          anchor="w",
      ).pack(fill="x")
      shown = infos if max_files is None else infos[:max_files]
      for info in shown:
        self._link(
            inner, "📄 " + info["filename"], white, info["path"], 10, True
        ).pack(anchor="w", pady=(8, 0))
        items = [x["name"] for x in info["sheets"] if x["kind"] == "data"]
        if max_items == 0:
          if items:
            tk.Label(
                inner, text=f"　共 {len(items)} 個項目", bg=white, fg=PAL["text2"],
                font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
        else:
          for name in items[:max_items]:
            tk.Label(
                inner, text="　• " + name, bg=white, fg=PAL["text2"],
                font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
          if len(items) > max_items:
            tk.Label(
                inner, text=f"　…另有 {len(items) - max_items} 項", bg=white,
                fg=CAL_MUTED, font=(self.fam, 9), anchor="w",
            ).pack(anchor="w")
        t_txt, h_txt = median_reading(info)
        if t_txt:
          tk.Label(
              inner, text=f"　🌡 {t_txt}°C　💧 {h_txt}%", bg=white, fg=CAL_MUTED,
              font=(self.fam, 9), anchor="w",
          ).pack(anchor="w")
      if max_files is not None and len(infos) > max_files:
        tk.Label(
            inner, text=f"…另有 {len(infos) - max_files} 個 Excel", bg=white,
            fg=CAL_MUTED, font=(self.fam, 9), anchor="w",
        ).pack(anchor="w", pady=(6, 0))
      pop.update_idletasks()
      return pop.winfo_reqheight()

    # 內容太高放不下時，逐步縮減顯示的項目數／檔案數
    for mf, mi in ((None, 14), (None, 8), (None, 5), (None, 3), (None, 0),
                   (8, 0), (5, 0), (3, 0), (2, 0)):
      if build(mf, mi) <= limit_h:
        break
    for w in self._all_widgets(pop):
      w.bind("<Enter>", lambda e: self._cancel_hide(), add="+")
      w.bind("<Leave>", lambda e: self._schedule_hide(), add="+")
    pw, ph = pop.winfo_reqwidth(), pop.winfo_reqheight()
    x = cell.winfo_rootx() + cell.winfo_width() - 12
    y = cell.winfo_rooty() + 22
    if x + pw > sw - 10:
      x = max(10, cell.winfo_rootx() - pw + 12)
    if y + ph > sh - 70:  # 下方放不下就往上移
      y = max(10, sh - ph - 70)
    pop.geometry(f"+{x}+{y}")
    self.popup = pop
    self.popup_owner = cell

  # ---------- 圖表 ----------
  def _build_chart_tab(self):
    t = self.tab_chart
    left = ttk.Frame(t)
    left.pack(side="left", fill="y", padx=(0, 10))
    self.chart_left = left
    right = ttk.Frame(t)
    right.pack(side="left", fill="both", expand=True)
    self.btn_collapse = ttk.Button(
        right, text="◀ 隱藏設定（圖表加寬）", command=self._toggle_chart_panel
    )
    self.btn_collapse.pack(side="top", anchor="w", pady=(0, 4))

    bf = ttk.Frame(left)
    bf.pack(fill="x", pady=(0, 8))
    ttk.Button(
        bf, text="📈 畫圖", style="Accent.TButton", command=self.draw
    ).pack(side="left", padx=(0, 6))
    ttk.Button(bf, text="匯出圖上數據…", command=self.export_points).pack(side="left")
    ttk.Button(bf, text="📤 輸出整理報表…", command=self.export_report).pack(
        side="left", padx=(6, 0)
    )

    ttk.Label(left, text="軸組合").pack(anchor="w")
    self.cmb_layout = ttk.Combobox(
        left, values=LAYOUT_LABELS, state="readonly", width=40
    )
    self.cmb_layout.current(0)
    self.cmb_layout.pack(anchor="w", pady=(0, 6))

    ttk.Label(left, text="圖型").pack(anchor="w")
    tf = ttk.Frame(left)
    tf.pack(anchor="w", pady=(0, 4))
    self.var_ctype = tk.StringVar(value="scatter")
    for text, val in (("散佈圖", "scatter"), ("折線圖", "line"), ("長條圖", "bar")):
      ttk.Radiobutton(tf, text=text, value=val, variable=self.var_ctype).pack(
          side="left", padx=(0, 10)
      )
    self.var_connect = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        left, text="折線圖把各點連起來", variable=self.var_connect
    ).pack(anchor="w", pady=(0, 6))

    ttk.Label(left, text="顯示方式").pack(anchor="w")
    vf = ttk.Frame(left)
    vf.pack(anchor="w", pady=(0, 6))
    self.var_view = tk.StringVar(value="browser" if HAS_PLOTLY else "inline")
    ttk.Radiobutton(
        vf, text="瀏覽器互動圖", value="browser", variable=self.var_view
    ).pack(side="left", padx=(0, 10))
    ttk.Radiobutton(
        vf, text="視窗內（舊式）", value="inline", variable=self.var_view
    ).pack(side="left")
    self.plot_dir = None
    self.plot_files = []

    ttk.Label(left, text="y 軸取值方式").pack(anchor="w")
    self.cmb_src = ttk.Combobox(
        left, values=SRC_LABELS, state="readonly", width=40
    )
    self.cmb_src.current(0)
    self.cmb_src.pack(anchor="w")
    self.cmb_src.bind("<<ComboboxSelected>>", self._on_src_change)
    self.var_param = tk.StringVar(value="")
    self.ent_param = ttk.Entry(left, textvariable=self.var_param, width=42)
    self.ent_param.pack(anchor="w", pady=(4, 0))
    self.lbl_param = ttk.Label(
        left, text="", style="Hint.TLabel", wraplength=self.app.px(330), justify="left"
    )
    self.lbl_param.pack(anchor="w")

    self.var_map = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        left,
        text="套用 setting.json 的項目合併／順序",
        variable=self.var_map,
        command=self._fill_chart_lists,
    ).pack(anchor="w", pady=(6, 6))

    self.chk_dates = CheckList(left, "日期（打勾選取）", self.app.px(80), self.fam)
    self.chk_dates.pack(fill="x", pady=(0, 6))
    self.chk_items = CheckList(left, "項目（打勾選取）", self.app.px(130), self.fam)
    self.chk_items.pack(fill="x", pady=(0, 6))

    self.lbl_chart = ttk.Label(
        right, text="", style="Hint.TLabel", wraplength=self.app.px(640), justify="left"
    )
    self.lbl_chart.pack(side="bottom", anchor="w", pady=(4, 0))
    self.fig = None
    self.canvas = None
    if HAS_MPL and HAS_MPL_TK:
      setup_mpl_fonts()
      self.fig = Figure(figsize=(7, 5), dpi=100, facecolor=PAL["plot_bg"])
      self.canvas = FigureCanvasTkAgg(self.fig, master=right)
      toolbar = NavigationToolbar2Tk(self.canvas, right, pack_toolbar=False)
      toolbar.update()
      toolbar.pack(side="bottom", fill="x")
      self.canvas.get_tk_widget().pack(side="top", fill="both", expand=True)
    else:
      ttk.Label(
          right,
          text="畫圖需要 matplotlib。請先在命令列執行：pip install matplotlib\n安裝後重新開啟程式即可。",
          justify="left",
      ).pack(expand=True)
    self._on_src_change()

  def _toggle_chart_panel(self):
    if self.chart_left.winfo_ismapped():
      self.chart_left.pack_forget()
      self.btn_collapse.config(text="▶ 顯示設定")
    else:
      self.chart_left.pack(side="left", fill="y", padx=(0, 10), before=self.btn_collapse.master)
      self.btn_collapse.config(text="◀ 隱藏設定（圖表加寬）")

  def _on_src_change(self, event=None):
    key = SRC_KEYS[self.cmb_src.current()]
    self.var_param.set(SRC_DEFAULT_PARAM[key])
    self.ent_param.config(state="normal" if key in ("regex", "cell") else "disabled")
    self.lbl_param.config(text=SRC_HINT[key])

  def _name_func(self):
    raw_to_final = load_item_mapping()[0] if self.var_map.get() else {}
    return lambda n: raw_to_final.get(n, n)

  def _fill_chart_lists(self):
    raw_to_final, order = load_item_mapping() if self.var_map.get() else ({}, [])
    names = []
    seen = set()
    for info in self.infos.values():
      if info.get("error"):
        continue
      for sh in info["sheets"]:
        if sh["kind"] == "data":
          final = raw_to_final.get(sh["name"], sh["name"])
          if final not in seen:
            seen.add(final)
            names.append(final)
    idx = {n: i for i, n in enumerate(order)}
    names.sort(key=lambda n: (idx.get(n, 10**6), natural_key(n)))
    self.chart_items = names
    self.chk_items.set_items(names)
    self.chart_dates = sorted(self.date_map.keys())
    self.chk_dates.set_items(
        [f"{d}（{len(self.date_map[d])} 個檔案）" for d in self.chart_dates]
    )

  def draw(self):
    browser = self.var_view.get() == "browser"
    if browser and not HAS_PLOTLY:
      messagebox.showinfo(
          "提示",
          "互動圖需要 plotly，請先在命令列執行：pip install plotly\n"
          "安裝後重新開啟程式即可。\n\n（也可以把「顯示方式」改成「視窗內」。）",
      )
      return
    if not browser and self.fig is None:
      messagebox.showinfo("提示", "需要先安裝 matplotlib：pip install matplotlib")
      return
    dates = [self.chart_dates[i] for i in self.chk_dates.checked()]
    items = [self.chart_items[i] for i in self.chk_items.checked()]
    if not dates or not items:
      messagebox.showwarning("提示", "請至少選一個日期和一個項目！")
      return
    source = SRC_KEYS[self.cmb_src.current()]
    param = self.var_param.get().strip()
    if source == "regex":
      try:
        re.compile(param)
      except re.error as e:
        messagebox.showerror("正則表達式錯誤", f"無法使用這個正則表達式：\n{e}")
        return
    if source == "cell" and coord_to_index(param) is None:
      messagebox.showerror("儲存格格式錯誤", "請輸入像 I3、J2 這樣的儲存格位置！")
      return
    entries = [(info, d) for d in dates for info in self.date_map[d]]
    points, warns, conflicts = build_points(
        entries, self.app.range_settings, self._name_func(), set(items), source, param
    )
    self.last_points = points
    layout = LAYOUT_KEYS[self.cmb_layout.current()]
    if browser:
      if self.plot_dir is None:
        self.plot_dir = tempfile.mkdtemp(prefix="excel_chart_")
      path = os.path.join(self.plot_dir, f"chart_{int(time.time() * 1000)}.html")
      try:
        st = build_plotly_html(
            points, layout, self.var_ctype.get(), self.var_connect.get(),
            Y_LABELS[source], self.chart_items, path,
        )
      except Exception as e:
        messagebox.showerror("畫圖失敗", f"產生互動圖時發生錯誤：\n{e}")
        return
      if st["points"] - st["missing"] <= 0:
        messagebox.showinfo("提示", "沒有可以畫的數據。")
      else:
        self.plot_files.append(path)
        for old in self.plot_files[:-5]:  # 只留最近 5 張，避免暫存檔越積越多
          try:
            os.remove(old)
          except OSError:
            pass
        self.plot_files = self.plot_files[-5:]
        webbrowser.open("file:///" + path.replace("\\", "/"))
      if self.fig is not None:
        self.fig.clear()
        self.fig.text(
            0.5, 0.5, "圖表已在瀏覽器開啟\n（要在這裡看，請把「顯示方式」改成「視窗內」）",
            ha="center", va="center", fontsize=13, color=PAL["muted"],
        )
        self.fig.set_facecolor(PAL["plot_bg"])
        self.canvas.draw_idle()
    else:
      st = draw_chart(
          self.fig, points, layout, self.var_ctype.get(),
          self.var_connect.get(), Y_LABELS[source], self.chart_items,
      )
      self.canvas.draw_idle()
    msg = f"共 {st['points']} 個點"
    if st["missing"]:
      msg += f"，其中 {st['missing']} 個沒有數值（圖上留空）"
    if st["skipped"]:
      msg += f"；{st['skipped']} 個項目名稱沒有數字（已略過）"
    if conflicts:
      msg += f"；{len(conflicts)} 個項目有重複數據（採用第一個）"
    self.lbl_chart.config(text=msg)
    if warns:
      uniq = list(dict.fromkeys(warns))
      more = f"\n...以及其他 {len(uniq) - 12} 筆" if len(uniq) > 12 else ""
      messagebox.showwarning(
          "取值範圍警告",
          "以下工作表的範圍有問題，數值已留空：\n\n" + "\n".join(uniq[:12]) + more,
      )

  def export_report(self):
    """用目前打勾的日期、項目與取值方式，輸出成 Excel：矩陣表＋長表。"""
    dates = [self.chart_dates[i] for i in self.chk_dates.checked()]
    items = [self.chart_items[i] for i in self.chk_items.checked()]
    if not dates or not items:
      messagebox.showwarning("提示", "請至少選一個日期和一個項目！")
      return
    source = SRC_KEYS[self.cmb_src.current()]
    param = self.var_param.get().strip()
    if source == "regex":
      try:
        re.compile(param)
      except re.error as e:
        messagebox.showerror("正則表達式錯誤", f"無法使用這個正則表達式：\n{e}")
        return
    if source == "cell" and coord_to_index(param) is None:
      messagebox.showerror("儲存格格式錯誤", "請輸入像 I3、J2 這樣的儲存格位置！")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", filetypes=[("Excel 檔案", "*.xlsx")],
        title="輸出整理報表",
    )
    if not path:
      return
    entries = [(info, d) for d in dates for info in self.date_map[d]]
    points, warns, conflicts = build_points(
        entries, self.app.range_settings, self._name_func(), set(items), source, param
    )
    order = {n: i for i, n in enumerate(self.chart_items)}
    points.sort(key=lambda p: (order.get(p["item"], 10**6), p["date"], p["file"]))
    label = Y_LABELS[source]
    long_df = pd.DataFrame([
        {"日期": p["date"], "檔案": p["file"], "項目": p["item"],
         "工作表": p["sheet"], label: p["value"]}
        for p in points
    ])
    matrix = {}
    for p in points:
      row = matrix.setdefault(p["item"], {})
      if p["value"] is not None and row.get(p["date"]) is None:
        row[p["date"]] = p["value"]
      else:
        row.setdefault(p["date"], None)
    wide_df = pd.DataFrame(
        [{"項目": it, **{d: matrix[it].get(d) for d in dates}}
         for it in items if it in matrix]
    )
    try:
      with pd.ExcelWriter(path, engine="openpyxl") as xw:
        wide_df.to_excel(xw, sheet_name="矩陣（項目×日期）", index=False)
        long_df.to_excel(xw, sheet_name="長表", index=False)
        for ws in xw.book.worksheets:
          ws.freeze_panes = "B2" if ws.title.startswith("矩陣") else "A2"
          for col in ws.columns:
            w = max((_disp_width(str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 10), 45)
      msg = f"已輸出：\n{path}\n\n共 {len(points)} 筆資料。"
      if conflicts:
        msg += f"\n有 {len(conflicts)} 個項目重複數據（矩陣採用第一個，長表全部保留）。"
      if warns:
        msg += f"\n有 {len(set(warns))} 個取值範圍警告，該數值已留空。"
      messagebox.showinfo("成功", msg)
    except Exception as e:
      messagebox.showerror("輸出失敗", str(e))

  def export_points(self):
    if not self.last_points:
      messagebox.showinfo("提示", "請先畫一張圖，再匯出圖上的數據。")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV 檔案", "*.csv"), ("Excel 檔案", "*.xlsx")],
        title="匯出圖上數據",
    )
    if not path:
      return
    df = pd.DataFrame([
        {"日期": p["date"], "檔案": p["file"], "項目": p["item"],
         "工作表": p["sheet"], "數值": p["value"]}
        for p in self.last_points
    ])
    try:
      if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False)
      else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
      messagebox.showinfo("成功", f"已匯出：\n{path}")
    except Exception as e:
      messagebox.showerror("匯出失敗", str(e))

  # ---------- 趨勢比較 ----------
  def _build_trend_tab(self):
    t = self.tab_trend
    bar = ttk.Frame(t)
    bar.pack(fill="x")
    ttk.Label(bar, text="項目").pack(side="left")
    self.cmb_tr_item = ttk.Combobox(bar, state="readonly", width=16)
    self.cmb_tr_item.pack(side="left", padx=(4, 14))
    ttk.Label(bar, text="基準日期").pack(side="left")
    self.cmb_tr_base = ttk.Combobox(bar, state="readonly", width=14)
    self.cmb_tr_base.pack(side="left", padx=(4, 14))
    ttk.Button(bar, text="更新", command=self._trend_update).pack(side="left")
    ttk.Button(bar, text="匯出這張表…", command=self._trend_export).pack(
        side="left", padx=8
    )
    self.cmb_tr_item.bind("<<ComboboxSelected>>", lambda e: self._trend_update())
    self.cmb_tr_base.bind("<<ComboboxSelected>>", lambda e: self._trend_update())
    ttk.Label(
        t,
        text="用目前的取值範圍算出各日期的 Rx，並和「基準日期」比較。變化% ＝（Rx − 基準 Rx）÷ 基準 Rx × 100。",
        style="Hint.TLabel",
    ).pack(anchor="w", pady=(6, 4))
    self.lbl_trend = ttk.Label(t, text="", style="Hint.TLabel")
    self.lbl_trend.pack(side="bottom", anchor="w")
    box = ttk.Frame(t)
    box.pack(fill="both", expand=True)
    self.tree_trend = self._make_tree(box, [
        ("item", "項目", 120), ("date", "日期", 100), ("file", "檔案", 230),
        ("rx", "Rx", 100), ("diff", "與基準差", 110), ("pct", "變化 %", 100),
    ])
    self.tree_trend.tag_configure("base", foreground=PAL["accent"])
    self.tree_trend.tag_configure("none", foreground=PAL["faint"])
    self.trend_rows = []

  def _refresh_trend_choices(self):
    items = ["全部項目"] + list(self.chart_items)
    self.cmb_tr_item.config(values=items)
    if self.cmb_tr_item.get() not in items:
      self.cmb_tr_item.set("全部項目")
    dates = list(self.chart_dates)
    self.cmb_tr_base.config(values=dates)
    if self.cmb_tr_base.get() not in dates:
      self.cmb_tr_base.set(dates[0] if dates else "")
    self._trend_update()

  def _trend_update(self):
    self.tree_trend.delete(*self.tree_trend.get_children())
    self.trend_rows = []
    base = self.cmb_tr_base.get()
    sel = self.cmb_tr_item.get()
    if not self.chart_dates or not base:
      self.lbl_trend.config(text="尚無資料。")
      return
    wanted = set(self.chart_items) if sel in ("", "全部項目") else {sel}
    entries = [(i, d) for d in self.chart_dates for i in self.date_map[d]]
    points, _w, _c = build_points(
        entries, self.app.range_settings, self._name_func(), wanted, "rx", ""
    )
    basev = {}
    for p in points:
      if p["date"] == base and p["value"] is not None:
        basev.setdefault(p["item"], p["value"])
    order = {n: i for i, n in enumerate(self.chart_items)}
    points.sort(key=lambda p: (order.get(p["item"], 10**6), p["date"], p["file"]))
    nobase = set()
    for p in points:
      v = p["value"]
      b = basev.get(p["item"])
      diff = pct = ""
      tags = ()
      if v is None:
        tags = ("none",)
      elif b is None:
        nobase.add(p["item"])
      else:
        diff = f"{v - b:+.3f}"
        pct = f"{(v - b) / b * 100:+.2f}" if b != 0 else ""
        if p["date"] == base:
          tags = ("base",)
      row = (
          p["item"], p["date"], p["file"], "" if v is None else f"{v:.3f}", diff, pct,
      )
      self.trend_rows.append(row)
      self.tree_trend.insert("", "end", values=row, tags=tags)
    msg = f"共 {len(points)} 筆。藍色＝基準日期；灰色＝沒有數值。"
    if nobase:
      msg += f"　{len(nobase)} 個項目在基準日期沒有數值，無法比較。"
    self.lbl_trend.config(text=msg)

  def _trend_export(self):
    if not self.trend_rows:
      messagebox.showinfo("提示", "目前沒有可匯出的資料。")
      return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV 檔案", "*.csv"), ("Excel 檔案", "*.xlsx")],
        title="匯出趨勢比較表",
    )
    if not path:
      return
    df = pd.DataFrame(
        self.trend_rows, columns=["項目", "日期", "檔案", "Rx", "與基準差", "變化%"]
    )
    for c in ("Rx", "與基準差", "變化%"):
      df[c] = pd.to_numeric(df[c], errors="coerce")
    try:
      if path.lower().endswith(".xlsx"):
        df.to_excel(path, index=False)
      else:
        df.to_csv(path, index=False, encoding="utf-8-sig")
      messagebox.showinfo("成功", f"已匯出：\n{path}")
    except Exception as e:
      messagebox.showerror("匯出失敗", str(e))

  # ---------- AI 分析（Gemini） ----------
  def _build_ai_tab(self):
    t = self.tab_ai
    gs = load_gemini_setting()
    self.ai_history = []
    self.ai_data_text = ""
    self.ai_files = []
    self.ai_queue = queue.Queue()
    self.ai_busy = False
    self.ai_stop = threading.Event()
    self.stream_text = ""
    self.stream_usage = {}
    self._stream_dirty = False
    self._stream_t = 0.0

    def sub(name, defaults):
      d = gs.get(name)
      d = d if isinstance(d, dict) else {}
      out = dict(defaults)
      for k in defaults:
        if isinstance(d.get(k), str):
          out[k] = d[k]
      return out

    old_key = gs.get("api_key") if isinstance(gs.get("api_key"), str) else ""
    self.ai_cfg = {
        "Gemini": sub("gemini", {
            "api_key": old_key,
            "model": gs.get("model") if isinstance(gs.get("model"), str) and gs.get("model") else GEMINI_DEFAULT_MODEL,
        }),
        "OpenAI 格式": sub("openai", {
            "api_key": "", "model": OPENAI_DEFAULT_MODEL, "base_url": OPENAI_DEFAULT_URL,
        }),
    }
    prov = gs.get("provider") if gs.get("provider") in PROVIDERS else PROVIDERS[0]
    self.ai_prov = prov

    top = ttk.Frame(t)
    top.pack(fill="x")
    ttk.Label(top, text="服務").grid(row=0, column=0, sticky="w")
    self.cmb_prov = ttk.Combobox(top, values=PROVIDERS, state="readonly", width=12)
    self.cmb_prov.set(prov)
    self.cmb_prov.grid(row=0, column=1, padx=(4, 12))
    self.cmb_prov.bind("<<ComboboxSelected>>", self._ai_on_provider)
    ttk.Label(top, text="金鑰").grid(row=0, column=2, sticky="w")
    self.var_key = tk.StringVar()
    ttk.Entry(top, textvariable=self.var_key, show="*", width=30).grid(
        row=0, column=3, padx=(4, 12)
    )
    ttk.Label(top, text="模型").grid(row=0, column=4, sticky="w")
    self.var_model = tk.StringVar()
    ttk.Entry(top, textvariable=self.var_model, width=22).grid(row=0, column=5, padx=(4, 12))
    self.var_remember = tk.BooleanVar(value=bool(gs.get("remember", False)))
    ttk.Checkbutton(
        top, text="記住金鑰（明文存在 gemini_setting.json）", variable=self.var_remember
    ).grid(row=0, column=6)
    ttk.Label(top, text="網址").grid(row=1, column=0, sticky="w", pady=(4, 0))
    self.var_url = tk.StringVar()
    self.ent_url = ttk.Entry(top, textvariable=self.var_url, width=48)
    self.ent_url.grid(row=1, column=1, columnspan=5, sticky="w", padx=(4, 0), pady=(4, 0))
    ttk.Label(
        top, text="（OpenAI 格式才需要；相容服務可改這裡）", style="Hint.TLabel"
    ).grid(row=1, column=6, sticky="w", pady=(4, 0))
    self._ai_load_fields()

    body = ttk.Frame(t)
    body.pack(fill="both", expand=True, pady=(8, 0))
    left = ttk.Frame(body)
    left.pack(side="left", fill="y", padx=(0, 10))
    self.chk_ai = CheckList(left, "要餵給 AI 的 Excel（打勾）", self.app.px(300), self.fam)
    self.chk_ai.pack(fill="x")
    rf = ttk.Frame(left)
    rf.pack(fill="x", pady=(8, 0))
    ttk.Label(rf, text="每個工作表最多取前").pack(side="left")
    self.var_rows = tk.StringVar(value="500")
    ttk.Entry(rf, textvariable=self.var_rows, width=7).pack(side="left", padx=4)
    ttk.Label(rf, text="列（留空＝全部）").pack(side="left")
    ttk.Button(
        left, text="載入選取的 Excel 成 CSV", style="Accent.TButton",
        command=self._ai_load_data,
    ).pack(fill="x", pady=(8, 0))
    self.lbl_ai_data = ttk.Label(
        left, text="尚未載入資料", style="Hint.TLabel",
        wraplength=self.app.px(300), justify="left",
    )
    self.lbl_ai_data.pack(anchor="w", pady=(4, 0))

    right = ttk.Frame(body)
    right.pack(side="left", fill="both", expand=True)
    box = ttk.Frame(right)
    box.pack(fill="both", expand=True)
    self.txt_chat = tk.Text(
        box, wrap="word", state="disabled", font=(self.fam, 10), height=12,
        padx=8, pady=6, bg=PAL["card"], fg=PAL["fg"], insertbackground=PAL["fg"],
        relief="flat", highlightthickness=1, highlightbackground=PAL["card_border"],
    )
    sb = ttk.Scrollbar(box, orient="vertical", command=self.txt_chat.yview)
    self.txt_chat.configure(yscrollcommand=sb.set)
    self.txt_chat.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    md_setup_tags(self.txt_chat, self.fam)
    self.txt_chat.tag_configure("me", foreground=PAL["accent"], font=(self.fam, 10, "bold"))
    self.txt_chat.tag_configure("ai", foreground=PAL["ai"], font=(self.fam, 10, "bold"))
    self.txt_chat.tag_configure("sys", foreground=PAL["faint"])

    inp = ttk.Frame(right)
    inp.pack(fill="x", pady=(6, 0))
    self.txt_ask = tk.Text(
        inp, height=4, wrap="word", font=(self.fam, 10), padx=6, pady=4,
        bg=PAL["card"], fg=PAL["fg"], insertbackground=PAL["fg"], relief="flat",
        highlightthickness=1, highlightbackground=PAL["card_border"],
        highlightcolor=PAL["accent"],
    )
    self.txt_ask.pack(side="left", fill="x", expand=True)
    self.txt_ask.bind("<Return>", self._ai_enter)
    bf = ttk.Frame(inp)
    bf.pack(side="left", padx=(8, 0))
    self.btn_ask = ttk.Button(bf, text="送出", style="Accent.TButton", command=self._ai_send)
    self.btn_ask.pack(fill="x")
    self.btn_stop = ttk.Button(bf, text="停止", command=self._ai_stop, state="disabled")
    self.btn_stop.pack(fill="x", pady=(4, 0))
    ttk.Button(bf, text="清空對話", command=self._ai_clear).pack(fill="x", pady=(4, 0))
    self.lbl_ai = ttk.Label(
        right,
        text="Enter 送出、Shift+Enter 換行。送出時會把已載入的 CSV 與對話一起傳給 Google。",
        style="Hint.TLabel",
    )
    self.lbl_ai.pack(anchor="w", pady=(4, 0))

  def _ai_load_fields(self):
    c = self.ai_cfg[self.ai_prov]
    self.var_key.set(c["api_key"])
    self.var_model.set(c["model"])
    self.var_url.set(c.get("base_url", ""))
    self.ent_url.config(state="normal" if self.ai_prov == "OpenAI 格式" else "disabled")

  def _ai_store_fields(self):
    c = self.ai_cfg[self.ai_prov]
    c["api_key"] = self.var_key.get().strip()
    c["model"] = self.var_model.get().strip()
    if "base_url" in c:
      c["base_url"] = self.var_url.get().strip()

  def _ai_on_provider(self, event=None):
    self._ai_store_fields()
    self.ai_prov = self.cmb_prov.get()
    self._ai_load_fields()

  def _refresh_ai_files(self):
    infos = [i for i in self.infos.values() if not i.get("error")]
    infos.sort(key=lambda i: (self.app.date_text_of(i) or "9999", i["filename"]))
    self.ai_files = infos
    self.chk_ai.set_items(
        [f"{self.app.date_text_of(i) or '無日期'}　{i['filename']}" for i in infos],
        checked=False,
    )

  def _chat_add(self, who, text):
    tag = {"我": "me", "AI": "ai"}.get(who, "sys")
    self.txt_chat.config(state="normal")
    self.txt_chat.insert(tk.END, f"{who}：", tag)
    if who == "AI":
      self.txt_chat.insert(tk.END, "\n")
      render_markdown(self.txt_chat, text)
      self.txt_chat.insert(tk.END, "\n")
    else:
      self.txt_chat.insert(tk.END, text + "\n\n")
    self.txt_chat.config(state="disabled")
    self.txt_chat.see(tk.END)

  def _ai_clear(self):
    if self.ai_busy:
      return
    self.ai_history = []
    self.txt_chat.config(state="normal")
    self.txt_chat.delete("1.0", tk.END)
    self.txt_chat.config(state="disabled")

  def _ai_load_data(self):
    idx = self.chk_ai.checked()
    if not idx:
      messagebox.showwarning("提示", "請先勾選至少一個 Excel！")
      return
    raw = self.var_rows.get().strip()
    limit = None
    if raw:
      try:
        limit = int(raw)
        if limit < 1:
          raise ValueError
      except ValueError:
        messagebox.showerror("格式錯誤", "列數請輸入大於 0 的整數，或留空。")
        return
    prog = ProgressWindow(self.app, self.frame, "轉換成 CSV", len(idx))
    parts, notes = [], []
    try:
      for i in idx:
        info = self.ai_files[i]
        prog.set_text(info["filename"])
        text, nt = build_ai_csv(info["path"], limit)
        if text:
          parts.append(text)
        notes.extend(nt)
        prog.step()
        if prog.cancelled:
          return
    finally:
      prog.close()
    self.ai_data_text = "\n".join(parts)
    n = len(self.ai_data_text)
    msg = f"已載入 {len(parts)} 個檔案，約 {n:,} 個字元（約 {n // 3:,}～{n // 2:,} tokens）。"
    if notes:
      msg += f"\n{len(notes)} 個工作表被截斷。"
    if n > 1_500_000:
      msg += "\n⚠ 資料很大，可能超過模型上限或很花額度，建議減少檔案或列數。"
    self.lbl_ai_data.config(text=msg)
    self._ai_clear()
    self._chat_add("系統", "資料已載入，可以開始提問。（換資料會清空對話）")

  def _ai_enter(self, event):
    if event.state & 0x1:  # Shift
      return None
    self._ai_send()
    return "break"

  def _ai_send(self):
    if self.ai_busy:
      return
    q = self.txt_ask.get("1.0", tk.END).strip()
    if not q:
      return
    self._ai_store_fields()
    prov = self.ai_prov
    cfg = self.ai_cfg[prov]
    env = "GEMINI_API_KEY" if prov == "Gemini" else "OPENAI_API_KEY"
    key = cfg["api_key"] or os.environ.get(env, "").strip()
    model = cfg["model"] or (GEMINI_DEFAULT_MODEL if prov == "Gemini" else OPENAI_DEFAULT_MODEL)
    base_url = cfg.get("base_url", "")
    if not key:
      messagebox.showwarning("提示", f"請先輸入 {prov} 的 API 金鑰（或設定環境變數 {env}）。")
      return
    if not self.ai_data_text:
      if not messagebox.askyesno("尚未載入資料", "還沒載入任何 Excel，AI 看不到你的數據。\n仍要直接提問嗎？"):
        return
    rem = bool(self.var_remember.get())
    save_gemini_setting({
        "provider": prov,
        "remember": rem,
        "gemini": {
            "model": self.ai_cfg["Gemini"]["model"],
            "api_key": self.ai_cfg["Gemini"]["api_key"] if rem else "",
        },
        "openai": {
            "model": self.ai_cfg["OpenAI 格式"]["model"],
            "base_url": self.ai_cfg["OpenAI 格式"]["base_url"],
            "api_key": self.ai_cfg["OpenAI 格式"]["api_key"] if rem else "",
        },
    })
    self.txt_ask.delete("1.0", tk.END)
    self._chat_add("我", q)
    self.ai_history.append(("user", q))
    system = AI_SYSTEM + ("\n\n以下是數據：\n\n" + self.ai_data_text if self.ai_data_text else "")
    hist = list(self.ai_history)
    self.ai_busy = True
    self.ai_stop.clear()
    self.stream_text = ""
    self.stream_usage = {}
    self._stream_dirty = False
    self.btn_ask.config(state="disabled")
    self.btn_stop.config(state="normal")
    self.lbl_ai.config(text="AI 回答中…（可按「停止」）")
    self._stream_begin()
    stop = self.ai_stop

    def work():
      gen = None
      try:
        if prov == "Gemini":
          gen = gemini_stream(key, model, system, hist, stop)
        else:
          gen = openai_stream(key, base_url, model, system, hist, stop)
        for kind, val in gen:
          if stop.is_set():
            break
          self.ai_queue.put(("chunk" if kind == "text" else "usage", val))
        self.ai_queue.put(("stopped" if stop.is_set() else "done", None))
      except Exception as e:
        self.ai_queue.put(("err", str(e)))
      finally:
        if gen is not None:
          try:
            gen.close()
          except Exception:
            pass

    threading.Thread(target=work, daemon=True).start()
    self.frame.after(60, self._ai_poll)

  def _ai_stop(self):
    self.ai_stop.set()
    self.btn_stop.config(state="disabled")

  def _stream_begin(self):
    t = self.txt_chat
    t.config(state="normal")
    t.mark_set("ai_head", tk.END + "-1c")
    t.mark_gravity("ai_head", "left")
    t.insert(tk.END, "AI：", "ai")
    t.insert(tk.END, "\n")
    t.mark_set("ai_body", tk.END + "-1c")
    t.mark_gravity("ai_body", "left")
    t.config(state="disabled")

  def _stream_render(self):
    t = self.txt_chat
    t.config(state="normal")
    t.delete("ai_body", tk.END)
    render_markdown(t, self.stream_text)
    t.config(state="disabled")
    t.see(tk.END)
    self._stream_dirty = False
    self._stream_t = time.time()

  def _ai_poll(self):
    done = None
    try:
      while True:
        kind, payload = self.ai_queue.get_nowait()
        if kind == "chunk":
          self.stream_text += payload
          self._stream_dirty = True
        elif kind == "usage":
          self.stream_usage = payload
        else:
          done = (kind, payload)
          break
    except queue.Empty:
      pass
    try:
      if self._stream_dirty and (done or time.time() - self._stream_t >= 0.2):
        self._stream_render()
      if done is None:
        self.frame.after(60, self._ai_poll)
        return
      self.ai_busy = False
      self.btn_ask.config(state="normal")
      self.btn_stop.config(state="disabled")
      kind, payload = done
      text = self.stream_text
      if kind == "err":
        self.ai_history.pop()  # 失敗的問題不留在對話裡
        if not text:
          self.txt_chat.config(state="normal")
          self.txt_chat.delete("ai_head", tk.END)
          self.txt_chat.config(state="disabled")
        self._chat_add("系統", "發生錯誤：" + payload)
        self.lbl_ai.config(text="發生錯誤，請檢查金鑰、模型名稱或網路。")
        return
      if kind == "stopped" and not text:
        self.ai_history.pop()
        self.txt_chat.config(state="normal")
        self.txt_chat.delete("ai_head", tk.END)
        self.txt_chat.config(state="disabled")
        self.lbl_ai.config(text="已停止。")
        return
      self.ai_history.append(("model", text))
      self.txt_chat.config(state="normal")
      self.txt_chat.insert(tk.END, "\n")
      self.txt_chat.config(state="disabled")
      tok = self.stream_usage.get("totalTokenCount")
      head = "已停止（保留已收到的內容）。" if kind == "stopped" else "完成。"
      self.lbl_ai.config(text=head + (f"本次共用 {tok:,} tokens。" if tok else ""))
    except tk.TclError:
      pass

  # ---------- 概況 ----------
  def _make_tree(self, parent, cols):
    tree = ttk.Treeview(parent, columns=[c[0] for c in cols], show="headings")
    for cid, text, width in cols:
      tree.heading(cid, text=text)
      tree.column(cid, width=self.app.px(width), anchor="w")
    tree.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    sb.pack(side="right", fill="y")
    tree.configure(yscrollcommand=sb.set)
    return tree

  def _build_overview_tab(self):
    t = self.tab_over
    self.lbl_over = ttk.Label(t, text="", justify="left", wraplength=self.app.px(1000))
    self.lbl_over.pack(anchor="w")
    pw = ttk.PanedWindow(t, orient="vertical")
    pw.pack(fill="both", expand=True, pady=6)
    f1 = ttk.LabelFrame(pw, text="每個 Excel", padding=4)
    f2 = ttk.LabelFrame(pw, text="每個項目的 Rx 統計（依目前取值範圍設定）", padding=4)
    pw.add(f1, weight=1)
    pw.add(f2, weight=1)
    self.tree_files = self._make_tree(f1, [
        ("date", "日期", 90), ("file", "檔案", 190), ("n", "數據項目", 70),
        ("ok", "有數值", 70), ("bad", "沒有數值的工作表", 260),
        ("t", "溫度中位數(°C)", 110), ("h", "濕度中位數(%)", 110),
    ])
    self.tree_items = self._make_tree(f2, [
        ("item", "項目", 150), ("cnt", "出現次數", 80), ("mean", "平均 Rx", 100),
        ("std", "標準差", 100), ("min", "最小", 90), ("max", "最大", 90),
        ("miss", "沒有數值次數", 110), ("odd", "疑似異常（偏離平均 >3 倍標準差）", 300),
    ])
    self.tree_items.tag_configure("odd", foreground=PAL["warn"])

  def _refresh_overview(self):
    self.tree_files.delete(*self.tree_files.get_children())
    self.tree_items.delete(*self.tree_items.get_children())
    n_files = len(self.infos)
    if not n_files:
      self.lbl_over.config(text="尚無資料。請先選擇根資料夾或額外加入 Excel。")
      return
    settings = self.app.range_settings
    n_sheets = 0
    for d in sorted(self.date_map):
      for info in self.date_map[d]:
        data_sheets = [s for s in info["sheets"] if s["kind"] == "data"]
        n_sheets += len(data_sheets)
        bad = []
        for s in data_sheets:
          v = get_values(s, settings.get(range_key(info["path"], s["name"])))
          if v["rx"] is None and v["a0"] is None:
            bad.append(s["name"] + ("(範圍有問題)" if v["warn"] else ""))
        t_txt, h_txt = median_reading(info)
        self.tree_files.insert("", "end", values=(
            d, info["filename"], len(data_sheets), len(data_sheets) - len(bad),
            "、".join(bad), t_txt, h_txt,
        ))
    entries = [(i, d) for d in sorted(self.date_map) for i in self.date_map[d]]
    name_of = self._name_func()
    names = set()
    for info, _ in entries:
      for s in info["sheets"]:
        if s["kind"] == "data":
          names.add(name_of(s["name"]))
    points, _w, _c = build_points(entries, settings, name_of, names, "rx", "")
    per = {}
    per_pts = {}
    for p in points:
      per.setdefault(p["item"], []).append(p["value"])
      per_pts.setdefault(p["item"], []).append((p["date"], p["value"]))
    for item in sorted(per, key=natural_key):
      vals = [v for v in per[item] if v is not None]
      odd = ""
      if vals:
        mean = sum(vals) / len(vals)
        std = statistics.stdev(vals) if len(vals) > 1 else 0.0
        if len(vals) >= 4 and std > 0:
          bad = [
              f"{d[5:]}（{v:.3f}）" for d, v in per_pts[item]
              if v is not None and abs(v - mean) > 3 * std
          ]
          odd = f"{len(bad)} 筆：" + "、".join(bad) if bad else ""
        elif len(vals) < 4:
          odd = "數據少於 4 筆，不判斷"
        row = (item, len(per[item]), f"{mean:.3f}", f"{std:.3f}",
               f"{min(vals):.3f}", f"{max(vals):.3f}", len(per[item]) - len(vals), odd)
      else:
        row = (item, len(per[item]), "", "", "", "", len(per[item]), "")
      tags = ("odd",) if odd and not odd.startswith("數據少於") else ()
      self.tree_items.insert("", "end", values=row, tags=tags)
    lines = [f"共 {n_files} 個 Excel、{n_sheets} 個數據工作表、{len(per)} 個項目。"]
    if self.date_map:
      days = sorted(self.date_map)
      d0 = datetime.strptime(days[0], "%Y-%m-%d")
      d1 = datetime.strptime(days[-1], "%Y-%m-%d")
      total_days = (d1 - d0).days + 1
      lines.append(
          f"日期範圍：{days[0]} ～ {days[-1]}，其中 {len(days)} 天有資料、"
          f"{total_days - len(days)} 天沒有。"
      )
    if self.undated:
      lines.append(f"另有 {len(self.undated)} 個檔案沒有日期或讀取失敗（見日曆下方清單）。")
    self.lbl_over.config(text="\n".join(lines))


class ExcelProcessorApp(tk.Tk):

  def __init__(self):
    super().__init__()
    self.title("Excel 資料處理與列印整合工具")

    # 【UI 美化】依螢幕縮放比例調整視窗大小
    # base_tk_scaling＝Windows 顯示縮放換算出來的原始值；ui_scale＝程式內「介面大小」倍率
    self.base_tk_scaling = float(self.tk.call("tk", "scaling"))
    self.ui_scale = load_ui_scale()
    self.tk.call("tk", "scaling", self.base_tk_scaling * self.ui_scale)
    self.scale = max(0.6, self.winfo_fpixels("1i") / 96.0)
    self.geometry(f"{self.px(900)}x{self.px(750)}")
    self.minsize(self.px(820), self.px(650))

    # 【UI 美化】套用主題、統一字型、設定樣式
    self.theme_name = load_theme_setting()
    if self.theme_name == "tech" and not HAS_SV_TTK:
      self.theme_name = "light"  # 科技深色需要 sv-ttk
    set_palette(self.theme_name)
    self.setup_theme()
    self.setup_fonts()
    self.setup_styles()

    # 數據整理全域變數
    self.file_list = []
    self.item_mapping = {}
    self.mode_var = tk.StringVar(value="RESISTOR")

    # 共用取值：檔案資訊快取、各工作表的取值範圍設定
    self.info_cache = InfoCache()
    self.range_settings = load_range_settings()
    self.print_info = None

    self.container = ttk.Frame(self, padding=10)
    self.container.pack(fill="both", expand=True)

    # 進入點：顯示系統主頁
    self.show_main_home()

  # ==================== UI 美化輔助函式 ====================
  def px(self, n):
    return int(n * self.scale)

  def setup_theme(self):
    if HAS_SV_TTK:
      sv_ttk.set_theme("dark" if self.theme_name == "tech" else "light")
    else:
      try:
        ttk.Style().theme_use("vista")
      except tk.TclError:
        pass

  def setup_fonts(self):
    families = set(tkfont.families())
    self.base_family = None
    for cand in (
        "Microsoft JhengHei UI",
        "Microsoft JhengHei",
        "PingFang TC",
        "Noto Sans CJK TC",
        "Arial",
    ):
      if cand in families:
        self.base_family = cand
        break
    if self.base_family is None:
      self.base_family = tkfont.nametofont("TkDefaultFont").actual("family")

    for name in (
        "TkDefaultFont",
        "TkTextFont",
        "TkMenuFont",
        "TkHeadingFont",
        "TkCaptionFont",
        "SunValleyCaptionFont",
        "SunValleyBodyFont",
        "SunValleyBodyStrongFont",
        "SunValleyBodyLargeFont",
        "SunValleySubtitleFont",
        "SunValleyTitleFont",
        "SunValleyTitleLargeFont",
        "SunValleyDisplayFont",
    ):
      try:
        tkfont.nametofont(name).configure(family=self.base_family)
      except tk.TclError:
        pass

    self.font_title = (self.base_family, 20, "bold")
    self.font_h2 = (self.base_family, 14, "bold")
    self.font_list = (self.base_family, 11)

  FONT_NAMES = (
      "TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
      "TkCaptionFont", "TkSmallCaptionFont", "TkFixedFont", "TkTooltipFont",
      "SunValleyCaptionFont", "SunValleyBodyFont", "SunValleyBodyStrongFont",
      "SunValleyBodyLargeFont", "SunValleySubtitleFont", "SunValleyTitleFont",
      "SunValleyTitleLargeFont", "SunValleyDisplayFont",
  )

  def scale_pixel_fonts(self):
    """sv-ttk 的字體大小是「像素」（負數），不會跟著 Windows 縮放，所以這裡手動乘上縮放倍率。"""
    if not hasattr(self, "_font_base"):
      self._font_base = {}
    for name in self.FONT_NAMES:
      try:
        f = tkfont.nametofont(name)
        if name not in self._font_base:
          size = int(f.cget("size"))
          self._font_base[name] = size if size < 0 else None
        base = self._font_base[name]
        if base is not None:
          f.configure(size=-max(8, int(round(abs(base) * self.scale))))
      except (tk.TclError, ValueError):
        pass

  def setup_styles(self):
    self.scale_pixel_fonts()
    style = ttk.Style()
    style.configure("Treeview", rowheight=self.px(28))
    style.configure("Title.TLabel", font=self.font_title)
    style.configure("H2.TLabel", font=self.font_h2)
    style.configure("Hint.TLabel", foreground=PAL["muted"])
    style.configure("Big.TButton", font=(self.base_family, 12))

    # 【首頁放大】只給主選單頁使用，其他頁面不受影響
    style.configure(
        "HomeTitle.TLabel", font=(self.base_family, 34, "bold")
    )
    style.configure(
        "HomeSub.TLabel", font=(self.base_family, 16), foreground=PAL["muted"]
    )
    style.configure(
        "Home.Accent.TButton", font=(self.base_family, 18, "bold")
    )

    # 給舊式 tk 元件 (Listbox / Canvas / Toplevel) 用的顏色，讓它們跟主題一致
    try:
      bg = style.lookup(".", "background")
      self.theme_bg = bg if bg else ("#1c1c1c" if PAL["dark"] else "#fafafa")
    except tk.TclError:
      self.theme_bg = "#1c1c1c" if PAL["dark"] else "#fafafa"
    self.theme_fg = PAL["fg"]
    self.theme_select_bg = PAL["sel_bg"]
    try:
      self.configure(bg=self.theme_bg)
    except tk.TclError:
      pass

  def apply_ui_scale(self, value):
    """改變介面大小倍率：重設 tk 縮放（字體與元件一起變），並回到主頁重建畫面。"""
    old = self.scale
    self.ui_scale = value
    self.tk.call("tk", "scaling", self.base_tk_scaling * value)
    self.scale = max(0.6, self.winfo_fpixels("1i") / 96.0)
    r = self.scale / old if old else 1.0
    try:
      w = max(self.px(820), int(self.winfo_width() * r))
      h = max(self.px(650), int(self.winfo_height() * r))
      self.minsize(self.px(820), self.px(650))
      self.geometry(f"{w}x{h}")
    except tk.TclError:
      pass
    self.setup_styles()
    save_ui_scale(value)
    self.show_main_home()

  def toggle_theme(self):
    """切換 淺色／科技深色，並重新顯示主頁（其他頁面進入時會用新主題重建）。"""
    new = "light" if self.theme_name == "tech" else "tech"
    if new == "tech" and not HAS_SV_TTK:
      messagebox.showinfo(
          "提示",
          "科技深色主題需要 sv-ttk。\n請先在命令列執行：pip install sv-ttk\n安裝後重新開啟程式即可。",
      )
      return
    self.theme_name = new
    set_palette(new)
    if HAS_SV_TTK:
      sv_ttk.set_theme("dark" if new == "tech" else "light")
    self.setup_styles()
    save_theme_setting(new)
    self.show_main_home()

  def center_dialog(self, top, w, h):
    w, h = self.px(w), self.px(h)
    self.update_idletasks()
    x = self.winfo_rootx() + (self.winfo_width() - w) // 2
    y = self.winfo_rooty() + (self.winfo_height() - h) // 2
    top.geometry(f"{w}x{h}+{max(x, 0)}+{max(y, 0)}")
    top.configure(bg=self.theme_bg)
    top.transient(self)

  def clear_container(self):
    for widget in self.container.winfo_children():
      widget.destroy()

  # ==================== 0. 系統主頁 (選單) ====================
  def show_main_home(self):
    self.clear_container()

    center = ttk.Frame(self.container)
    center.place(relx=0.5, rely=0.42, anchor="center")

    lbl_title = ttk.Label(
        center, text="Excel 資料處理系統", style="HomeTitle.TLabel"
    )
    lbl_title.pack(pady=(0, 10))

    lbl_sub = ttk.Label(
        center, text="請選擇要使用的功能", style="HomeSub.TLabel"
    )
    lbl_sub.pack(pady=(0, 40))

    btn_frame = ttk.Frame(center)
    btn_frame.pack()

    btn_print = ttk.Button(
        btn_frame,
        text="🖨️ 數據列印",
        width=22,
        style="Home.Accent.TButton",
        command=self.show_print_mode_select,
    )
    btn_print.pack(pady=14, ipady=self.px(18))

    btn_organize = ttk.Button(
        btn_frame,
        text="📊 數據整理",
        width=22,
        style="Home.Accent.TButton",
        command=self.show_step1_home,
    )
    btn_organize.pack(pady=14, ipady=self.px(18))

    btn_view = ttk.Button(
        btn_frame,
        text="📈 檢視",
        width=22,
        style="Home.Accent.TButton",
        command=self.show_view_mode,
    )
    btn_view.pack(pady=14, ipady=self.px(18))

    opt = ttk.Frame(center)
    opt.pack(pady=(20, 0))
    ttk.Button(
        opt,
        text="🌓 切換主題（目前：" + ("科技深色" if self.theme_name == "tech" else "淺色") + "）",
        command=self.toggle_theme,
    ).pack(side="left", padx=(0, 16))
    ttk.Label(opt, text="介面大小").pack(side="left", padx=(0, 6))
    cmb_scale = ttk.Combobox(
        opt, state="readonly", width=6,
        values=[f"{int(round(c * 100))}%" for c in UI_SCALE_CHOICES],
    )
    cur = min(UI_SCALE_CHOICES, key=lambda c: abs(c - self.ui_scale))
    cmb_scale.set(f"{int(round(cur * 100))}%")
    cmb_scale.pack(side="left")
    cmb_scale.bind(
        "<<ComboboxSelected>>",
        lambda e: self.apply_ui_scale(UI_SCALE_CHOICES[cmb_scale.current()]),
    )
    ttk.Label(
        center,
        text="介面大小是在 Windows 顯示縮放之外再加的倍率，會記住設定。",
        style="Hint.TLabel",
    ).pack(pady=(6, 0))

  # ==================== 共用：讀檔、取值範圍、檢視入口 ====================
  def load_infos(self, paths, parent=None, title="讀取 Excel 資料中…"):
    """讀取（或從快取取得）多個檔案的資訊；會顯示可取消的進度視窗。取消時回傳 None。"""
    parent = parent or self
    paths = list(paths)
    infos = {}
    pending = [p for p in paths if self.info_cache.peek(p) is None]
    prog = ProgressWindow(self, parent, title, len(pending)) if pending else None
    try:
      for p in paths:
        if prog is not None and self.info_cache.peek(p) is None:
          if prog.cancelled:
            return None
          prog.set_text(os.path.basename(p))
          infos[p] = self.info_cache.get(p)
          prog.step()
        else:
          infos[p] = self.info_cache.get(p)
    finally:
      if prog is not None:
        prog.close()
      self.info_cache.save()
    return infos

  def date_text_of(self, info):
    """從檔名或資料夾名解析日期（YYYY-MM-DD），解析不到回傳空字串。"""
    d = self.parse_date_from_string(info["filename"])
    if not d:
      d = self.parse_date_from_string(
          os.path.basename(os.path.dirname(info["path"]))
      )
    return d or ""

  def open_range_dialog(self, paths):
    paths = list(paths)
    if not paths:
      messagebox.showinfo("提示", "目前沒有可以設定的檔案，請先加入 Excel 檔案！")
      return False
    infos = self.load_infos(paths, title="讀取工作表資訊中…")
    if infos is None:
      return False
    dlg = RangeDialog(self, [infos[p] for p in paths])
    self.wait_window(dlg)
    return dlg.saved

  def refresh_print_th(self):
    """依「溫濕度取法」重新填入列印頁的溫度與濕度。"""
    info = self.print_info
    temp, hum = "", ""
    if info is not None:
      if self.th_mode_var.get() == TH_MODE_MEDIAN:
        temp, hum = median_reading(info)
      else:
        temp, hum = first_reading(info)
    self.entry_temp.delete(0, tk.END)
    self.entry_humidity.delete(0, tk.END)
    if temp:
      self.entry_temp.insert(0, temp)
    if hum:
      self.entry_humidity.insert(0, hum)

  def show_view_mode(self):
    self.clear_container()
    w = min(max(self.winfo_width(), self.px(1150)), self.winfo_screenwidth() - 60)
    h = min(max(self.winfo_height(), self.px(820)), self.winfo_screenheight() - 90)
    self.geometry(f"{w}x{h}")
    self.view_page = ViewPage(self)

  # ==================== 數據列印功能區 ====================
  def show_print_mode_select(self):
    file_path = filedialog.askopenfilename(
        title="選取數據列印用的 Excel 檔案",
        filetypes=[("Excel 檔案", "*.xlsx *.xls *.xlsm"), ("所有檔案", "*.*")],
    )
    if not file_path:
      return

    filename = os.path.basename(file_path)
    folder_name = os.path.basename(os.path.dirname(file_path))

    # 解析日期
    date_str = self.parse_date_from_string(filename)
    if not date_str:
      date_str = self.parse_date_from_string(folder_name)
    if not date_str:
      user_choice = self.request_manual_date(filename, folder_name)
      if user_choice == "SKIP" or user_choice is None:
        return
      date_str = user_choice

    self.show_print_layout_page(file_path, date_str)

  def show_print_layout_page(self, file_path, date_str):
    self.clear_container()

    lbl_title = ttk.Label(
        self.container, text="數據列印設定", style="H2.TLabel"
    )
    lbl_title.pack(pady=10)

    # 1. 上方設定區域 (日期與氣象參數)
    top_frame = ttk.LabelFrame(
        self.container, text="基本資訊設定", padding=12
    )
    top_frame.pack(fill="x", padx=20, pady=5)

    # 上方左邊：日期
    ttk.Label(top_frame, text="日期 (YYYY MM DD):").grid(
        row=0, column=0, sticky="w", padx=5, pady=5
    )
    self.entry_print_date = ttk.Entry(top_frame, width=15)
    self.entry_print_date.grid(row=0, column=1, sticky="w", padx=5, pady=5)
    formatted_date_input = date_str.replace("-", " ")
    self.entry_print_date.insert(0, formatted_date_input)

    # 上方右邊：溫度、濕度、天氣狀態
    ttk.Label(top_frame, text="溫度 (°C):").grid(
        row=0, column=2, sticky="w", padx=(20, 5), pady=5
    )
    self.entry_temp = ttk.Entry(top_frame, width=10)
    self.entry_temp.grid(row=0, column=3, sticky="w", padx=5, pady=5)

    ttk.Label(top_frame, text="濕度 (%):").grid(
        row=0, column=4, sticky="w", padx=(10, 5), pady=5
    )
    self.entry_humidity = ttk.Entry(top_frame, width=10)
    self.entry_humidity.grid(row=0, column=5, sticky="w", padx=5, pady=5)

    ttk.Label(top_frame, text="天氣狀態:").grid(
        row=0, column=6, sticky="w", padx=(10, 5), pady=5
    )
    self.entry_weather = ttk.Entry(top_frame, width=12)
    self.entry_weather.grid(row=0, column=7, sticky="w", padx=5, pady=5)
    self.entry_weather.insert(0, "晴")

    # 讀取 Excel（共用取值層）：工作表分類、溫濕度
    self.print_file_path = file_path
    sheet_names = []
    self.print_info = None
    infos = self.load_infos([file_path], title="讀取 Excel 檔案中…")
    if infos is not None:
      info = infos[file_path]
      if info.get("error"):
        messagebox.showerror("錯誤", f"讀取 Excel 檔案失敗：\n{info['error']}")
      else:
        self.print_info = info
        sheet_names = [s["name"] for s in info["sheets"]]

    # 溫濕度取法（預設維持原本做法：開頭第一筆）
    ttk.Label(top_frame, text="溫濕度取法:").grid(
        row=1, column=0, sticky="w", padx=5, pady=5
    )
    self.th_mode_var = tk.StringVar(value=TH_MODE_FIRST)
    cmb_th = ttk.Combobox(
        top_frame,
        textvariable=self.th_mode_var,
        values=[TH_MODE_FIRST, TH_MODE_MEDIAN],
        state="readonly",
        width=30,
    )
    cmb_th.grid(row=1, column=1, columnspan=5, sticky="w", padx=5, pady=5)
    cmb_th.bind("<<ComboboxSelected>>", lambda e: self.refresh_print_th())
    self.refresh_print_th()

    # 2. 中央列表：項目勾選清單
    list_frame = ttk.LabelFrame(
        self.container, text="選取列印項目 (工作表)", padding=10
    )
    list_frame.pack(fill="both", expand=True, padx=20, pady=10)

    canvas = tk.Canvas(
        list_frame, bg=self.theme_bg, highlightthickness=0, borderwidth=0
    )
    scrollbar = ttk.Scrollbar(
        list_frame, orient="vertical", command=canvas.yview
    )
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    self.print_sheet_vars = {}
    sheet_by_name = (
        {s["name"]: s for s in self.print_info["sheets"]}
        if self.print_info
        else {}
    )
    for sname in sheet_names:
      sh = sheet_by_name.get(sname)
      is_check = (
          sname not in ["氣溫", "溫度"]
          and "plxdaq" not in sname.lower()
          and sh is not None
          and sh["kind"] not in ("temp", "plxdaq")
          and not sh["hidden"]
          and not (sh["kind"] == "data" and not sheet_has_values(sh))
      )
      label = sname
      if sh is not None and sh["kind"] == "temp":
        label += "（溫濕度紀錄）"
      elif sh is not None and sh["kind"] == "data" and not sheet_has_values(sh):
        label += "（沒有數值）"
      var = tk.BooleanVar(value=is_check)
      chk = ttk.Checkbutton(scroll_frame, text=label, variable=var)
      chk.pack(anchor="w", pady=2)
      self.print_sheet_vars[sname] = var

    # 3. 下方控制按鈕
    bottom_frame = ttk.Frame(self.container)
    bottom_frame.pack(fill="x", padx=20, pady=10)

    btn_back = ttk.Button(
        bottom_frame, text="返回主選單", command=self.show_main_home
    )
    btn_back.pack(side="left")

    btn_range = ttk.Button(
        bottom_frame,
        text="⚙ 取值範圍",
        command=lambda: self.open_range_dialog([self.print_file_path]),
    )
    btn_range.pack(side="left", padx=8)

    btn_confirm = ttk.Button(
        bottom_frame,
        text="確定 (生成列印用 Excel)",
        style="Accent.TButton",
        command=self.generate_print_excel,
    )
    btn_confirm.pack(side="right")

  def generate_print_excel(self):
    date_val = self.entry_print_date.get().strip()
    match = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", date_val)
    if not match:
      messagebox.showerror(
          "格式錯誤", "日期格式不正確，請輸入 YYYY MM DD（例如：2026 10 15）"
      )
      return

    year, month, day = match.groups()
    temp_val = self.entry_temp.get().strip()
    try:
      temp_val = str(int(float(temp_val)))
    except ValueError:
      pass

    humidity_val = self.entry_humidity.get().strip()
    weather_val = self.entry_weather.get().strip()

    selected_sheets = [
        s for s, var in self.print_sheet_vars.items() if var.get()
    ]
    if not selected_sheets:
      messagebox.showerror("錯誤", "請至少勾選一個項目！")
      return

    try:
      info = self.info_cache.get(self.print_file_path)
      if info.get("error"):
        raise RuntimeError(info["error"])
      by_name = {s["name"]: s for s in info["sheets"]}
      export_rows = []
      range_warns = []

      for sname in selected_sheets:
        rx_val = ""
        a0_val = ""
        sh = by_name.get(sname)
        if sh is not None:
          v = get_values(
              sh, self.range_settings.get(range_key(self.print_file_path, sname))
          )
          if v["rx"] is not None:
            rx_val = v["rx"]
          if v["a0"] is not None:
            a0_val = v["a0"]
          for w in v["warn"]:
            range_warns.append(f"{sname}：{w}")
        export_rows.append((sname, rx_val, a0_val))

      out_wb = openpyxl.Workbook()
      out_sheet = out_wb.active
      out_sheet.title = "數據列印"

      # 自動計算星期幾（只顯示數字，例如：一）
      dt_obj = datetime(int(year), int(month), int(day))
      weekdays_zh = ["一", "二", "三", "四", "五", "六", "日"]
      weekday_str = weekdays_zh[dt_obj.weekday()]

      # 填入 A1 整合資訊
      out_sheet["A1"] = (
          f"{year}年 {int(month)}月 {int(day)}日 ({weekday_str})"
          f" {temp_val}℃ {humidity_val}% {weather_val}"
      )

      # 填入表格標頭 (從 A3 開始)
      out_sheet["A3"] = "項目"
      out_sheet["B3"] = "Rx"
      out_sheet["C3"] = "A0"

      current_row = 4
      for name, rx, a0 in export_rows:
        out_sheet[f"A{current_row}"] = name
        out_sheet[f"B{current_row}"] = rx
        out_sheet[f"C{current_row}"] = a0
        current_row += 1

      # 字體大小改為 10
      for row in out_sheet.iter_rows():
        for cell in row:
          current_font = cell.font
          bold = current_font.bold if current_font else False
          italic = current_font.italic if current_font else False
          cell.font = Font(name="Arial", size=10, bold=bold, italic=italic)

      # 替項目數據表格加入所有框線 (從第 3 行到最後一列，A 欄到 C 欄)
      thin_side = Side(style="thin", color="000000")
      table_border = Border(
          left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
      )
      for r in range(3, current_row):
        for c in range(1, 4):
          out_sheet.cell(row=r, column=c).border = table_border

      # 將所有縱列寬度一律設定為 19.38
      for col in out_sheet.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        out_sheet.column_dimensions[col_letter].width = 19.38

      save_path = filedialog.asksaveasfilename(
          defaultextension=".xlsx",
          filetypes=[("Excel 檔案", "*.xlsx")],
          title="選擇儲存結果位置與檔名",
      )

      if not save_path:
        return

      out_wb.save(save_path)
      messagebox.showinfo("成功", f"列印用 Excel 檔案已成功生成至：\n{save_path}")
      if range_warns:
        more = (
            f"\n...以及其他 {len(range_warns) - 12} 筆" if len(range_warns) > 12 else ""
        )
        messagebox.showwarning(
            "取值範圍警告",
            "以下項目的範圍有問題，數值已留空：\n\n"
            + "\n".join(range_warns[:12])
            + more,
        )

      # 自動用 Excel 開啟剛剛儲存的檔案 (Windows 專用，視窗保持開啟)
      try:
        os.startfile(save_path)
      except Exception as e:
        print(f"自動開啟檔案失敗: {e}")

    except Exception as e:
      messagebox.showerror("重大錯誤", f"生成列印檔案時發生錯誤：\n{e}")

  # ==================== 1. 數據整理 - 步驟一：選擇檔案 (Home) ====================
  def show_step1_home(self):
    self.clear_container()

    lbl_title = ttk.Label(
        self.container,
        text="【數據整理】步驟一：選擇模式與管理 Excel 檔案",
        style="H2.TLabel",
    )
    lbl_title.pack(pady=10)

    mode_frame = ttk.LabelFrame(
        self.container, text="搜尋模式設定", padding=12
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

    btn_frame = ttk.Frame(self.container)
    btn_frame.pack(fill="x", padx=20, pady=8)

    btn_add_files = ttk.Button(
        btn_frame, text="加入 Excel 檔案", command=self.add_excel_files
    )
    btn_add_files.pack(side="left", padx=5)

    btn_add_folder = ttk.Button(
        btn_frame, text="從資料夾遞迴加入", command=self.add_excel_folder
    )
    btn_add_folder.pack(side="left", padx=5)

    list_frame = ttk.Frame(self.container)
    list_frame.pack(fill="both", expand=True, padx=20, pady=10)

    self.tree_step1 = ttk.Treeview(
        list_frame, columns=("Filename", "Date", "Path"), show="headings"
    )
    self.tree_step1.heading("Filename", text="檔案名稱")
    self.tree_step1.heading("Date", text="解析日期 (YYYY-MM-DD)")
    self.tree_step1.heading("Path", text="完整路徑")
    self.tree_step1.column("Filename", width=self.px(220))
    self.tree_step1.column("Date", width=self.px(150))
    self.tree_step1.column("Path", width=self.px(380))
    self.tree_step1.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        list_frame, orient="vertical", command=self.tree_step1.yview
    )
    scrollbar.pack(side="right", fill="y")
    self.tree_step1.configure(yscrollcommand=scrollbar.set)

    bottom_frame = ttk.Frame(self.container)
    bottom_frame.pack(fill="x", padx=20, pady=10)

    btn_home = ttk.Button(
        bottom_frame, text="返回主選單", command=self.show_main_home
    )
    btn_home.pack(side="left", padx=5)

    btn_remove = ttk.Button(
        bottom_frame, text="移除選取檔案 (X)", command=self.remove_selected_file
    )
    btn_remove.pack(side="left", padx=5)

    btn_range = ttk.Button(
        bottom_frame,
        text="⚙ 取值範圍設定",
        command=lambda: self.open_range_dialog(
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

  def parse_date_from_string(self, text):
    match = re.search(r"(\d{4})_(\d{1,2})_(\d{1,2})", text)
    if match:
      y, m, d = match.groups()
      try:
        dt = datetime(int(y), int(m), int(d))
        return dt.strftime("%Y-%m-%d")
      except ValueError:
        pass
    return None

  def request_manual_date(self, filename, folder_name):
    top = tk.Toplevel(self)
    top.title(f"無法辨識日期: {filename}")
    self.center_dialog(top, 460, 220)
    top.grab_set()

    body = ttk.Frame(top, padding=15)
    body.pack(fill="both", expand=True)

    ttk.Label(
        body,
        text=(
            f"無法從檔名 '{filename}' 或資料夾 '{folder_name}' 辨識日期。\n"
            "請手動輸入正確日期（格式：YYYY MM DD），或點擊「不加入」略過："
        ),
        wraplength=self.px(420),
    ).pack(anchor="w")

    entry_date = ttk.Entry(body, width=25)
    entry_date.pack(pady=10)
    entry_date.insert(0, "2026 01 01")

    result_date = [None]

    def on_submit():
      val = entry_date.get().strip()
      match = re.match(r"^(\d{4})\s+(\d{1,2})\s+(\d{1,2})$", val)
      if match:
        y, m, d = match.groups()
        try:
          dt = datetime(int(y), int(m), int(d))
          result_date[0] = dt.strftime("%Y-%m-%d")
          top.destroy()
          return
        except ValueError:
          pass
      messagebox.showerror(
          "格式錯誤",
          "日期格式不正確，請重新輸入（例如：2026 10 15）",
          parent=top,
      )

    def on_cancel():
      result_date[0] = "SKIP"
      top.destroy()

    btn_frame = ttk.Frame(body)
    btn_frame.pack(pady=10)

    ttk.Button(
        btn_frame, text="確定加入", style="Accent.TButton", command=on_submit
    ).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="不加入", command=on_cancel).pack(
        side="left", padx=5
    )

    self.wait_window(top)
    return result_date[0]

  def add_files_logic(self, file_paths):
    for path in file_paths:
      if not path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        continue
      filename = os.path.basename(path)
      folder_name = os.path.basename(os.path.dirname(path))

      date_str = self.parse_date_from_string(filename)
      if not date_str:
        date_str = self.parse_date_from_string(folder_name)
      if not date_str:
        user_choice = self.request_manual_date(filename, folder_name)
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

  # ==================== 2. 數據整理 - 步驟二：項目管理 (Management) ====================
  def show_step2_management(self):
    self.clear_container()

    lbl_title = ttk.Label(
        self.container,
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

    setting_file_path = os.path.join(os.getcwd(), "setting.json")
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
        self.container,
        text=(
            "在此可調整項目上下順序、合併項目，或刪除不需要的項目。\n"
            "(提示：設定會自動儲存至 setting.json)"
        ),
        style="Hint.TLabel",
    )
    lbl_info.pack(anchor="w", padx=20, pady=5)

    main_mgmt_frame = ttk.Frame(self.container)
    main_mgmt_frame.pack(fill="both", expand=True, padx=20, pady=5)

    # 【UI 美化】Listbox 是舊式 tk 元件，手動設定顏色與邊框讓它跟主題一致
    self.listbox_items = tk.Listbox(
        main_mgmt_frame,
        selectmode=tk.EXTENDED,
        font=self.font_list,
        bg=PAL["card"],
        fg=self.theme_fg,
        selectbackground=self.theme_select_bg,
        selectforeground=PAL["sel_fg"],
        activestyle="none",
        relief="flat",
        borderwidth=0,
        highlightthickness=1,
        highlightbackground=PAL["card_border"],
        highlightcolor=self.theme_select_bg,
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

    bottom_frame = ttk.Frame(self.container)
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

    top = tk.Toplevel(self)
    top.title("合併項目與選擇代表名稱")
    self.center_dialog(top, 460, 260)
    top.grab_set()

    body = ttk.Frame(top, padding=15)
    body.pack(fill="both", expand=True)

    ttk.Label(
        body,
        text=(
            f"準備合併以下項目：\n{', '.join(selected_names)}\n\n"
            "請選擇其中一個作為最終的代表名稱："
        ),
        wraplength=self.px(420),
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

  # ==================== 3. 數據整理 - 步驟三：計算與矩陣輸出 (Output) ====================
  def show_step3_output(self):
    self.clear_container()

    lbl_title = ttk.Label(
        self.container,
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
        self.container,
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
        self.container, text="選擇匯出檔案格式", padding=12
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
        self.container,
        text="開始計算並儲存",
        style="Accent.TButton",
        command=self.process_and_export,
    )
    btn_run.pack(pady=10, ipady=4)

    btn_back_frame = ttk.Frame(self.container)
    btn_back = ttk.Button(
        btn_back_frame, text="上一步", command=self.show_step2_management
    )
    btn_back.pack()
    btn_back_frame.pack(pady=5)

  def process_and_export(self):
    try:
      setting_file_path = os.path.join(os.getcwd(), "setting.json")
      try:
        with open(setting_file_path, "w", encoding="utf-8") as sf:
          json.dump(self.item_mapping, sf, ensure_ascii=False, indent=4)
      except Exception as se:
        print(f"儲存 setting.json 失敗: {se}")

      matrix_data = {}
      source = "rx" if self.mode_var.get() == "RESISTOR" else "a0"
      conflict_reports = []
      range_warns = []

      infos = self.load_infos(
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
            v = get_values(sh, self.range_settings.get(range_key(f_path, sname)))
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
          if hasattr(self, "export_format_var")
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


if __name__ == "__main__":
  # 【UI 美化】必須在建立視窗之前呼叫，高 DPI 設定才會生效
  enable_high_dpi()
  app = ExcelProcessorApp()
  app.mainloop()