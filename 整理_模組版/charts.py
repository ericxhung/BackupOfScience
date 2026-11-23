"""圖表繪製：matplotlib（視窗內）與 plotly（瀏覽器互動圖）。"""

import importlib.util
import os
import re
import math
from datetime import datetime
from collections import OrderedDict, defaultdict

from core import PAL, natural_key


# ---------- 大套件改成「用到才載入」，讓程式啟動更快 ----------
# plotly、matplotlib 載入各要一兩秒，所以不在啟動時載入：
#   HAS_PLOTLY：只檢查有沒有安裝（很快），真的畫圖時才會載入 plotly。
#   ensure_mpl()：第一次需要 matplotlib（進入「檢視」模式）時才載入。
HAS_PLOTLY = importlib.util.find_spec("plotly") is not None


class _LazyGo:
  """第一次用到 go.Figure 之類的東西時，才載入 plotly。"""

  def __getattr__(self, name):
    import plotly.graph_objects as _go

    return getattr(_go, name)


go = _LazyGo()

matplotlib = mdates = mpl_font_manager = Figure = None
FigureCanvasTkAgg = NavigationToolbar2Tk = None
HAS_MPL = importlib.util.find_spec("matplotlib") is not None
HAS_MPL_TK = HAS_MPL  # 載入之後會更新成實際結果
_mpl_loaded = False


def ensure_mpl():
  """載入 matplotlib（只會做一次）。成功回傳 True，沒安裝回傳 False。"""
  global matplotlib, mdates, mpl_font_manager, Figure
  global FigureCanvasTkAgg, NavigationToolbar2Tk, HAS_MPL, HAS_MPL_TK, _mpl_loaded
  if _mpl_loaded:
    return HAS_MPL
  _mpl_loaded = True
  try:
    import matplotlib as _m
    import matplotlib.dates as _d
    from matplotlib import font_manager as _f
    from matplotlib.figure import Figure as _Fig

    matplotlib, mdates, mpl_font_manager, Figure = _m, _d, _f, _Fig
    HAS_MPL = True
  except Exception:
    HAS_MPL = HAS_MPL_TK = False
    return False
  try:
    from matplotlib.backends.backend_tkagg import (
        FigureCanvasTkAgg as _Canvas,
        NavigationToolbar2Tk as _Toolbar,
    )

    FigureCanvasTkAgg, NavigationToolbar2Tk = _Canvas, _Toolbar
    HAS_MPL_TK = True
  except Exception:
    HAS_MPL_TK = False
  return True


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
  if not ensure_mpl():
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
