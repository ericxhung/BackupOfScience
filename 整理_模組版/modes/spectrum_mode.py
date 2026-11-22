# -*- coding: utf-8 -*-
"""光譜校正模式（modes/spectrum_mode.py）
把「光譜校正工具」做成主程式的一個模式：主頁會多一個「🌈 光譜校正」按鈕。
流程：1 裁切（一條線） -> 2 對圖（樣本當底圖） -> 3 比對與校正 -> 4 輸出
必要套件：numpy、matplotlib、Pillow　選用套件：pillow-heif（直接讀 iPad 的 HEIC）
視窗、主題、介面大小、高 DPI 都由 main.py 處理；本檔只負責光譜校正的畫面與運算。
設定檔：SpectrumSetting.json（放在主程式資料夾，不能寫入時改放使用者資料夾）
"""
import os
import json
import csv
import math
import numpy as np
from PIL import Image, ImageOps

try:
    from PIL import ImageTk
    IMGTK_OK = True
except Exception:
    IMGTK_OK = False

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_OK = True
except Exception:
    HEIF_OK = False

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from modes.base import Mode

_CanvasBase = tk.Canvas

RS = getattr(Image, "Resampling", Image)

# 從你傳的氖燈參考圖抄下來的標示值（請對照確認，可在第 2 頁直接修改）
DEFAULT_REFS = ("534.04, 540.47, 576.36, 585.14, 587.99, 594.30, 597.27, 602.71, 606.92, 609.39, 613.97, "
                "621.39, 626.22, 630.43, 633.27, 640.32, 650.84, 653.69, 660.62, 668.66, 672.49")
IMG_TYPES = [("圖片", "*.jpg *.jpeg *.png *.heic *.heif *.tif *.tiff *.bmp"), ("全部檔案", "*.*")]


def load_image(path):
    """讀圖，依 EXIF 自動轉正（iPad 直拍的方向），回傳 uint8 的 H x W x 3"""
    im = Image.open(path)
    im = ImageOps.exif_transpose(im).convert("RGB")
    return np.asarray(im, dtype=np.uint8)


def sample_bilinear(img, xs, ys):
    """雙線性內插取樣。座標以「像素中心＝整數」計。xs、ys 同形狀；超出圖片的地方補 0"""
    h, w = img.shape[:2]
    x0 = np.floor(xs).astype(np.int64)
    y0 = np.floor(ys).astype(np.int64)
    fx = (xs - x0)[..., None].astype(np.float32)
    fy = (ys - y0)[..., None].astype(np.float32)

    def g(yy, xx):
        ok = (xx >= 0) & (xx < w) & (yy >= 0) & (yy < h)
        out = img[np.clip(yy, 0, h - 1), np.clip(xx, 0, w - 1)].astype(np.float32)
        out[~ok] = 0
        return out

    return (g(y0, x0) * (1 - fx) * (1 - fy) + g(y0, x0 + 1) * fx * (1 - fy)
            + g(y0 + 1, x0) * (1 - fx) * fy + g(y0 + 1, x0 + 1) * fx * fy)


def extract_strip(img, p0, p1, width, shear_deg=0.0):
    """沿 p0->p1 這條線切出「條帶」，並轉成水平。
    回傳 nv x nu x 3；橫向每格剛好 1 px，所以欄位編號 = 沿線距離(px)。width=1 就是單一條線。
    shear_deg：亮線（狹縫像）相對於「垂直於裁切線」的傾斜角，用來把歪的亮線拉直。"""
    p0 = np.asarray(p0, float)
    p1 = np.asarray(p1, float)
    d = p1 - p0
    L = float(np.hypot(d[0], d[1]))
    e = d / L
    n = np.array([-e[1], e[0]])
    t = math.tan(math.radians(shear_deg))
    u = np.arange(0.0, L + 1e-9, 1.0)[None, :]
    nv = max(1, int(round(width)))
    v = (np.arange(nv) - (nv - 1) / 2.0)[:, None]
    uu = u + v * t
    xs = p0[0] + uu * e[0] + v * n[0]
    ys = p0[1] + uu * e[1] + v * n[1]
    return sample_bilinear(img, xs, ys)


def strip_to_profile(strip, mode="ImageJ(RGB平均)", degamma=False):
    """條帶 -> 一維亮度曲線（沿線方向）。
    預設「ImageJ(RGB平均)」＝ (R+G+B)/3，單位是灰階值 0~255，和 ImageJ 預設的 Plot Profile 一樣；
    條帶寬度內的像素會取平均（和 ImageJ 的線寬平均一樣）。"""
    s = strip
    if degamma:
        s = 255.0 * (np.clip(s, 0, 255) / 255.0) ** 2.2
    if mode == "最大通道":
        g = s.max(axis=2)
    elif mode == "RGB總和":
        g = s.sum(axis=2)
    elif mode in ("亮度", "亮度(加權)"):
        g = s @ np.array([0.299, 0.587, 0.114], dtype=np.float32)
    else:
        g = s.mean(axis=2)
    return g.mean(axis=0)


def gauss_smooth(y, sigma):
    y = np.asarray(y, float)
    if sigma <= 0:
        return y.copy()
    r = int(math.ceil(sigma * 4))
    k = np.exp(-0.5 * (np.arange(-r, r + 1) / sigma) ** 2)
    k /= k.sum()
    return np.convolve(np.pad(y, (r, r), mode="edge"), k, mode="valid")


def blur2d(a, sigma):
    """二維高斯模糊（可分離）"""
    a = np.asarray(a, np.float32)
    r = int(math.ceil(sigma * 3))
    k = np.exp(-0.5 * (np.arange(-r, r + 1) / sigma) ** 2).astype(np.float32)
    k /= k.sum()
    for ax in (0, 1):
        p = np.pad(a, [(r, r) if i == ax else (0, 0) for i in range(2)], mode="edge")
        out = np.zeros_like(a)
        n = a.shape[ax]
        for i, kv in enumerate(k):
            out += kv * (p[i:i + n] if ax == 0 else p[:, i:i + n])
        a = out
    return a


def find_peaks(y, sigma=1.5, thr_frac=0.02, min_dist=3, half=4):
    """找亮線中心。回傳 (位置 px（含小數）, 高度)。thr_frac：最低高度占最高峰的比例"""
    y = np.asarray(y, float)
    ys = gauss_smooth(y, sigma)
    z = ys - np.percentile(ys, 5)
    top = z.max() if len(z) else 0
    if top <= 0:
        return np.array([]), np.array([])
    cand = [i for i in range(1, len(z) - 1)
            if z[i] >= z[i - 1] and z[i] > z[i + 1] and z[i] >= thr_frac * top]
    cand.sort(key=lambda i: -z[i])
    kept = []
    for i in cand:
        if all(abs(i - j) >= min_dist for j in kept):
            kept.append(i)
    kept.sort()
    xs, hs = [], []
    for i in kept:
        a, b = max(0, i - half), min(len(z), i + half + 1)
        seg = z[a:b] - z[a:b].min()
        xs.append(float((np.arange(a, b) * seg).sum() / seg.sum()) if seg.sum() > 0 else float(i))
        hs.append(float(z[i]))
    return np.array(xs), np.array(hs)


def _dilate(m, r):
    d = m.copy()
    for _ in range(r):
        n = d.copy()
        n[1:] |= d[:-1]
        n[:-1] |= d[1:]
        n[:, 1:] |= d[:, :-1]
        n[:, :-1] |= d[:, 1:]
        d = n
    return d


def _label(m):
    """連通區域標記（4 鄰接）"""
    gh, gw = m.shape
    lab = np.zeros(m.shape, np.int32)
    n = 0
    ys, xs = np.nonzero(m)
    for y0, x0 in zip(ys.tolist(), xs.tolist()):
        if lab[y0, x0]:
            continue
        n += 1
        lab[y0, x0] = n
        st = [(y0, x0)]
        while st:
            y, x = st.pop()
            for yy, xx in ((y - 1, x), (y + 1, x), (y, x - 1), (y, x + 1)):
                if 0 <= yy < gh and 0 <= xx < gw and m[yy, xx] and not lab[yy, xx]:
                    lab[yy, xx] = n
                    st.append((yy, xx))
    return lab, n


def auto_line(img):
    """自動找光譜位置與色散方向（抗雜訊版）。回傳 (起點, 終點, 亮線橫向寬度)。
    1) 縮成粗格子（把單點雜訊平均掉）→ 2) 挑最亮的一團 → 3) 用「梯度方向」找色散軸
       （亮線是細長的，亮度沿著色散方向變化最大）→ 4) 沿色散軸量出範圍"""
    h, w = img.shape[:2]
    cell = max(4, max(h, w) // 300)
    gh, gw = h // cell, w // cell
    blk = img[:gh * cell, :gw * cell].astype(np.float32).reshape(gh, cell, gw, cell, 3).mean(axis=(1, 3))
    lum = blk.max(axis=2)
    bg = float(np.median(lum))
    sig = 1.4826 * float(np.median(np.abs(lum - bg))) + 0.5
    top = float(lum.max())
    thr = bg + max(0.12 * (top - bg), 6 * sig)
    mask = lum >= thr
    if mask.sum() < 3:
        raise ValueError("找不到足夠的亮點。請改用手動畫線，或調整曝光後重拍。")
    lab, n = _label(_dilate(mask, max(2, int(0.045 * max(h, w) / cell))))
    best, bs = 0, -1.0
    for i in range(1, n + 1):
        s = float(((lum - bg) * (mask & (lab == i))).sum())
        if s > bs:
            best, bs = i, s
    sel = mask & (lab == best)
    cy, cx = np.nonzero(sel)
    wgt = lum[sel] - bg
    pts = np.column_stack([(cx + 0.5) * cell, (cy + 0.5) * cell]).astype(np.float64)
    c = (pts * wgt[:, None]).sum(axis=0) / wgt.sum()

    # 色散方向：結構張量（全解析度、只看這一團附近）
    m = 2 * cell
    x0, x1 = max(0, int(cx.min() * cell - m)), min(w, int((cx.max() + 1) * cell + m))
    y0, y1 = max(0, int(cy.min() * cell - m)), min(h, int((cy.max() + 1) * cell + m))
    b = blur2d(img[y0:y1, x0:x1].max(axis=2).astype(np.float32), 1.2)
    gy, gx = np.gradient(b)
    jxx, jyy, jxy = float((gx * gx).sum()), float((gy * gy).sum()), float((gx * gy).sum())
    th = 0.5 * math.atan2(2 * jxy, jxx - jyy)
    e = np.array([math.cos(th), math.sin(th)])
    # 連續光譜（白光、太陽光）是一整條色帶：它的「梯度方向」會指向色帶的橫向，結構張量就會判錯。
    # 所以先看這團亮光本身的形狀：夠細長就直接用它的長軸（主成分），不夠細長才用上面的梯度方向。
    dp = pts - c
    cov = (dp * wgt[:, None]).T @ dp / wgt.sum()
    ev, evec = np.linalg.eigh(cov)
    if ev[1] > 2.5 * max(ev[0], 1e-9):
        e = evec[:, 1] / (np.hypot(*evec[:, 1]) or 1e-9)
    nrm = np.array([-e[1], e[0]])
    t = (pts - c) @ e
    q = (pts - c) @ nrm
    lo, hi = float(t.min()) - cell / 2.0, float(t.max()) + cell / 2.0
    wid = float(np.percentile(q, 97.5) - np.percentile(q, 2.5)) + cell
    # 延伸：藍紫端的光通常比紅綠暗很多，第一次的門檻會把它切掉。
    # 改用很低的門檻，沿著色散軸往兩端找「接得上」的暗光（只看線兩側 wid 範圍內，避免被雜訊帶偏）
    thr2 = bg + max(0.025 * (top - bg), 4 * sig)
    gy2, gx2 = np.nonzero(lum >= thr2)
    allp = np.column_stack([(gx2 + 0.5) * cell, (gy2 + 0.5) * cell]).astype(np.float64)
    if len(allp):
        tq = (allp - c) @ e
        qq = (allp - c) @ nrm
        keep = np.abs(qq) <= max(wid, 3 * cell)
        ts = np.sort(tq[keep])
        gap = max(0.1 * (hi - lo), 4 * cell)       # 兩塊亮處相距不超過這個距離才算接得上
        cur_hi, cur_lo = hi, lo
        for v in ts[ts > hi]:
            if v - cur_hi <= gap:
                cur_hi = v
            else:
                break
        for v in ts[ts < lo][::-1]:
            if cur_lo - v <= gap:
                cur_lo = v
            else:
                break
        lo, hi = cur_lo - cell / 2.0, cur_hi + cell / 2.0
    pad = 0.03 * (hi - lo) + cell / 2.0
    tl, th_ = lo - pad, hi + pad
    for ax, size in ((0, w), (1, h)):          # 不要超出照片邊界
        if abs(e[ax]) > 1e-9:
            t1, t2 = (0.0 - c[ax]) / e[ax], ((size - 1.0) - c[ax]) / e[ax]
            tl, th_ = max(tl, min(t1, t2)), min(th_, max(t1, t2))
    return c + e * tl, c + e * th_, wid


def auto_line_multi(imgs):
    """多張照片（氖燈＋其他光譜）一起找那條共用線。
    每張各自偵測，再合併：方向取加權平均、橫向位置取中位數、長度涵蓋所有光譜的範圍。
    回傳 (起點, 終點, 寬度, 每張的結果清單；偵測失敗的那張是 None)"""
    res = []
    for im in imgs:
        try:
            res.append(auto_line(im))
        except ValueError:
            res.append(None)
    ok = [r for r in res if r is not None]
    if not ok:
        raise ValueError("找不到足夠的亮點。請改用手動畫線，或調整曝光後重拍。")
    if len(ok) == 1:
        return ok[0][0], ok[0][1], ok[0][2], res
    d0 = (ok[0][1] - ok[0][0])
    d0 = d0 / (np.hypot(*d0) or 1e-9)
    e = np.zeros(2)
    for a, b, _ in ok:
        d = b - a
        L = float(np.hypot(*d)) or 1e-9
        d = d / L
        e += (L if d @ d0 >= 0 else -L) * d
    e = e / (np.hypot(*e) or 1e-9)
    nrm = np.array([-e[1], e[0]])
    q = float(np.median([((a + b) / 2.0) @ nrm for a, b, _ in ok]))
    ts = [float(pt @ e) for a, b, _ in ok for pt in (a, b)]
    wid = max(w for _, _, w in ok)
    return q * nrm + e * min(ts), q * nrm + e * max(ts), wid, res


def auto_flip(strip):
    """短波長端偏藍綠、長波長端偏紅（氖燈只有綠～紅，其他光譜可能有藍紫）。
    若藍綠在後面，回傳 True（要把起點終點對調）"""
    r = strip[..., 0].mean(axis=0)
    g = (strip[..., 1].mean(axis=0) + strip[..., 2].mean(axis=0)) / 2.0
    x = np.arange(r.size, dtype=float)
    gw = np.clip(g - r, 0, None)
    rw = np.clip(r - g, 0, None)
    if gw.sum() < 1e-6 or rw.sum() < 1e-6:
        return False
    return bool((gw * x).sum() / gw.sum() > (rw * x).sum() / rw.sum())


def auto_shear(img, p0, p1, width, rng=10.0, step=0.5):
    """試很多個傾斜角，挑「亮線最銳利」的那個。寬度小於 3 時沒有意義，直接回傳 0"""
    width = min(float(width), 61.0)
    if width < 3:
        return 0.0

    def score(s):
        pr = extract_strip(img, p0, p1, width, s).max(axis=2).mean(axis=0)
        hp = pr - gauss_smooth(pr, 6)
        return float(np.sum(hp * hp))

    best_s, best_v = 0.0, -1.0
    for s in np.arange(-rng, rng + 1e-9, step):
        v = score(s)
        if v > best_v:
            best_s, best_v = float(s), v
    for s in np.arange(best_s - step, best_s + step + 1e-9, 0.1):
        v = score(s)
        if v > best_v:
            best_s, best_v = float(s), v
    return best_s


def coarse_align(xs, hs, refs, L):
    """自動粗對：找「每個亮線位置 -> 波長」的線性關係，使亮線盡量落在參考波長上。
    只是給你一個起點，請務必用眼睛確認。"""
    xs = np.asarray(xs, float)
    hs = np.asarray(hs, float)
    refs = np.sort(np.asarray(refs, float))
    if len(xs) < 3 or len(refs) < 3:
        return None
    w = np.sqrt(hs / hs.max())

    def score(l0, k, sig):
        lam = l0[:, None] + k * xs[None, :]
        j = np.clip(np.searchsorted(refs, lam), 1, len(refs) - 1)
        d = np.minimum(np.abs(lam - refs[j - 1]), np.abs(lam - refs[j]))
        return (np.exp(-(d / sig) ** 2) * w).sum(axis=1)

    l0s = np.arange(300.0, 800.0, 0.5)
    best = (-1.0, 0.0, 0.0)
    for k in np.geomspace(0.1, 2.0, 500):
        s = score(l0s, k, 1.5)
        i = int(s.argmax())
        if s[i] > best[0]:
            best = (float(s[i]), float(l0s[i]), float(k))
    _, l0, k = best
    best = (-1.0, l0, k)
    for kk in k * np.linspace(0.97, 1.03, 61):
        ls = l0 + np.arange(-4, 4.001, 0.1)
        s = score(ls, kk, 0.6)
        i = int(s.argmax())
        if s[i] > best[0]:
            best = (float(s[i]), float(ls[i]), float(kk))
    sc, l0, k = best
    return {"lam0": l0, "k": k, "score": sc / float(w.sum()), "center": l0 + k * L / 2.0, "span": k * L}


def match_peaks(lams, refs, tol):
    """每個偵測峰配到最近的參考波長（誤差 <= tol，且一個參考值只能用一次）。回傳 refs 索引或 -1"""
    pairs = sorted((abs(l - r), i, j) for i, l in enumerate(lams) for j, r in enumerate(refs) if abs(l - r) <= tol)
    used_i, used_j, out = set(), set(), [-1] * len(lams)
    for d, i, j in pairs:
        if i in used_i or j in used_j:
            continue
        out[i] = j
        used_i.add(i)
        used_j.add(j)
    return out


def fit_poly(xs, ys, deg):
    c = np.polyfit(xs, ys, deg)
    return c, ys - np.polyval(c, xs)


def loo_rms(xs, ys, deg):
    """留一法交叉驗證：每次拿掉一個點、用其他點擬合、看預測誤差。比單純的殘差更誠實"""
    n = len(xs)
    if n < deg + 3:
        return float("nan")
    errs = []
    for i in range(n):
        m = np.arange(n) != i
        c = np.polyfit(xs[m], ys[m], deg)
        errs.append(ys[i] - np.polyval(c, xs[i]))
    return float(np.sqrt(np.mean(np.square(errs))))


def fget(var, default=0.0):
    try:
        return float(var.get())
    except Exception:
        return default


SETTINGS_NAME = "SpectrumSetting.json"
# 會存進設定檔的介面變數（App 裡同名的 tk 變數）
PARAM_VARS = ["var_sens", "var_tol", "var_dpi", "var_mode", "var_degamma", "var_deg", "var_norm", "var_bg",
              "var_label", "var_hand", "var_snap", "var_drag", "var_alpha", "var_refs"]


def _settings_candidates():
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))     # 主程式資料夾
    return [os.path.join(here, SETTINGS_NAME), os.path.join(os.path.expanduser("~"), SETTINGS_NAME)]


def load_settings():
    """讀 SpectrumSetting.json（先找程式旁邊，再找使用者資料夾）。回傳 (內容, 路徑)；沒有就回傳空設定"""
    for q in _settings_candidates():
        try:
            with open(q, "r", encoding="utf-8") as fh:
                d = json.load(fh)
            if isinstance(d, dict):
                return d, q
        except Exception:
            continue
    return {}, _settings_candidates()[0]


def save_settings(data, path):
    """寫入設定檔；程式資料夾不能寫（例如裝在 Program Files）就改寫到使用者資料夾。回傳實際路徑或 None"""
    for q in [path] + [c for c in _settings_candidates() if c != path]:
        try:
            with open(q, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
            return q
        except Exception:
            continue
    return None


def ref_band_rows(img):
    """底圖裡「深色色條」所在的列範圍（最長的一段暗色列）；找不到就用全部列"""
    g = img.astype(np.float32).mean(axis=(1, 2))
    lo, hi = float(np.percentile(g, 10)), float(np.percentile(g, 90))
    if hi - lo < 10:
        return 0, img.shape[0]
    dark = list(g < (lo + hi) / 2.0) + [False]
    best, start = (0, 0), None
    for i, d in enumerate(dark):
        if d and start is None:
            start = i
        elif not d and start is not None:
            if i - start > best[1] - best[0]:
                best = (start, i)
            start = None
    if best[1] - best[0] < 8:
        return 0, img.shape[0]
    return best


def detect_base_peaks(img, thr_frac=0.06):
    """偵測底圖色條裡的亮線。回傳 (位置（像素邊界座標，和 ref_map 一致）, 高度)"""
    y0, y1 = ref_band_rows(img)
    prof = img[y0:y1].astype(np.float32).mean(axis=(0, 2))
    xs, hs = find_peaks(prof, sigma=1.5, thr_frac=thr_frac, min_dist=3, half=4)
    return xs + 0.5, hs


def put_canvas_image(cv, x, y):
    """更新畫布上的影像：重用同一個影像項目（不刪掉重建），縮放時才不會閃"""
    if cv._img_id is None:
        cv._img_id = cv.create_image(x, y, image=cv._ph, anchor="nw")
    else:
        cv.itemconfigure(cv._img_id, image=cv._ph)
        cv.coords(cv._img_id, x, y)
    cv.tag_lower(cv._img_id)


def debounce_mpl(canvas, ms=120):
    """matplotlib 圖在拖曳視窗邊框時每一格都重畫會閃：改成「停下來 ms 毫秒後才重畫一次」"""
    try:
        w = canvas.get_tk_widget()
        orig = canvas.resize
        st = {"id": None, "ev": None}

        def fire():
            st["id"] = None
            try:
                orig(st["ev"])
            except Exception:       # 畫面已經被關掉就不用重畫
                pass

        def handler(ev):
            st["ev"] = ev
            if st["id"] is not None:
                w.after_cancel(st["id"])
            st["id"] = w.after(ms, fire)

        w.bind("<Configure>", handler)
    except Exception:
        pass


def _cancel_after(cv, e):
    """畫布被銷毀（離開模式）時取消還沒執行的重畫，避免對已消失的元件操作"""
    if e.widget is cv and getattr(cv, "_after", None) is not None:
        try:
            cv.after_cancel(cv._after)
        except Exception:
            pass
        cv._after = None


def to_photo(im):
    if IMGTK_OK:
        return ImageTk.PhotoImage(im)
    import io
    import base64
    b = io.BytesIO()
    im.save(b, format="PNG")
    return tk.PhotoImage(data=base64.b64encode(b.getvalue()))


# ======================================================================
#  可縮放影像檢視：只畫「看得到的那一塊」，並直接用螢幕像素，所以不卡也不糊
# ======================================================================
class ZoomView(_CanvasBase):
    """滾輪＝縮放（以游標為中心）；右鍵／中鍵拖曳＝平移。
    座標一律用「取樣座標」（像素中心＝整數），和 extract_strip 一致。"""

    def __init__(self, master, **kw):
        super().__init__(master, bg="#101010", highlightthickness=0, takefocus=1, cursor="crosshair", **kw)
        self.bind("<Destroy>", lambda e: _cancel_after(self, e), add="+")
        self.src = None
        self.disp = None
        self.gain = 1.0
        self.bright = True
        self.scale = 1.0
        self.ox = 0.0
        self.oy = 0.0
        self.draw_overlay = None
        self.hook_press = self.hook_motion = self.hook_release = None
        self.hover_cb = None
        self.click_cb = None
        self._ph = None
        self._img_id = None
        self._after = None
        self._need_fit = False
        self._left = None
        self._pan = None
        self._moved = False
        self.bind("<Configure>", self._on_conf)
        self.bind("<MouseWheel>", self._wheel)
        self.bind("<Button-4>", lambda e: self._wheel(e, 1))
        self.bind("<Button-5>", lambda e: self._wheel(e, -1))
        self.bind("<ButtonPress-1>", self._p1)
        self.bind("<B1-Motion>", self._m1)
        self.bind("<ButtonRelease-1>", self._r1)
        for b in (2, 3):
            self.bind(f"<ButtonPress-{b}>", self._pp)
            self.bind(f"<B{b}-Motion>", self._pm)
        self.bind("<Motion>", self._hover)

    # ---- 座標 ----
    def to_screen(self, x, y):
        return (x + 0.5 - self.ox) * self.scale, (y + 0.5 - self.oy) * self.scale

    def to_image(self, sx, sy):
        return sx / self.scale + self.ox - 0.5, sy / self.scale + self.oy - 0.5

    # ---- 影像 ----
    def set_image(self, img, gain=1.0, bright=True, fit=True):
        self.src = img
        self.gain = gain
        self.bright = bright
        self._make_disp()
        if fit:
            if self.winfo_width() > 20:
                self.fit()
            else:
                self._need_fit = True
        self.schedule()

    def set_gain(self, gain):
        self.gain = gain
        self._make_disp()
        self.schedule()

    def _make_disp(self):
        if self.src is None:
            self.disp = None
        elif self.bright:
            lut = (np.clip(np.arange(256) / 255.0 * self.gain, 0, 1) ** 0.5 * 255).astype(np.uint8)
            self.disp = lut[self.src]
        else:
            self.disp = self.src

    def fit(self):
        if self.src is None:
            return
        H, W = self.src.shape[:2]
        w, h = max(20, self.winfo_width()), max(20, self.winfo_height())
        self.scale = max(0.02, min(w / W, h / H))
        self.ox = W / 2.0 - w / (2 * self.scale)
        self.oy = H / 2.0 - h / (2 * self.scale)
        self.schedule()

    def zoom_to_rect(self, x0, y0, x1, y1):
        w, h = max(20, self.winfo_width()), max(20, self.winfo_height())
        rw, rh = max(1.0, x1 - x0), max(1.0, y1 - y0)
        self.scale = min(80.0, max(0.02, min(w / rw, h / rh)))
        cx, cy = (x0 + x1) / 2.0 + 0.5, (y0 + y1) / 2.0 + 0.5
        self.ox = cx - w / (2 * self.scale)
        self.oy = cy - h / (2 * self.scale)
        self.schedule()

    # ---- 繪製 ----
    def schedule(self):
        if self._after is None:
            self._after = self.after(8, self.redraw)

    def redraw(self):
        self._after = None
        self.delete("ov")
        w, h = self.winfo_width(), self.winfo_height()
        if self.disp is None or w < 4 or h < 4:
            self._clear_img()
            return
        H, W = self.disp.shape[:2]
        s = self.scale
        ix0, iy0 = max(0, int(math.floor(self.ox))), max(0, int(math.floor(self.oy)))
        ix1, iy1 = min(W, int(math.ceil(self.ox + w / s))), min(H, int(math.ceil(self.oy + h / s)))
        if ix1 > ix0 and iy1 > iy0:
            crop = Image.fromarray(np.ascontiguousarray(self.disp[iy0:iy1, ix0:ix1]))
            dw, dh = max(1, int(round((ix1 - ix0) * s))), max(1, int(round((iy1 - iy0) * s)))
            if s >= 1.0:
                im = crop.resize((dw, dh), RS.NEAREST)
            else:
                im = crop.resize((dw, dh), RS.BILINEAR, reducing_gap=2.0)
            self._ph = to_photo(im)
            put_canvas_image(self, (ix0 - self.ox) * s, (iy0 - self.oy) * s)
        else:
            self._clear_img()
        if self.draw_overlay:
            self.draw_overlay(self)

    def _clear_img(self):
        if self._img_id is not None:
            self.delete(self._img_id)
            self._img_id = None

    def redraw_overlay(self):
        self.delete("ov")
        if self.draw_overlay:
            self.draw_overlay(self)

    # ---- 滑鼠 ----
    def _wheel(self, ev, dirn=None):
        if self.src is None:
            return
        if dirn is None:
            dirn = 1 if ev.delta > 0 else -1
        f = 1.25 if dirn > 0 else 0.8
        ns = min(80.0, max(0.02, self.scale * f))
        ix, iy = ev.x / self.scale + self.ox, ev.y / self.scale + self.oy
        self.scale = ns
        self.ox, self.oy = ix - ev.x / ns, iy - ev.y / ns
        self.schedule()

    def _p1(self, ev):
        self.focus_set()
        self._left = None
        if self.hook_press and self.hook_press(ev):
            self._left = "hook"
            return
        self._left = "pan"
        self._pan = (ev.x, ev.y, self.ox, self.oy)
        self._moved = False

    def _m1(self, ev):
        if self._left == "hook":
            if self.hook_motion:
                self.hook_motion(ev)
        elif self._left == "pan":
            sx0, sy0, ox0, oy0 = self._pan
            if abs(ev.x - sx0) + abs(ev.y - sy0) > 3:
                self._moved = True
            self.ox = ox0 - (ev.x - sx0) / self.scale
            self.oy = oy0 - (ev.y - sy0) / self.scale
            self.schedule()

    def _r1(self, ev):
        if self._left == "hook":
            if self.hook_release:
                self.hook_release(ev)
        elif self._left == "pan" and not self._moved and self.click_cb:
            self.click_cb(*self.to_image(ev.x, ev.y))
        self._left = None

    def _pp(self, ev):
        self.focus_set()
        self._pan = (ev.x, ev.y, self.ox, self.oy)

    def _pm(self, ev):
        if self._pan is None:
            return
        sx0, sy0, ox0, oy0 = self._pan
        self.ox = ox0 - (ev.x - sx0) / self.scale
        self.oy = oy0 - (ev.y - sy0) / self.scale
        self.schedule()

    def _hover(self, ev):
        if self.hover_cb and self.src is not None:
            self.hover_cb(*self.to_image(ev.x, ev.y))

    def _on_conf(self, e):
        if self._need_fit and e.width > 20 and e.height > 20:
            self._need_fit = False
            self.fit()
        # 拖曳視窗邊框時不要每一格都重畫（會閃）：停下來 70 毫秒後才畫一次
        if self._after is not None:
            self.after_cancel(self._after)
        self._after = self.after(70, self.redraw)


# ======================================================================
#  對圖檢視：樣本當底圖、你的光譜疊在上面、下面是亮度曲線（共用同一條波長軸）
# ======================================================================
class AlignView(_CanvasBase):
    TOP, BOT = 44, 18

    def __init__(self, master, app):
        super().__init__(master, bg="#101010", highlightthickness=0, takefocus=1, cursor="crosshair")
        self.bind("<Destroy>", lambda e: _cancel_after(self, e), add="+")
        self.app = app
        self.vc, self.vs = 600.0, 300.0
        self.ref_pil = None
        self.ov_pil = None
        self.dirty = True
        self.fitted = False
        self._ph = None
        self._img_id = None
        self._after = None
        self._drag = None
        self.pick = None            # 兩點對齊進行中的狀態
        self.bind("<Configure>", self.on_conf)
        self.bind("<MouseWheel>", self.on_wheel)
        self.bind("<Button-4>", lambda e: self.on_wheel(e, 1))
        self.bind("<Button-5>", lambda e: self.on_wheel(e, -1))
        self.bind("<ButtonPress-1>", self.on_p1)
        self.bind("<B1-Motion>", self.on_m1)
        for b in (2, 3):
            self.bind(f"<ButtonPress-{b}>", self.on_pp)
            self.bind(f"<B{b}-Motion>", self.on_pm)
        self.bind("<Motion>", self.on_hover)
        self.bind("<Key>", self.on_key)

    def on_conf(self, e):
        # 拖曳視窗邊框時不要每一格都重畫（會閃）：停下來 70 毫秒後才畫一次
        if self._after is not None:
            self.after_cancel(self._after)
        self._after = self.after(70, self.redraw)

    # ---- 轉換 ----
    def layout(self):
        W, H = self.winfo_width(), self.winfo_height()
        body = max(40, H - self.TOP - self.BOT - 6)
        hb = int(body * 0.55)
        py0 = self.TOP + hb + 6
        py1 = H - self.BOT
        return W, H, hb, py0, py1

    def x2lam(self, x, W):
        return self.vc - self.vs / 2.0 + x / W * self.vs

    def lam2x(self, lam, W):
        return (lam - (self.vc - self.vs / 2.0)) / self.vs * W

    def fit(self):
        a = self.app
        refs = a.refs_cache
        if refs:
            lo, hi = min(refs) - 25, max(refs) + 25
        elif a.prof is not None:
            lo, hi = float(a.lam_lin(0)), float(a.lam_lin(len(a.prof) - 1))
        else:
            lo, hi = 300.0, 800.0
        self.vc, self.vs = (lo + hi) / 2.0, max(5.0, hi - lo)
        self.fitted = True
        self.schedule()

    def refresh_sources(self):
        a = self.app
        self.ref_pil = Image.fromarray(a.ref_img) if a.ref_img is not None else None
        self.ov_pil = None
        if a.strips:
            st = a.strips[0]
            m = max(float(np.percentile(st, 99.8)), 20.0)
            self.ov_pil = Image.fromarray((np.clip(st / m, 0, 1) ** 0.5 * 255).astype(np.uint8))
        self.dirty = False

    # ---- 繪製 ----
    def schedule(self):
        if self._after is None:
            self._after = self.after(8, self.redraw)

    def redraw(self):
        self._after = None
        self.delete(*[i for i in self.find_all() if i != self._img_id])   # 保留底圖項目，只換掉其他東西
        W, H, hb, py0, py1 = self.layout()
        if W < 50 or H < 120:
            return
        a = self.app
        if self.dirty:
            self.refresh_sources()
        top = self.TOP
        ks = self.app.k
        left = self.vc - self.vs / 2.0
        k = self.vs / W
        band = np.empty((hb, W, 3), np.float32)
        band[:] = (16, 16, 16)
        have_base = False
        if self.ref_pil is not None and a.ref_map:
            x1, l1, x2, l2 = a.ref_map
            kref = (l2 - l1) / (x2 - x1)
            Wr, Hr = self.ref_pil.size
            lamL, lamR = a.ref_lam(0), a.ref_lam(Wr)
            lo, hi = min(lamL, lamR), max(lamL, lamR)
            va, vb = max(left, lo), min(left + self.vs, hi)
            if vb > va:
                pa, pb = sorted(((va - l1) / kref + x1, (vb - l1) / kref + x1))
                xa, xb = int(round((va - left) / k)), int(round((vb - left) / k))
                xa, xb = max(0, xa), min(W, xb)
                if xb - xa >= 1:
                    im = self.ref_pil.resize((xb - xa, hb), RS.BILINEAR,
                                             box=(max(0.0, pa), 0, min(float(Wr), pb), float(Hr)))
                    band[:, xa:xb] = np.asarray(im, np.float32)
                    have_base = True
        if self.ov_pil is not None and a.prof is not None:
            nu, nv = self.ov_pil.size
            lamS0, lamS1 = float(a.lam_lin(-0.5)), float(a.lam_lin(nu - 0.5))
            va, vb = max(left, lamS0), min(left + self.vs, lamS1)
            if vb > va:
                ua = (va - lamS0) / (lamS1 - lamS0) * nu
                ub = (vb - lamS0) / (lamS1 - lamS0) * nu
                xa, xb = int(round((va - left) / k)), int(round((vb - left) / k))
                xa, xb = max(0, xa), min(W, xb)
                if xb - xa >= 1:
                    zoomed = (xb - xa) > (ub - ua)
                    im = self.ov_pil.resize((xb - xa, hb), RS.NEAREST if zoomed else RS.BILINEAR,
                                            box=(ua, 0, ub, float(nv)))
                    al = fget(a.var_alpha, 0.5) if have_base else 1.0
                    band[:, xa:xb] = band[:, xa:xb] * (1 - al) + np.asarray(im, np.float32) * al
        self._ph = to_photo(Image.fromarray(np.clip(band, 0, 255).astype(np.uint8)))
        put_canvas_image(self, 0, top)
        if not have_base:
            self.create_text(W // 2, top + hb // 2 + 14, fill="#888", text="（尚未載入樣本底圖：按上方「載入參考圖」）")

        # ---- 波長刻度（下方）----
        step = 1.0
        for st in (0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50, 100, 200, 500):
            step = st
            if st / k >= 70:
                break
        t = math.ceil(left / step) * step
        while t <= left + self.vs:
            x = (t - left) / k
            self.create_line(x, py1, x, py1 + int(5 * ks), fill="#aaa")
            self.create_text(x + 2, py1 + int(6 * ks), text=f"{t:g}", fill="#bbb", anchor="nw", font=self.app.f_small)
            t += step

        # ---- 參考波長（綠色虛線）----
        for r in a.refs_cache:
            x = self.lam2x(r, W)
            if -5 <= x <= W + 5:
                self.create_line(x, top, x, py1, fill="#2ecc71", dash=(2, 3))
                try:
                    self.create_text(x, top - 3, text=f"{r:g}", fill="#35d07f", anchor="w", angle=90,
                                     font=self.app.f_small)
                except tk.TclError:
                    self.create_text(x, top - 8, text=f"{r:g}", fill="#35d07f", font=self.app.f_small)

        # ---- 底圖自動偵測到的亮線（青色短線，畫在底圖上緣）----
        for lam in a.base_pk_lam:
            x = self.lam2x(float(lam), W)
            if 0 <= x <= W:
                self.create_line(x, top, x, top + int(10 * ks), fill="#00e5ff", width=2)

        # ---- 亮度曲線 ----
        self.create_rectangle(0, py0, W, py1, outline="#333")
        if a.prof is not None and len(a.prof) > 2:
            n = len(a.prof)
            dl = float(a.lam_lin(1) - a.lam_lin(0))
            lam0 = float(a.lam_lin(0))
            ia = max(0, int(math.floor((left - lam0) / dl)) - 1)
            ib = min(n - 1, int(math.ceil((left + self.vs - lam0) / dl)) + 1)
            pr = a.prof
            rng = float(pr.max() - pr.min()) or 1.0
            hh = py1 - py0 - 10

            def yv(v):
                return py1 - 5 - (v - float(pr.min())) / rng * hh

            if ib > ia:
                idx = np.arange(ia, ib + 1)
                if len(idx) > 2 * W:
                    idx = idx[::int(math.ceil(len(idx) / (2 * W)))]
                xs = (lam0 + idx * dl - left) / k
                ys = yv(pr[idx])
                co = np.column_stack([xs, ys]).ravel().tolist()
                if len(co) >= 4:
                    self.create_line(*co, fill="#58a6ff")
            for px in a.pk_x:
                x = (lam0 + px * dl - left) / k
                if 0 <= x <= W:
                    y = yv(float(np.interp(px, np.arange(n), pr)))
                    self.create_polygon(x, y - 3 * ks, x - 4 * ks, y - 11 * ks, x + 4 * ks, y - 11 * ks,
                                        fill="#ff4d4d", outline="")
                    self.create_line(x, top + hb - 9 * ks, x, top + hb, fill="#ff4d4d")
        if a.prof is not None:
            self.draw_handles(W, hb, py1)
        self.draw_pick(W, hb, py1)

    # ---- 左右兩端的把手（光譜在中間）----
    def handle_x(self):
        a = self.app
        W = max(1, self.winfo_width())
        c, s = fget(a.var_center, 600.0), fget(a.var_span, 200.0)
        return self.lam2x(c - s / 2.0, W), self.lam2x(c + s / 2.0, W)

    def hit_handle(self, x):
        if self.app.prof is None:
            return None
        xl, xr = self.handle_x()
        dl, dr = abs(x - xl), abs(x - xr)
        if min(dl, dr) > 10 * self.app.k:
            return None
        return "L" if dl <= dr else "R"

    def draw_handles(self, W, hb, py1):
        top = self.TOP
        ks = self.app.k
        col = "#ff9100"
        for x in self.handle_x():
            if 0 <= x <= W:
                ym = top + hb // 2
                self.create_line(x, top, x, py1, fill=col, width=3)
                self.create_rectangle(x - 9 * ks, ym - 22 * ks, x + 9 * ks, ym + 22 * ks, fill=col, outline="black")
                self.create_text(x, ym, text="◀▶", fill="black", font=self.app.f_small)
            else:
                left = x < 0
                self.create_text(12 * ks if left else W - 12 * ks, top + hb // 2, fill=col, anchor="w" if left else "e",
                                 text="◀ 左端（縮小畫面可看到）" if left else "右端（縮小畫面可看到） ▶")

    # ---- 兩點對齊：我的光譜 2 點 ↔ 底圖 2 點 ----
    PICK_TEXT = ["我的光譜的第 1 點（點藍色曲線上的亮線）", "底圖上對應的第 1 點", "我的光譜的第 2 點", "底圖上對應的第 2 點"]

    def start_pick(self):
        if self.app.prof is None:
            messagebox.showinfo("提示", "請先在第 1 頁開啟照片並設好裁切線。")
            return
        if self.pick:
            self.cancel_pick()
            return
        self.pick = {"step": 0, "u": [], "lam": []}
        self.configure(cursor="tcross")
        self.focus_set()
        self.say_pick()
        self.schedule()

    def cancel_pick(self):
        self.pick = None
        self.configure(cursor="crosshair")
        self.app.status.set("兩點對齊已取消。")
        self.schedule()

    def say_pick(self):
        s = self.pick["step"]
        self.app.status.set(f"兩點對齊 {s + 1}/4：請點「{self.PICK_TEXT[s]}」。右鍵拖曳可平移、滾輪可縮放，Esc 取消。")

    def pick_click(self, x):
        a, p = self.app, self.pick
        W = max(1, self.winfo_width())
        lam = self.x2lam(x, W)
        snap = bool(a.var_snap.get())
        if p["step"] % 2 == 0:          # 我的光譜上的點 -> 換成「沿線位置 u (px)」
            L = max(1, len(a.prof) - 1)
            l0, l1 = float(a.lam_lin(0)), float(a.lam_lin(L))
            u = (lam - l0) / (l1 - l0) * L
            if snap and len(a.pk_x):
                xs = (np.asarray(a.lam_lin(a.pk_x), float) - (self.vc - self.vs / 2.0)) / self.vs * W
                i = int(np.argmin(np.abs(xs - x)))
                if abs(xs[i] - x) <= 10 * a.k:
                    u = float(a.pk_x[i])
            p["u"].append(u)
        else:                           # 底圖上的點 -> 波長
            cands = list(a.refs_cache) + [float(v) for v in a.base_pk_lam]    # 參考波長＋底圖偵測到的亮線
            if snap and cands:
                rx = np.array([self.lam2x(r, W) for r in cands])
                i = int(np.argmin(np.abs(rx - x)))
                if abs(rx[i] - x) <= 10 * a.k:
                    lam = cands[i]
            p["lam"].append(lam)
        p["step"] += 1
        if p["step"] >= 4:
            self.finish_pick()
        else:
            self.say_pick()
        self.schedule()

    def finish_pick(self):
        a = self.app
        (u1, u2), (l1, l2) = self.pick["u"], self.pick["lam"]
        self.pick = None
        self.configure(cursor="crosshair")
        L = max(1, len(a.prof) - 1)
        if abs(u2 - u1) < 5 or abs(l2 - l1) < 1.0:
            messagebox.showwarning("兩點對齊", "兩個點太近了，請選距離遠一點的兩點。")
            a.status.set("兩點對齊已取消（點太近）。")
            return
        span = (l2 - l1) * L / (u2 - u1)
        if span <= 0:
            messagebox.showwarning("兩點對齊", "兩個點的左右順序和底圖相反（光譜方向相反）。請檢查點的順序，"
                                               "或回第 1 頁按「⇄ 反向」。")
            a.status.set("兩點對齊已取消（方向相反）。")
            return
        lam0 = l1 - span * u1 / L
        a.var_span.set(round(span, 3))
        a.var_center.set(round(lam0 + span / 2.0, 3))
        a.status.set(f"兩點對齊完成：中心 {lam0 + span / 2.0:.2f} nm、寬度 {span:.1f} nm。"
                     "請再用把手或方向鍵檢查、微調。")

    def draw_pick(self, W, hb, py1):
        p = self.pick
        if not p:
            return
        a = self.app
        self.create_text(8, 6, anchor="nw", fill="#ffd54f", font=self.app.f_bold,
                         text=f"兩點對齊 {p['step'] + 1}/4：請點「{self.PICK_TEXT[p['step']]}」（Esc 取消）")
        for i, u in enumerate(p["u"]):
            x = self.lam2x(float(a.lam_lin(u)), W)
            self.create_line(x, self.TOP, x, py1, fill="#ff9100", dash=(4, 3), width=2)
            self.create_text(x + 3, self.TOP + 14 * self.app.k, anchor="w", fill="#ff9100", text=f"我{i + 1}")
        for i, lam in enumerate(p["lam"]):
            x = self.lam2x(lam, W)
            self.create_line(x, self.TOP, x, py1, fill="#00e5ff", dash=(4, 3), width=2)
            self.create_text(x + 3, self.TOP + 28 * self.app.k, anchor="w", fill="#00e5ff", text=f"底{i + 1}")

    def draw_cursor(self, x):
        self.delete("cur")
        W, H, hb, py0, py1 = self.layout()
        if W < 50:
            return
        lam = self.x2lam(x, W)
        self.create_line(x, self.TOP, x, py1, fill="#ffd54f", tags="cur")
        self.app.var_cur.set(f"游標：{lam:.2f} nm　　視窗寬 {self.vs:.1f} nm")

    # ---- 滑鼠與鍵盤 ----
    def on_hover(self, ev):
        self.draw_cursor(ev.x)
        self.configure(cursor="sb_h_double_arrow" if self.hit_handle(ev.x) else "crosshair")

    def on_wheel(self, ev, dirn=None):
        if dirn is None:
            dirn = 1 if ev.delta > 0 else -1
        W = max(1, self.winfo_width())
        lam = self.x2lam(ev.x, W)
        f = 0.85 if dirn > 0 else 1 / 0.85
        a = self.app
        if ev.state & 0x1:      # Shift＋滾輪：以游標為中心縮放「我的光譜」
            c, s = fget(a.var_center, 600.0), fget(a.var_span, 200.0)
            u = (lam - (c - s / 2.0)) / s
            s2 = min(2000.0, max(5.0, s / f))
            a.var_center.set(round(lam - u * s2 + s2 / 2.0, 3))
            a.var_span.set(round(s2, 3))
        else:                   # 滾輪：縮放畫面
            self.vs = min(3000.0, max(2.0, self.vs * f))
            self.vc = lam - (lam - self.vc) * f
        self.schedule()
        self.draw_cursor(ev.x)

    def on_p1(self, ev):
        self.focus_set()
        a = self.app
        if self.pick:                   # 兩點對齊中：左鍵＝選點
            self._drag = None
            self.pick_click(ev.x)
            return
        h = self.hit_handle(ev.x)
        if h:       # 抓住左／右把手：記下目前左右兩端的波長
            c, s = fget(a.var_center, 600.0), fget(a.var_span, 200.0)
            self._drag = ("H" + h, ev.x, c - s / 2.0, c + s / 2.0)
            return
        self._drag = ("spec" if a.var_drag.get() == "spec" else "pan", ev.x, fget(a.var_center, 600.0), self.vc)

    def on_m1(self, ev):
        if not self._drag:
            return
        if self._drag[0] in ("HL", "HR"):   # 拖把手：被拖的那端跟著游標，另一端不動
            mode, _, lam_l, lam_r = self._drag
            lam = self.x2lam(ev.x, max(1, self.winfo_width()))
            if mode == "HL":
                lam_l = min(lam, lam_r - 5.0)
            else:
                lam_r = max(lam, lam_l + 5.0)
            self.app.var_center.set(round((lam_l + lam_r) / 2.0, 3))
            self.app.var_span.set(round(lam_r - lam_l, 3))
            self.schedule()
            self.draw_cursor(ev.x)
            return
        mode, x0, c0, vc0 = self._drag
        W = max(1, self.winfo_width())
        dl = (ev.x - x0) / W * self.vs
        if mode == "spec":
            self.app.var_center.set(round(c0 + dl, 3))
        else:
            self.vc = vc0 - dl
        self.schedule()
        self.draw_cursor(ev.x)

    def on_pp(self, ev):
        self.focus_set()
        self._drag = ("pan", ev.x, fget(self.app.var_center, 600.0), self.vc)

    def on_pm(self, ev):
        if not self._drag:
            return
        _, x0, _, vc0 = self._drag
        self.vc = vc0 - (ev.x - x0) / max(1, self.winfo_width()) * self.vs
        self.schedule()

    def on_key(self, ev):
        a = self.app
        ks = ev.keysym
        if ks == "Escape" and self.pick:
            self.cancel_pick()
            return
        shift, ctrl = bool(ev.state & 0x1), bool(ev.state & 0x4)
        if ks in ("Left", "Right"):
            step = 1.0 if shift else (0.02 if ctrl else 0.1)
            a.var_center.set(round(fget(a.var_center, 600.0) + (step if ks == "Right" else -step), 3))
        elif ks in ("Up", "Down"):
            f = 1.0 + (0.01 if shift else (0.0002 if ctrl else 0.001))
            a.var_span.set(round(fget(a.var_span, 200.0) * (f if ks == "Up" else 1 / f), 3))
        else:
            return
        self.schedule()


# ======================================================================
#  視窗
# ======================================================================
# 提示文字的顏色跟著主程式色票（淺色／深色都看得到）；找不到色票就用灰色
class SpectrumMode(Mode):
    title = "🌈 光譜校正"
    order = 50

    def __init__(self, app):
        super().__init__(app)
        self.root = app                       # 舊程式碼裡的 self.root 就是主視窗
        self.cfg, self.cfg_path = load_settings()      # SpectrumSetting.json
        self.k = max(0.6, float(getattr(app, "scale", 1.0)))
        self.page = None
        self._restored = False
        self._hooked = False
        self._save_id = None
        self._rc_old = None

        # ---- 狀態（離開模式再回來還會在）----
        self.photos = []                 # [{name, path, img}]，第 0 張是氖燈
        self.p0 = self.p1 = None         # 裁切線起點、終點（原圖取樣座標）
        self.strips = []
        self.sel = "body"
        self.drag = None
        self.prof = None
        self.pk_x = np.array([])
        self.pk_h = np.array([])
        self.pk_ref = []
        self.prof_dirty = True
        self.init_scale = True
        self.coef = None
        self.calib = None
        self.ref_img = None
        self.ref_map = None
        self.refs_cache = self._parse_refs(DEFAULT_REFS)
        self._timers = {}
        self._scales = []
        self._tree_w = [40, 85, 90, 90, 90, 85]
        self.ref_path = None                                  # 底圖（參考圖）檔案路徑
        self.base_pk_lam = np.array([])                       # 底圖自動偵測到的亮線波長
        self.dirs = dict(self.cfg["dirs"]) if isinstance(self.cfg.get("dirs"), dict) else {}

        # ---- 變數 ----
        D = lambda **kw: tk.DoubleVar(master=app, **kw)
        S = lambda **kw: tk.StringVar(master=app, **kw)
        B = lambda **kw: tk.BooleanVar(master=app, **kw)
        self.var_width = D(value=1.0)
        self.var_shear = D(value=0.0)
        self.var_gain = D(value=1.0)
        self.var_view = S(value="")
        self.view_idx = 0
        self.var_center = D(value=600.0)
        self.var_span = D(value=200.0)
        self.var_alpha = D(value=0.5)
        self.var_sens = D(value=2.0)
        self.var_tol = D(value=1.5)
        self.var_dpi = D(value=300.0)
        self.var_mode = S(value="ImageJ(RGB平均)")
        self.var_degamma = B(value=False)
        self.var_deg = S(value="1")
        self.var_norm = B(value=False)     # 預設輸出原始灰階值（同 ImageJ）
        self.var_bg = B(value=False)
        self.var_label = B(value=False)
        self.var_hand = B(value=False)
        self.var_snap = B(value=True)
        self.var_drag = S(value="spec")
        self.var_refs = S(value=DEFAULT_REFS)
        self.var_hover = S(value="")
        self.var_cur = S(value="")
        self.status = S(value="請先按「開啟氖燈照片」。")
        self.apply_cfg_params()

    # ---------------- 進入模式：畫出整個畫面 ----------------
    def start(self):
        app = self.app
        self._cancel_timers()
        app.clear_container()
        self.k = max(0.6, float(getattr(app, "scale", 1.0)))
        self._scales = []
        self.page = ttk.Frame(app.container)
        self.page.pack(fill="both", expand=True)
        self.page.bind("<Destroy>", self._on_page_destroy, add="+")
        self._init_fonts()
        try:                                  # matplotlib 字體依主程式的「介面大小」放大；離開時還原
            self._rc_old = float(matplotlib.rcParams["font.size"])
            matplotlib.rcParams["font.size"] = self._rc_old * float(getattr(app, "ui_scale", 1.0))
        except Exception:
            self._rc_old = None
        ttk.Label(self.page, textvariable=self.status, anchor="w", relief="sunken").pack(side="bottom", fill="x")
        ub = ttk.Frame(self.page)
        ub.pack(side="top", fill="x")
        ttk.Button(ub, text="返回主選單", command=self.go_home).pack(side="left", padx=6, pady=2)
        ttk.Label(ub, text="（介面大小、主題請在主選單調整）", style="Hint.TLabel").pack(side="left", padx=8)
        self.nb = ttk.Notebook(self.page)
        self.nb.pack(fill="both", expand=True)
        self._build1()
        self._build2()
        self._build3()
        self._build4()
        self.nb.bind("<<NotebookTabChanged>>", self.on_tab)
        self.apply_ui_scale()
        self.page.update_idletasks()
        if not self._restored:
            self._restored = True
            self.restore_reference()
        self.restore_ui()
        if not self._hooked:                  # 參數一改就（延遲）存檔
            self._hooked = True
            for name in PARAM_VARS:
                try:
                    getattr(self, name).trace_add("write", lambda *a: self._queue_save())
                except Exception:
                    pass

    def go_home(self):
        self.save_settings_now()
        self.app.show_main_home()

    def _on_page_destroy(self, e):
        if e.widget is not self.page:
            return
        self._cancel_timers()
        try:
            if self._rc_old is not None:
                matplotlib.rcParams["font.size"] = self._rc_old
            self.save_settings_now()
        except Exception:
            pass

    def _cancel_timers(self):
        for key, tid in list(self._timers.items()):
            try:
                self.app.after_cancel(tid)
            except Exception:
                pass
        self._timers = {}

    def _queue_save(self):
        if self._save_id is not None:
            try:
                self.app.after_cancel(self._save_id)
            except Exception:
                pass
        self._save_id = self.app.after(1500, self.save_settings_now)

    def restore_ui(self):
        """回到模式時，把之前載入的照片、裁切線、底圖重新顯示在新建的畫面上"""
        if self.photos:
            self.update_view_list()
            ph = self.photos[min(self.view_idx, len(self.photos) - 1)]
            self.var_gain.set(ph.get("gain", fget(self.var_gain, 1.0)))
            self.view.set_image(ph["img"], fget(self.var_gain, 1.0), bright=True, fit=True)
            if self.p0 is not None:
                self.focus_line()
                self.refresh_strips()
        if self.ref_img is not None:
            self.align.dirty = True
            self.align.schedule()

    # ---------------- 字體／介面縮放（尺寸由 main.py 決定）----------------
    def _init_fonts(self):
        base = tkfont.nametofont("TkDefaultFont")
        sz, fam = int(base.cget("size")), base.cget("family")

        def mk(ratio, bold=False):
            n = sz * ratio
            n = max(6, int(round(n))) if sz > 0 else -max(8, int(round(-n)))
            return tkfont.Font(root=self.app, family=fam, size=n, weight="bold" if bold else "normal")
        self.f_small = mk(8 / 9)
        self.f_bold = mk(10 / 9, True)
        self.f_res = mk(11 / 9, True)

    def apply_ui_scale(self):
        """版面倍率 k 直接用主程式的縮放；像素大小的東西（滑桿長度、表格欄寬、把手…）乘上它"""
        self.k = max(0.6, float(getattr(self.app, "scale", 1.0)))
        for sc in self._scales:
            sc.configure(length=int(420 * self.k))
        for c, w in zip(self.tree["columns"], self._tree_w):
            self.tree.column(c, width=int(w * self.k))
        self.align.TOP = max(int(44 * self.k), self.f_small.measure("000.00") + int(14 * self.k))
        self.align.BOT = int(6 * self.k + self.f_small.metrics("linespace") + 6)
        self.view.schedule()
        self.align.schedule()

    # ---------------- 設定檔 SpectrumSetting.json ----------------
    def apply_cfg_params(self):
        p = self.cfg.get("params")
        if not isinstance(p, dict):
            return
        for name in PARAM_VARS:
            if name in p:
                try:
                    getattr(self, name).set(p[name])
                except Exception:
                    pass
        self.refs_cache = self._parse_refs(self.var_refs.get())

    def collect_cfg(self):
        params = {}
        for name in PARAM_VARS:
            try:
                params[name] = getattr(self, name).get()
            except Exception:
                pass
        ref = {"path": self.ref_path,
               "map": [float(v) for v in self.ref_map] if self.ref_map else None,
               "size": [int(self.ref_img.shape[1]), int(self.ref_img.shape[0])] if self.ref_img is not None else None}
        return {"version": 2, "reference": ref, "params": params, "dirs": self.dirs}

    def save_settings_now(self):
        self._save_id = None
        q = save_settings(self.collect_cfg(), self.cfg_path)
        if q:
            self.cfg_path = q
        return q

    def restore_reference(self):
        """啟動時自動載入上次的底圖；橫軸校準只在「圖片尺寸沒變」時才還原"""
        r = self.cfg.get("reference")
        if not isinstance(r, dict) or not r.get("path"):
            return
        p = r["path"]
        if not os.path.isfile(p):
            self.status.set(f"找不到上次的底圖：{p}（請在第 2 頁重新載入）")
            return
        try:
            self.ref_img = np.asarray(Image.open(p).convert("RGB"))
        except Exception:
            return
        self.ref_path = p
        m, sz = r.get("map"), r.get("size")
        if (isinstance(m, list) and len(m) == 4 and isinstance(sz, list) and len(sz) == 2
                and (int(sz[0]), int(sz[1])) == (self.ref_img.shape[1], self.ref_img.shape[0])):
            try:
                self.ref_map = tuple(float(v) for v in m)
            except Exception:
                self.ref_map = None
        self.update_base_peaks()
        self.align.dirty = True
        self.align.schedule()
        self.status.set("已載入上次的底圖" + ("與橫軸校準。" if self.ref_map else "，但橫軸尚未校準（第 2 頁按「重新校準參考圖橫軸」）。"))

    def _d(self, key):
        d = self.dirs.get(key)
        return d if isinstance(d, str) and os.path.isdir(d) else None

    def _rem(self, key, path):
        if path:
            self.dirs[key] = os.path.dirname(path)

    # ---------------- 底圖自動偵測與自動對齊 ----------------
    def update_base_peaks(self):
        self.base_pk_lam = np.array([])
        if self.ref_img is None or not self.ref_map:
            return
        try:
            bx, _ = detect_base_peaks(self.ref_img)
            self.base_pk_lam = np.array([self.ref_lam(x) for x in bx], float)
        except Exception:
            pass

    def auto_align_base(self):
        """偵測底圖色條裡的亮線，和我的亮線自動配對、算出縮放與平移"""
        if not self.compute_profile():
            messagebox.showinfo("提示", "請先在第 1 頁開啟照片並設好裁切線。")
            return
        if self.ref_img is None or not self.ref_map:
            messagebox.showinfo("提示", "請先載入底圖並校準橫軸（按「載入參考圖」）。")
            return
        self.update_base_peaks()
        blam = np.sort(self.base_pk_lam)
        if len(self.pk_x) < 3 or len(blam) < 3:
            messagebox.showinfo("提示", f"偵測到的亮線不夠（我的 {len(self.pk_x)} 條、底圖 {len(blam)} 條），"
                                        "無法自動對齊。可調整「峰值靈敏度」。")
            self.align.schedule()
            return
        self.status.set("自動偵測對齊中……")
        self.root.update_idletasks()
        L = len(self.prof) - 1
        r = coarse_align(self.pk_x, self.pk_h, blam, L)
        lam0, k = r["lam0"], r["k"]
        idx = match_peaks(lam0 + k * self.pk_x, blam, 1.2)
        sel = [i for i, j in enumerate(idx) if j >= 0]
        rms = float("nan")
        if len(sel) >= 3:           # 用配到的線做直線擬合，比粗對更準
            c, res = fit_poly(self.pk_x[sel], np.array([blam[idx[i]] for i in sel]), 1)
            if c[0] > 0:
                k, lam0 = float(c[0]), float(c[1])
                rms = float(np.sqrt(np.mean(res ** 2)))
        span = k * L
        self.var_span.set(round(min(1000.0, max(10.0, span)), 3))
        self.var_center.set(round(lam0 + span / 2.0, 3))
        self.align.schedule()
        warn = "" if len(sel) >= 0.6 * min(len(self.pk_x), len(blam)) else "　⚠ 配到的線偏少，結果可能不可信，請看圖確認"
        self.status.set(f"自動偵測對齊：我的 {len(self.pk_x)} 條線、底圖 {len(blam)} 條線，配到 {len(sel)} 條，"
                        f"RMS {rms:.2f} nm（吻合分數 {r['score']:.2f}）。請看圖確認，再用把手微調。" + warn)

    # ---------------- 共用小工具 ----------------
    def _later(self, key, fn, ms=60):
        if key in self._timers:
            self.root.after_cancel(self._timers[key])
        self._timers[key] = self.root.after(ms, fn)

    def _slider(self, parent, text, var, lo, hi, cmd, row, step=0.1):
        ttk.Label(parent, text=text).grid(row=row, column=0, sticky="w")
        sc = ttk.Scale(parent, from_=lo, to=hi, variable=var, length=int(420 * self.k),
                       command=lambda _v: cmd())
        sc.grid(row=row, column=1, padx=6, pady=1)
        self._scales.append(sc)
        sp = ttk.Spinbox(parent, from_=lo, to=hi, increment=step, textvariable=var, width=8, command=cmd)
        sp.grid(row=row, column=2)
        sp.bind("<Return>", lambda e: cmd())

    def _mpl(self, parent, fig):
        tbf = ttk.Frame(parent)
        tbf.pack(side="bottom", fill="x")
        canvas = FigureCanvasTkAgg(fig, master=parent)
        tb = NavigationToolbar2Tk(canvas, tbf)
        tb.update()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        debounce_mpl(canvas)
        return canvas, tb

    def try_load(self, p):
        try:
            return load_image(p)
        except Exception as e:
            hint = ""
            if p.lower().endswith((".heic", ".heif")) and not HEIF_OK:
                hint = "\n\n這是 HEIC 檔：請 pip install pillow-heif，或把 iPad 相機格式改成「最相容」(JPEG)。"
            messagebox.showerror("開啟失敗", f"{e}{hint}")
            return None

    def basis(self):
        d = self.p1 - self.p0
        L = float(np.hypot(d[0], d[1])) or 1e-9
        e = d / L
        return e, np.array([-e[1], e[0]]), L

    def cur_width(self):
        return max(1, int(round(fget(self.var_width, 1.0))))

    # ======================================================================
    #  第 1 頁：裁切（一條線）
    # ======================================================================
    def _build1(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text=" 1 裁切（一條線） ")
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=6, pady=4)
        for t, c in [("開啟氖燈照片", self.open_neon), ("加入其他光譜照片", self.add_samples),
                     ("自動偵測", self.auto_detect), ("自動求斜率", self.auto_shear_btn),
                     ("⇄ 反向", self.swap_ends), ("聚焦到線", self.focus_line), ("看全圖", self.show_full)]:
            ttk.Button(bar, text=t, command=c).pack(side="left", padx=2)
        ttk.Label(bar, text="  檢視照片：").pack(side="left")
        self.cmb_view = ttk.Combobox(bar, textvariable=self.var_view, state="readonly", width=18, values=[])
        self.cmb_view.pack(side="left")
        self.cmb_view.bind("<<ComboboxSelected>>", lambda e: self.show_photo(self.cmb_view.current()))
        ttk.Button(bar, text="◀", width=3, command=lambda: self.step_photo(-1)).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="▶", width=3, command=lambda: self.step_photo(1)).pack(side="left")
        ttk.Checkbutton(bar, text="左鍵＝平移畫面", variable=self.var_hand).pack(side="left", padx=10)
        sl = ttk.Frame(f)
        sl.pack(fill="x", padx=6)
        geom = lambda: self._later("geom", self.on_geom, 150)
        self._slider(sl, "裁切寬度 (px)　1＝單線", self.var_width, 1, 500, geom, 0, 1)
        self._slider(sl, "斜率修正 (度)　寬度>1才有用", self.var_shear, -15, 15, geom, 1, 0.05)
        self._slider(sl, "顯示亮度（只影響畫面）", self.var_gain, 1, 80, lambda: self._later("gain", self.apply_gain, 80), 2, 0.5)
        ttk.Label(f, text="滾輪＝縮放　右鍵/中鍵拖曳＝平移　拖曳綠點(起點＝短波長)/紅點(終點)/線身＝調整　"
                          "空白處左鍵拖曳＝畫新線　點選後方向鍵微調（Shift 大步、Ctrl 極細）",
                  style="Hint.TLabel").pack(anchor="w", padx=8)
        ttk.Label(f, textvariable=self.var_hover, anchor="w").pack(fill="x", padx=8)
        pw = ttk.PanedWindow(f, orient="vertical")
        pw.pack(fill="both", expand=True)
        top, bot = ttk.Frame(pw), ttk.Frame(pw)
        pw.add(top, weight=3)
        pw.add(bot, weight=2)
        self.view = ZoomView(top)
        self.view.pack(fill="both", expand=True)
        self.view.draw_overlay = self.draw_ov1
        self.view.hook_press = self.v1_press
        self.view.hook_motion = self.v1_motion
        self.view.hook_release = self.v1_release
        self.view.hover_cb = self.v1_hover
        self.view.bind("<Key>", self.v1_key)
        self.fig1 = Figure(figsize=(10, 3), constrained_layout=True)
        self.ax_strip, self.ax_prof1 = self.fig1.subplots(2, 1, sharex=True, gridspec_kw={"height_ratios": [1, 2]})
        self.canvas1, self.tb1 = self._mpl(bot, self.fig1)

    def load_neon_path(self, p):
        img = self.try_load(p)
        if img is None:
            return
        self.photos = [{"name": "氖燈", "path": p, "img": img}]
        self.p0 = self.p1 = None
        self.strips = []
        self.prof = None
        self.coef = None
        self.calib = None
        self.pk_x, self.pk_h, self.pk_ref = np.array([]), np.array([]), []
        self.prof_dirty = True
        self.init_scale = True
        self.var_width.set(1.0)
        self.var_shear.set(0.0)
        self.photos[0]["gain"] = self.auto_gain(img)
        self.var_gain.set(self.photos[0]["gain"])
        self.view_idx = 0
        self.view.set_image(img, fget(self.var_gain, 1.0), bright=True, fit=True)
        self.update_view_list()
        self.auto_detect()

    def open_neon(self):
        p = filedialog.askopenfilename(title="選擇氖燈照片", filetypes=IMG_TYPES, initialdir=self._d("photo"))
        if p:
            self._rem("photo", p)
            self.load_neon_path(p)

    @staticmethod
    def auto_gain(img):
        top = float(np.percentile(img[::4, ::4].max(axis=2), 99.99))
        return round(min(80.0, max(1.0, 255.0 / max(top, 30.0))), 1)

    def update_view_list(self):
        names = [f"{i + 1} {ph['name']}" for i, ph in enumerate(self.photos)]
        self.cmb_view.configure(values=names)
        if self.photos:
            self.view_idx = min(self.view_idx, len(self.photos) - 1)
            self.var_view.set(names[self.view_idx])

    def show_photo(self, i):
        """切換第 1 頁上方顯示哪一張照片（裁切線是共用的，所以可以拿任何一張來檢查）"""
        if not self.photos or i < 0 or i >= len(self.photos):
            return
        self.view_idx = i
        ph = self.photos[i]
        self.var_gain.set(ph.get("gain", 1.0))
        self.view.set_image(ph["img"], fget(self.var_gain, 1.0), bright=True, fit=False)
        self.update_view_list()

    def step_photo(self, d):
        if self.photos:
            self.show_photo((self.view_idx + d) % len(self.photos))

    def add_sample_path(self, p):
        img = self.try_load(p)
        if img is None:
            return
        if img.shape != self.photos[0]["img"].shape:
            messagebox.showwarning("尺寸不同", f"{os.path.basename(p)} 的尺寸和氖燈照片不同，無法共用同一條線，已略過。")
            return
        self.photos.append({"name": os.path.splitext(os.path.basename(p))[0], "path": p, "img": img,
                            "gain": self.auto_gain(img)})

    def add_samples(self):
        if not self.photos:
            messagebox.showinfo("提示", "請先開啟氖燈照片。")
            return
        ps = filedialog.askopenfilenames(title="選擇其他光譜照片（可多選）", filetypes=IMG_TYPES, initialdir=self._d("photo"))
        if ps:
            self._rem("photo", ps[0])
        for p in ps:
            self.add_sample_path(p)
        self.update_view_list()
        if ps and len(self.photos) > 1:
            self.show_photo(len(self.photos) - 1)      # 直接切到剛加入的照片讓你檢查裁切線
        self.refresh_strips()

    def auto_detect(self):
        if not self.photos:
            messagebox.showinfo("提示", "請先開啟氖燈照片。")
            return
        try:
            p0, p1, wsup, res = auto_line_multi([ph["img"] for ph in self.photos])
        except ValueError as e:
            messagebox.showwarning("自動偵測", str(e))
            return
        sts = []
        for ph in self.photos:      # 每張各自正規化後疊在一起判斷方向，避免亮的照片蓋過暗的氖燈
            st = extract_strip(ph["img"], p0, p1, min(wsup, 300.0), 0.0).astype(np.float32)
            sts.append(st / max(float(st.mean(axis=(0, 2)).max()), 1e-6))
        if auto_flip(np.concatenate(sts, axis=0)):
            p0, p1 = p1, p0
        miss = [self.photos[i]["name"] for i, r in enumerate(res) if r is None]
        self.p0, self.p1 = p0, p1
        self.var_shear.set(0.0)
        self.focus_line()
        self.refresh_strips()
        _, _, L = self.basis()
        self.status.set(f"自動偵測（合併 {len(self.photos) - len(miss)}/{len(self.photos)} 張照片，線涵蓋所有光譜範圍）：線長 {L:.0f} px。"
                        + (f"沒偵測到：{'、'.join(miss)}（已略過）。" if miss else "") + f"亮線橫向約 {wsup:.0f} px，若覺得曲線雜訊多，可把寬度加到 {max(1, int(wsup * 0.6))} 左右"
                        "（寬度 1＝單線）。請檢查綠點是短波長端(綠光側)；不準就拖曳或放大後微調。")

    def auto_shear_btn(self):
        if self.p0 is None:
            return
        if self.cur_width() < 3:
            messagebox.showinfo("提示", "寬度為 1（單線）時不需要斜率修正。要用的話請先把寬度調到 3 以上。")
            return
        self.var_shear.set(round(auto_shear(self.photos[0]["img"], self.p0, self.p1, self.cur_width()), 2))
        self.on_geom()

    def swap_ends(self):
        if self.p0 is None:
            return
        self.p0, self.p1 = self.p1, self.p0
        self.view.redraw_overlay()
        self.refresh_strips()

    def focus_line(self):
        if self.p0 is None:
            return
        pts = np.array([self.p0, self.p1])
        L = float(np.hypot(*(self.p1 - self.p0)))
        m = max(40.0, 0.2 * L) + self.cur_width() / 2.0
        x0, y0 = pts.min(axis=0) - m
        x1, y1 = pts.max(axis=0) + m
        self.view.zoom_to_rect(x0, y0, x1, y1)

    def show_full(self):
        self.view.fit()

    def apply_gain(self):
        if self.photos:
            self.photos[self.view_idx]["gain"] = fget(self.var_gain, 1.0)   # 每張照片各記自己的顯示亮度
        self.view.set_gain(fget(self.var_gain, 1.0))

    def on_geom(self):
        self.view.redraw_overlay()
        self.refresh_strips()

    # ---- 影像上的疊圖與滑鼠操作 ----
    def draw_ov1(self, v):
        if self.p0 is None or self.p1 is None:
            return
        e, n, L = self.basis()
        t = math.tan(math.radians(fget(self.var_shear, 0)))
        f = n + t * e
        nv = self.cur_width()
        a = v.to_screen(*self.p0)
        b = v.to_screen(*self.p1)
        if nv > 1:
            hw = nv / 2.0
            c = [v.to_screen(*(self.p0 - hw * f)), v.to_screen(*(self.p0 + hw * f)),
                 v.to_screen(*(self.p1 + hw * f)), v.to_screen(*(self.p1 - hw * f))]
            v.create_polygon(*[q for pt in c for q in pt], outline="#ffd600", fill="", tags="ov")
        v.create_line(a[0], a[1], b[0], b[1], fill="#00ff90", tags="ov")
        r, o = 7 * self.k, 11 * self.k
        v.create_oval(a[0] - r, a[1] - r, a[0] + r, a[1] + r, fill="#00e676", outline="black", tags="ov")
        v.create_rectangle(b[0] - r, b[1] - r, b[0] + r, b[1] + r, fill="#ff1744", outline="black", tags="ov")
        v.create_text(a[0] + o, a[1] - o, text="起點(短波長)", fill="#00e676", anchor="w", tags="ov")
        v.create_text(b[0] + o, b[1] + o, text="終點(長波長)", fill="#ff5c7a", anchor="w", tags="ov")

    def hit(self, sx, sy):
        v = self.view
        a = np.array(v.to_screen(*self.p0))
        b = np.array(v.to_screen(*self.p1))
        m = np.array([sx, sy], float)
        if np.hypot(*(m - a)) <= 12 * self.k:
            return "p0"
        if np.hypot(*(m - b)) <= 12 * self.k:
            return "p1"
        d = b - a
        l2 = float(d @ d)
        if l2 > 1:
            u = float(np.clip(((m - a) @ d) / l2, 0, 1))
            if np.hypot(*(m - (a + u * d))) <= 7 * self.k:
                return "body"
        return None

    def v1_press(self, ev):
        self.view.focus_set()
        if not self.photos:
            return True
        pt = np.array(self.view.to_image(ev.x, ev.y))
        h = self.hit(ev.x, ev.y) if self.p0 is not None else None
        if h:
            self.sel = h
            self.drag = {"mode": h, "start": pt, "p0": self.p0.copy(), "p1": self.p1.copy()}
            return True
        if self.var_hand.get():
            return False
        self.drag = {"mode": "new", "start": pt,
                     "p0": None if self.p0 is None else self.p0.copy(),
                     "p1": None if self.p1 is None else self.p1.copy()}
        self.p0, self.p1 = pt.copy(), pt.copy()
        self.sel = "p1"
        return True

    def v1_motion(self, ev):
        d = self.drag
        if not d:
            return
        pt = np.array(self.view.to_image(ev.x, ev.y))
        m = d["mode"]
        if m == "p0":
            self.p0 = pt
        elif m in ("p1", "new"):
            self.p1 = pt
        else:
            dv = pt - d["start"]
            self.p0, self.p1 = d["p0"] + dv, d["p1"] + dv
        self.view.redraw_overlay()
        self.show_line_info()

    def v1_release(self, ev):
        d = self.drag
        self.drag = None
        if not d:
            return
        if d["mode"] == "new" and float(np.hypot(*(self.p1 - self.p0))) < 5:
            self.p0, self.p1 = d["p0"], d["p1"]
            self.view.redraw_overlay()
            return
        self.refresh_strips()

    def v1_key(self, ev):
        if self.p0 is None:
            return
        ks = ev.keysym
        if ks not in ("Left", "Right", "Up", "Down"):
            return
        shift, ctrl = bool(ev.state & 0x1), bool(ev.state & 0x4)
        st = 2.0 if shift else (0.05 if ctrl else 0.25)
        dv = np.array({"Left": (-st, 0.0), "Right": (st, 0.0), "Up": (0.0, -st), "Down": (0.0, st)}[ks])
        if self.sel == "p0":
            self.p0 = self.p0 + dv
        elif self.sel == "p1":
            self.p1 = self.p1 + dv
        else:
            self.p0, self.p1 = self.p0 + dv, self.p1 + dv
        self.view.redraw_overlay()
        self.show_line_info()
        self._later("nudge", self.refresh_strips, 250)

    def show_line_info(self):
        if self.p0 is None:
            return
        e, _, L = self.basis()
        ang = math.degrees(math.atan2(e[1], e[0]))
        self.var_hover.set(f"起點 ({self.p0[0]:.2f}, {self.p0[1]:.2f})　終點 ({self.p1[0]:.2f}, {self.p1[1]:.2f})　"
                           f"線長 {L:.1f} px　角度 {ang:.2f}°　（目前選取：{ {'p0': '起點', 'p1': '終點', 'body': '整條線'}[self.sel] }）")

    def v1_hover(self, x, y):
        if self.drag or not self.photos:
            return
        img = self.photos[self.view_idx]["img"]
        xi, yi = int(round(x)), int(round(y))
        if 0 <= xi < img.shape[1] and 0 <= yi < img.shape[0]:
            r, g, b = (int(c) for c in img[yi, xi])
            self.var_hover.set(f"游標 ({x:.1f}, {y:.1f})　像素值 R{r} G{g} B{b}")

    def refresh_strips(self):
        ax1, ax2 = self.ax_strip, self.ax_prof1
        ax1.clear()
        ax2.clear()
        self.strips = []
        self.prof_dirty = True
        self.coef = None
        self.align.dirty = True
        if self.p0 is None or not self.photos:
            self.canvas1.draw_idle()
            return
        _, _, L = self.basis()
        w = self.cur_width()
        s = fget(self.var_shear, 0) if w >= 3 else 0.0
        rows, centers, y = [], [], 0
        for ph in self.photos:
            st = extract_strip(ph["img"], self.p0, self.p1, w, s)
            self.strips.append(st)
            m = max(float(np.percentile(st, 99.8)), 20.0)
            im = np.clip(st / m, 0, 1) ** 0.5
            im = np.repeat(im, max(1, int(math.ceil(10 / im.shape[0]))), axis=0)
            rows.append(im)
            centers.append(y + im.shape[0] / 2.0)
            y += im.shape[0]
            rows.append(np.ones((2, im.shape[1], 3), np.float32))
            y += 2
        img = np.concatenate(rows[:-1], axis=0)
        ax1.imshow(img, aspect="auto", extent=(-0.5, img.shape[1] - 0.5, img.shape[0], 0), interpolation="nearest")
        ax1.set_yticks(centers)
        ax1.set_yticklabels([ph["name"][:14] for ph in self.photos])
        for ph, st in zip(self.photos, self.strips):
            p = strip_to_profile(st, self.var_mode.get(), self.var_degamma.get())
            ax2.plot(np.arange(len(p)), p, lw=0.9, label=ph["name"][:14])
        ax2.set_xlabel("沿裁切線的位置 (px)　（每張光譜都用同一條線）")
        ax2.set_ylabel("灰階值 (0~255，同 ImageJ)")
        ax2.grid(True, alpha=0.3)
        ax2.legend(fontsize=8, loc="upper right")
        m = self.strips[0].max(axis=2)
        sat = float((m >= 250).sum()) / max(1, int((m >= 80).sum()))
        self.status.set(f"裁切線長 {L:.0f} px、寬 {w} px；氖燈亮處的飽和比例 {sat * 100:.0f}%"
                        + ("　⚠ 偏高，峰位置會不準，建議降低曝光重拍" if sat > 0.15 else ""))
        self.show_line_info()
        self.canvas1.draw_idle()

    # ======================================================================
    #  第 2 頁：對圖（樣本當底圖）
    # ======================================================================
    def _build2(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text=" 2 對圖（樣本當底圖） ")
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=6, pady=4)
        for t, c in [("載入參考圖", self.load_reference), ("重新校準參考圖橫軸", self.calibrate_reference),
                     ("自動偵測對齊", self.auto_align_base), ("自動粗對", self.do_coarse),
                     ("兩點對齊", lambda: self.align.start_pick()),
                     ("重設視窗", lambda: (self.align.fit(), None))]:
            ttk.Button(bar, text=t, command=c).pack(side="left", padx=2)
        ttk.Label(bar, text="  峰值靈敏度(%)").pack(side="left")
        ttk.Spinbox(bar, from_=0.2, to=50, increment=0.5, textvariable=self.var_sens, width=6,
                    command=self.on_prof_param).pack(side="left")
        ttk.Label(bar, text="  亮度算法").pack(side="left")
        cb = ttk.Combobox(bar, textvariable=self.var_mode, values=["ImageJ(RGB平均)", "亮度(加權)", "最大通道", "RGB總和"], width=15, state="readonly")
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self.on_prof_param())
        ttk.Checkbutton(bar, text="去 gamma", variable=self.var_degamma, command=self.on_prof_param).pack(side="left", padx=6)
        bar2 = ttk.Frame(f)
        bar2.pack(fill="x", padx=6)
        ttk.Label(bar2, text="左鍵拖曳：").pack(side="left")
        ttk.Radiobutton(bar2, text="移動我的光譜", variable=self.var_drag, value="spec").pack(side="left")
        ttk.Radiobutton(bar2, text="平移畫面", variable=self.var_drag, value="pan").pack(side="left")
        r = ttk.Frame(f)
        r.pack(fill="x", padx=6)
        ttk.Label(r, text="參考波長 (nm，逗號分隔)").pack(side="left")
        en = ttk.Entry(r, textvariable=self.var_refs)
        en.pack(side="left", fill="x", expand=True, padx=4)
        en.bind("<Return>", lambda ev: self.draw_align())
        ttk.Button(r, text="套用清單", command=self.draw_align).pack(side="left")
        ttk.Checkbutton(r, text="兩點對齊時吸附到亮線／參考線", variable=self.var_snap).pack(side="left", padx=8)
        sl = ttk.Frame(f)
        sl.pack(fill="x", padx=6)
        upd = lambda: self.align.schedule()
        self._slider(sl, "平移：中心波長 (nm)", self.var_center, 250, 900, upd, 0, 0.1)
        self._slider(sl, "縮放：條帶總寬度 (nm)", self.var_span, 10, 1000, upd, 1, 0.1)
        self._slider(sl, "我的光譜透明度（0＝只看樣本）", self.var_alpha, 0, 1, upd, 2, 0.05)
        ttk.Label(f, text="拖曳左右兩端的橘色把手＝拉伸／移動我的光譜(光譜在兩把手中間)　滾輪＝縮放畫面　Shift＋滾輪＝縮放我的光譜　右鍵/中鍵拖曳＝平移畫面　"
                          "點一下圖後：←→＝移動我的光譜(Shift 大步、Ctrl 極細)　↑↓＝縮放我的光譜",
                  style="Hint.TLabel").pack(anchor="w", padx=8)
        ttk.Label(f, textvariable=self.var_cur, anchor="w").pack(fill="x", padx=8)
        self.align = AlignView(f, self)
        self.align.pack(fill="both", expand=True)

    def compute_profile(self, force=False):
        if not self.strips:
            return False
        if not (force or self.prof_dirty):
            return True
        self.prof = strip_to_profile(self.strips[0], self.var_mode.get(), self.var_degamma.get())
        self.pk_x, self.pk_h = find_peaks(self.prof, thr_frac=fget(self.var_sens, 2.0) / 100.0)
        self.pk_ref = [None] * len(self.pk_x)
        self.prof_dirty = False
        self.coef = None
        if self.init_scale:
            self.init_scale = False
            self.var_span.set(round(0.6 * (len(self.prof) - 1), 1))
        return True

    def lam_lin(self, x):
        """第 2 頁滑桿決定的「位置 -> 波長」（直線）"""
        L = max(1, len(self.prof) - 1)
        span, c = fget(self.var_span, 200.0), fget(self.var_center, 600.0)
        return (c - span / 2.0) + span * np.asarray(x, float) / L

    def lam_final(self, x):
        return np.polyval(self.coef, np.asarray(x, float)) if self.coef is not None else self.lam_lin(x)

    def on_prof_param(self):
        self.prof_dirty = True
        if self.compute_profile():
            self.draw_align()

    def _parse_refs(self, txt):
        txt = txt.replace("，", ",").replace(";", ",").replace("、", ",")
        out = []
        for tok in txt.replace("\n", ",").replace(" ", ",").split(","):
            if tok.strip():
                try:
                    out.append(float(tok))
                except ValueError:
                    pass
        return sorted(set(out))

    def get_refs(self):
        txt = self.var_refs.get().replace("，", ",").replace(";", ",").replace("、", ",")
        out = []
        for tok in txt.replace("\n", ",").replace(" ", ",").split(","):
            if tok.strip():
                try:
                    out.append(float(tok))
                except ValueError:
                    messagebox.showwarning("參考波長", f"「{tok}」不是數字，已略過。")
        return sorted(set(out))

    def load_reference(self):
        p = filedialog.askopenfilename(title="選擇網路上的氖燈參考圖（樣本）", filetypes=IMG_TYPES,
                                       initialdir=self._d("ref") or self._d("photo"))
        if not p:
            return
        self.load_reference_path(p)

    def load_reference_path(self, p):
        try:
            self.ref_img = np.asarray(Image.open(p).convert("RGB"))
        except Exception as e:
            messagebox.showerror("開啟失敗", str(e))
            return
        self.ref_map = None
        self.ref_path = p
        self._rem("ref", p)
        self.base_pk_lam = np.array([])
        self.align.dirty = True
        self.calibrate_reference()

    def calibrate_reference(self):
        if self.ref_img is None:
            self.load_reference()
            return
        top = tk.Toplevel(self.root)
        top.title("校準參考圖橫軸")
        top.geometry("1250x650")
        ttk.Label(top, text="請在圖上點兩個「已知波長」的位置（兩點離越遠越準）。滾輪可放大、右鍵拖曳平移，"
                            "放大後再點最準；每點一下會問你波長。").pack(anchor="w", padx=6, pady=4)
        zv = ZoomView(top)
        zv.pack(fill="both", expand=True)
        zv.set_image(self.ref_img, 1.0, bright=False, fit=True)
        pts = []

        def ov(v):
            for x, y, lam in pts:
                sx, sy = v.to_screen(x, y)
                v.create_line(sx - 12, sy, sx + 12, sy, fill="#ff3b3b", width=2, tags="ov")
                v.create_line(sx, sy - 12, sx, sy + 12, fill="#ff3b3b", width=2, tags="ov")
                v.create_text(sx + 8, sy - 14, text=f"{lam:g} nm", fill="#ffe082", anchor="w", tags="ov")

        def onclick(ix, iy):
            lam = simpledialog.askfloat("波長", f"第 {len(pts) + 1} 個點是幾 nm？", parent=top)
            if lam is None:
                return
            pts.append((ix, iy, lam))
            zv.redraw_overlay()
            if len(pts) == 2:
                (xa, _, la), (xb, _, lb) = sorted(pts, key=lambda q: q[0])
                xa, xb = xa + 0.5, xb + 0.5      # 取樣座標 -> 像素邊界座標
                if abs(xb - xa) < 5 or la == lb:
                    messagebox.showwarning("校準", "兩點太近或波長相同，請重新點。", parent=top)
                    pts.clear()
                    zv.redraw_overlay()
                    return
                if lb < la:
                    messagebox.showwarning("校準", "右邊的點波長比左邊小，這張圖的波長方向反了，目前不支援。", parent=top)
                    pts.clear()
                    zv.redraw_overlay()
                    return
                self.ref_map = (xa, la, xb, lb)
                top.destroy()
                self.align.dirty = True
                self.draw_align()
                self.save_settings_now()          # 底圖路徑與橫軸校準立刻存進 SpectrumSetting.json

        zv.draw_overlay = ov
        zv.click_cb = onclick

    def ref_lam(self, px):
        x1, l1, x2, l2 = self.ref_map
        return l1 + (px - x1) * (l2 - l1) / (x2 - x1)

    def draw_align(self):
        if not self.compute_profile():
            return
        self.refs_cache = self.get_refs()
        self.update_base_peaks()
        self.align.dirty = True
        if not self.align.fitted:
            self.align.fit()
        self.align.schedule()
        self.status.set(f"偵測到 {len(self.pk_x)} 條亮線（紅色三角），參考清單 {len(self.refs_cache)} 條（綠色虛線）。"
                        "把紅三角拖到綠虛線上，對齊後到第 3 頁按「自動比對」。")

    def update_profile_artists(self):
        self.align.schedule()

    def do_coarse(self):
        if not self.compute_profile():
            messagebox.showinfo("提示", "請先在第 1 頁開啟照片並設好裁切線。")
            return
        refs = self.get_refs()
        if len(self.pk_x) < 3 or len(refs) < 3:
            messagebox.showinfo("提示", "偵測到的亮線或參考波長少於 3 個，無法自動粗對。")
            return
        self.status.set("自動粗對中……")
        self.root.update_idletasks()
        r = coarse_align(self.pk_x, self.pk_h, refs, len(self.prof) - 1)
        self.var_center.set(round(min(900.0, max(250.0, r["center"])), 2))
        self.var_span.set(round(min(1000.0, max(10.0, r["span"])), 2))
        self.align.schedule()
        self.status.set(f"粗對完成：吻合分數 {r['score']:.2f}（1＝全部吻合；低於 0.5 通常不可信）。請看圖確認，必要時再手動微調。")

    # ======================================================================
    #  第 3 頁：比對與校正
    # ======================================================================
    def _build3(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text=" 3 比對與校正 ")
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=6, pady=4)
        ttk.Button(bar, text="自動比對", command=self.auto_match).pack(side="left", padx=2)
        ttk.Label(bar, text="容許誤差(nm)").pack(side="left")
        ttk.Spinbox(bar, from_=0.2, to=10, increment=0.1, textvariable=self.var_tol, width=6).pack(side="left", padx=2)
        ttk.Label(bar, text="  擬合階數").pack(side="left")
        ttk.Combobox(bar, textvariable=self.var_deg, values=["1", "2"], width=4, state="readonly").pack(side="left", padx=2)
        ttk.Button(bar, text="計算校正", command=self.calc_calib).pack(side="left", padx=6)
        ttk.Button(bar, text="刪除選取的峰", command=self.del_peak).pack(side="left", padx=2)
        ttk.Label(bar, text="　雙擊表格列＝手動指定波長（留空＝不使用）").pack(side="left")
        self.lbl_res = tk.StringVar(value="尚未計算校正")
        ttk.Label(f, textvariable=self.lbl_res, font=self.f_res).pack(anchor="w", padx=8)
        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)
        cols = ("no", "x", "rough", "ref", "fit", "res")
        self.tree = ttk.Treeview(body, columns=cols, show="headings", height=24)
        for c, t, w in zip(cols, ["#", "位置(px)", "概略波長", "比對波長", "校正波長", "殘差(nm)"], [40, 85, 90, 90, 90, 85]):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center")
        self.tree.pack(side="left", fill="y", padx=6, pady=4)
        self.tree.bind("<Double-1>", self.edit_match)
        self.fig3 = Figure(figsize=(6, 5), constrained_layout=True)
        self.fig3.add_subplot(111)
        self.canvas3 = FigureCanvasTkAgg(self.fig3, master=body)
        self.canvas3.get_tk_widget().pack(side="left", fill="both", expand=True)
        debounce_mpl(self.canvas3)

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        if self.prof is None:
            return
        rough = self.lam_lin(self.pk_x)
        fit = self.lam_final(self.pk_x) if self.coef is not None else None
        for i, x in enumerate(self.pk_x):
            ref = self.pk_ref[i]
            fit_s = f"{fit[i]:.2f}" if fit is not None else ""
            res_s = f"{ref - fit[i]:+.3f}" if (fit is not None and ref is not None) else ""
            self.tree.insert("", "end", iid=str(i),
                             values=(i + 1, f"{x:.2f}", f"{rough[i]:.1f}", "" if ref is None else f"{ref:.2f}", fit_s, res_s))

    def auto_match(self):
        if not self.compute_profile():
            return
        refs = self.get_refs()
        if not refs:
            messagebox.showinfo("提示", "參考波長清單是空的。")
            return
        rough = self.lam_final(self.pk_x)
        idx = match_peaks(rough, refs, fget(self.var_tol, 1.5))
        self.pk_ref = [refs[j] if j >= 0 else None for j in idx]
        self.coef = None
        self.refresh_table()
        n = sum(r is not None for r in self.pk_ref)
        self.status.set(f"自動比對：{len(self.pk_x)} 個峰中配到 {n} 個。請檢查表格；配錯的雙擊修改，淡的雜訊峰可刪除。")

    def edit_match(self, ev):
        sel = self.tree.identify_row(ev.y)
        if not sel:
            return
        i = int(sel)
        cur = self.pk_ref[i]
        v = simpledialog.askstring("手動指定", f"第 {i + 1} 個峰（位置 {self.pk_x[i]:.1f} px）\n對應的波長 (nm)，留空＝不使用：",
                                   initialvalue="" if cur is None else f"{cur}", parent=self.root)
        if v is None:
            return
        v = v.strip()
        try:
            self.pk_ref[i] = float(v) if v else None
        except ValueError:
            messagebox.showwarning("輸入錯誤", "請輸入數字。")
            return
        self.coef = None
        self.refresh_table()

    def del_peak(self):
        sel = self.tree.selection()
        if not sel:
            return
        drop = sorted((int(s) for s in sel), reverse=True)
        keep = [i for i in range(len(self.pk_x)) if i not in drop]
        self.pk_x, self.pk_h = self.pk_x[keep], self.pk_h[keep]
        self.pk_ref = [self.pk_ref[i] for i in keep]
        self.coef = None
        self.refresh_table()
        self.align.schedule()

    def calc_calib(self):
        if not self.compute_profile():
            return
        idx = [i for i, r in enumerate(self.pk_ref) if r is not None]
        deg = int(self.var_deg.get())
        if len(idx) < deg + 2:
            messagebox.showwarning("點數不足", f"{deg} 階擬合至少要 {deg + 2} 個已比對的峰（目前 {len(idx)} 個）。")
            return
        xs = self.pk_x[idx]
        ls = np.array([self.pk_ref[i] for i in idx], float)
        c, res = fit_poly(xs, ls, deg)
        rms = float(np.sqrt(np.mean(res ** 2)))
        loo = loo_rms(xs, ls, deg)
        self.coef = c
        self.calib = {"deg": deg, "coef_highest_first": [float(v) for v in c], "n": len(idx), "rms_nm": rms,
                      "loo_rms_nm": loo, "max_abs_nm": float(np.abs(res).max()),
                      "points": [{"x_px": float(x), "ref_nm": float(l), "res_nm": float(r)} for x, l, r in zip(xs, ls, res)]}
        self.refresh_table()
        ax = self.fig3.axes[0]
        ax.clear()
        ax.axhline(0, color="gray", lw=0.8)
        ax.plot(np.polyval(c, xs), res, "o")
        ax.set_xlabel("波長 (nm)")
        ax.set_ylabel("殘差＝參考值 − 擬合值 (nm)")
        ax.set_title("校正殘差（越貼近 0、越沒有彎曲規律越好）", fontsize=10)
        ax.grid(True, alpha=0.3)
        self.canvas3.draw_idle()
        self.lbl_res.set(f"N={len(idx)}　RMS={rms:.3f} nm　最大誤差={np.abs(res).max():.3f} nm　留一法 RMS={loo:.3f} nm")
        self.status.set("校正完成。若殘差有明顯彎曲規律，可改用 2 階；也可再按一次「自動比對」用校正結果抓回更多亮線。")

    # ======================================================================
    #  第 4 頁：輸出
    # ======================================================================
    def _build4(self):
        f = ttk.Frame(self.nb)
        self.nb.add(f, text=" 4 輸出 ")
        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=6, pady=4)
        for t, c in [("更新圖", self.draw_output), ("存 PNG", self.save_png), ("存 CSV", self.save_csv),
                     ("存校正記錄 JSON", self.save_json)]:
            ttk.Button(bar, text=t, command=c).pack(side="left", padx=2)
        ttk.Label(bar, text="  輸出 DPI").pack(side="left")
        ttk.Spinbox(bar, from_=72, to=1200, increment=50, textvariable=self.var_dpi, width=6).pack(side="left")
        for t, v in [("正規化(最大=1)", self.var_norm), ("扣背景", self.var_bg), ("標示峰值波長", self.var_label)]:
            ttk.Checkbutton(bar, text=t, variable=v, command=self.draw_output).pack(side="left", padx=8)
        self.fig4 = Figure(figsize=(10, 6), constrained_layout=True)
        self.canvas4, self.tb4 = self._mpl(f, self.fig4)

    def output_data(self):
        x = np.arange(self.strips[0].shape[1])
        lam = np.polyval(self.coef, x)
        out = []
        for ph, st in zip(self.photos, self.strips):
            p = strip_to_profile(st, self.var_mode.get(), self.var_degamma.get())
            if self.var_bg.get():
                p = p - np.percentile(p, 5)
            if self.var_norm.get():
                p = p / (p.max() or 1.0)
            out.append((ph["name"], p))
        return x, lam, out

    def draw_output(self):
        fig = self.fig4
        fig.clear()
        ax = fig.add_subplot(111)
        if self.coef is None or not self.strips:
            ax.text(0.5, 0.5, "請先在第 3 頁完成「計算校正」", ha="center", va="center", transform=ax.transAxes)
            self.canvas4.draw_idle()
            return
        x, lam, out = self.output_data()
        for name, p in out:
            ax.plot(lam, p, lw=1, label=name)
            if self.var_label.get():
                px, _ = find_peaks(p, thr_frac=max(0.05, fget(self.var_sens, 2.0) / 100.0))
                for q in px:
                    ax.annotate(f"{np.polyval(self.coef, q):.1f}", (np.polyval(self.coef, q), np.interp(q, x, p)),
                                rotation=90, fontsize=7, ha="center", va="bottom")
        ax.set_xlabel("波長 (nm)")
        ax.set_ylabel("相對亮度" if self.var_norm.get() else "灰階值 (0~255，同 ImageJ)")
        ax.grid(True, alpha=0.3)
        ax.legend()
        self.canvas4.draw_idle()

    def _need_calib(self):
        if self.coef is None or not self.strips:
            messagebox.showinfo("提示", "請先在第 3 頁完成「計算校正」。")
            return False
        return True

    def save_png(self):
        if not self._need_calib():
            return
        self.draw_output()
        p = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")], initialdir=self._d("out"))
        self._rem("out", p)
        if p:
            self.fig4.savefig(p, dpi=int(fget(self.var_dpi, 300)))
            self.status.set(f"已存：{p}（DPI {int(fget(self.var_dpi, 300))}）")

    def save_csv(self):
        if not self._need_calib():
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialdir=self._d("out"))
        self._rem("out", p)
        if not p:
            return
        x, lam, out = self.output_data()
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["wavelength_nm", "pixel"] + [n for n, _ in out])
            for i in range(len(x)):
                w.writerow([f"{lam[i]:.4f}", int(x[i])] + [f"{pr[i]:.5f}" for _, pr in out])
        self.status.set(f"已存：{p}")

    def save_json(self):
        if not self._need_calib():
            return
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")], initialdir=self._d("out"))
        self._rem("out", p)
        if not p:
            return
        data = {"tool": "spectrum_mode",
                "photos": [{"name": ph["name"], "path": ph["path"]} for ph in self.photos],
                "line": {"p0": [float(v) for v in self.p0], "p1": [float(v) for v in self.p1],
                         "width_px": self.cur_width(), "shear_deg": fget(self.var_shear)},
                "profile_mode": self.var_mode.get(), "degamma": bool(self.var_degamma.get()),
                "reference_nm": self.get_refs(), "calibration": self.calib}
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        self.status.set(f"已存：{p}")

    # ---------------- 切換分頁 ----------------
    def on_tab(self, _e=None):
        idx = self.nb.index("current")
        if idx >= 1:
            if not self.strips:
                messagebox.showinfo("提示", "請先在第 1 頁開啟照片，並設好裁切線。")
                self.nb.select(0)
                return
            self.compute_profile()
        if idx == 1:
            self.draw_align()
        elif idx == 2:
            self.refresh_table()
        elif idx == 3:
            self.draw_output()