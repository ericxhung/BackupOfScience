import io
import json
import os
import re
import string
from datetime import datetime
from flask import (
    Flask,
    abort,
    jsonify,
    make_response,
    render_template,
    request,
    send_file,
)
import markdown
import openpyxl
import pandas as pd

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(CURRENT_DIR, "templates")
ROOT_DIR = os.path.dirname(CURRENT_DIR)  # C:\Users\USER\Documents\科資
RESISTOR_DIR = os.path.join(ROOT_DIR, "電阻")

app = Flask(__name__, template_folder=TEMPLATE_DIR)


def parse_date_from_string(text):
  match = re.search(r"(\d{4})_(\d{1,2})_(\d{1,2})", text)
  if match:
    y, m, d = match.groups()
    try:
      dt = datetime(int(y), int(m), int(d))
      return dt.strftime("%Y-%m-%d")
    except ValueError:
      pass
  return None


def get_all_resistor_data(mode="RESISTOR"):
  file_list = []
  if os.path.exists(RESISTOR_DIR):
    for root, dirs, files in os.walk(RESISTOR_DIR):
      for file in files:
        if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
          f_path = os.path.join(root, file)
          folder_name = os.path.basename(root)
          date_str = parse_date_from_string(file) or parse_date_from_string(
              folder_name
          )
          if not date_str:
            date_str = "2026-01-01"
          file_list.append({"path": f_path, "filename": file, "date": date_str})

  if not file_list:
    return pd.DataFrame()

  setting_path = os.path.join(ROOT_DIR, "setting.json")
  item_mapping = {}
  if os.path.exists(setting_path):
    try:
      with open(setting_path, "r", encoding="utf-8") as sf:
        item_mapping = json.load(sf)
    except Exception:
      pass

  if not item_mapping:
    all_sheets = set()
    for f in file_list:
      try:
        wb = openpyxl.load_workbook(f["path"], read_only=True, data_only=True)
        for sname in wb.sheetnames:
          if (
              sname not in ["氣溫", "溫度"]
              and "plxdaq" not in sname.lower()
          ):
            all_sheets.add(sname)
        wb.close()
      except Exception:
        pass
    item_mapping = {sname: [sname] for sname in sorted(list(all_sheets))}

  matrix_data = {}
  cols = list(string.ascii_uppercase[:16])
  rows = range(1, 31)
  pattern = (
      re.compile(r"^=\d+\*")
      if mode == "RESISTOR"
      else re.compile(r"^=AVERAGE\(", re.IGNORECASE)
  )

  for f in file_list:
    date_key = f["date"]
    if date_key not in matrix_data:
      matrix_data[date_key] = {}

    try:
      wb_f = openpyxl.load_workbook(f["path"], data_only=False)
      wb_d = openpyxl.load_workbook(f["path"], data_only=True)

      for item_name, raw_sheets in item_mapping.items():
        val_found = None
        for sname in raw_sheets:
          if sname in wb_f.sheetnames and sname in wb_d.sheetnames:
            sf, sd = wb_f[sname], wb_d[sname]
            coord = None
            for c in cols:
              for r in rows:
                cell_ref = f"{c}{r}"
                v = sf[cell_ref].value
                if v and isinstance(v, str) and pattern.match(v):
                  coord = cell_ref
                  break
              if coord:
                break

            if coord and sd[coord].value is not None:
              try:
                val_found = float(sd[coord].value)
                break
              except ValueError:
                pass
        matrix_data[date_key][item_name] = (
            val_found if val_found is not None else ""
        )

      wb_f.close()
      wb_d.close()
    except Exception:
      pass

  df = pd.DataFrame(matrix_data)
  final_order = list(item_mapping.keys())
  for s in final_order:
    if s not in df.index:
      df.loc[s] = ""
  df = df.reindex(index=final_order)
  df = df.reindex(sorted(df.columns), axis=1)
  return df


# ------------------- 路由 -------------------


@app.route("/")
def index():
  df = get_all_resistor_data()
  dates = list(df.columns) if not df.empty else []
  items = list(df.index) if not df.empty else []

  chart_datasets = []
  colors = [
      "#3b82f6",
      "#ef4444",
      "#10b981",
      "#f59e0b",
      "#8b5cf6",
      "#ec4899",
      "#06b6d4",
      "#84cc16",
  ]

  for idx, item in enumerate(items):
    row_vals = [
        df.loc[item, d] if pd.notna(df.loc[item, d]) and df.loc[item, d] != "" else None
        for d in dates
    ]
    chart_datasets.append({
        "label": item,
        "data": row_vals,
        "borderColor": colors[idx % len(colors)],
        "backgroundColor": colors[idx % len(colors)],
        "fill": False,
    })

  summary_chart = {"labels": dates, "datasets": chart_datasets}
  return render_template(
      "index.html", dates=dates, items=items, summary_chart=summary_chart
  )


@app.route("/chart")
def chart_page():
  df = get_all_resistor_data()
  dates = list(df.columns) if not df.empty else []
  items = list(df.index) if not df.empty else []

  selected_query = request.args.get("items", "")
  default_selected = (
      [i.strip() for i in selected_query.split(",") if i.strip()]
      if selected_query
      else []
  )

  item_data_map = {}
  for item in items:
    points = []
    for d in dates:
      v = df.loc[item, d]
      if pd.notna(v) and v != "":
        points.append({"date": d, "value": float(v)})
    item_data_map[item] = points

  return render_template(
      "chart.html",
      items=items,
      dates=dates,
      default_selected=json.dumps(default_selected, ensure_ascii=False),
      item_data_map=json.dumps(item_data_map, ensure_ascii=False),
  )


@app.route("/files")
def files_page():
  rel_path = request.args.get("path", "")
  target_dir = os.path.normpath(os.path.join(ROOT_DIR, rel_path))
  if not target_dir.startswith(ROOT_DIR):
    target_dir = ROOT_DIR
    rel_path = ""

  if os.path.isfile(target_dir):
    filename = os.path.basename(target_dir)
    ext = os.path.splitext(filename)[1].lower()
    content_type = "unknown"
    preview_data = ""

    try:
      if ext in [".txt", ".bat"]:
        content_type = "text"
        with open(target_dir, "r", encoding="utf-8", errors="ignore") as f:
          preview_data = f.read()

      elif ext == ".py":
        content_type = "code"
        with open(target_dir, "r", encoding="utf-8", errors="ignore") as f:
          preview_data = f.read()

      elif ext == ".json":
        content_type = "json"
        with open(target_dir, "r", encoding="utf-8", errors="ignore") as f:
          preview_data = f.read()

      elif ext == ".md":
        content_type = "markdown"
        with open(target_dir, "r", encoding="utf-8", errors="ignore") as f:
          raw_md = f.read()
          preview_data = {
              "raw": raw_md,
              "html": markdown.markdown(
                  raw_md, extensions=["tables", "fenced_code"]
              ),
          }

      elif ext in [".jpg", ".png", ".jpeg"]:
        content_type = "image"
        preview_data = f"/raw_file?path={rel_path}"

      elif ext in [".xlsx", ".xlsm"]:
        content_type = "excel"
        xls = pd.ExcelFile(target_dir)
        sheets_html = {}
        for sname in xls.sheet_names[:5]:
          df_s = pd.read_excel(xls, sheet_name=sname).head(50)
          sheets_html[sname] = df_s.to_html(
              classes=(
                  "table-auto w-full border-collapse border border-gray-300"
                  " text-xs text-center"
              ),
              index=False,
          )
        preview_data = sheets_html

      elif ext in [".doc", ".docx"]:
        content_type = "docx_preview"
        preview_data = f"/raw_file?path={rel_path}"

      elif ext == ".html":
        content_type = "html"
        preview_data = f"/raw_file?path={rel_path}"

      else:
        content_type = "unsupported"

    except Exception as e:
      content_type = "error"
      preview_data = str(e)

    parent_rel = os.path.dirname(rel_path)
    return render_template(
        "files.html",
        mode="file",
        rel_path=rel_path,
        parent_rel=parent_rel,
        filename=filename,
        ext=ext,
        content_type=content_type,
        preview_data=preview_data,
    )

  else:
    items = []
    try:
      for name in os.listdir(target_dir):
        full_p = os.path.join(target_dir, name)
        is_dir = os.path.isdir(full_p)
        child_rel = os.path.join(rel_path, name).replace("\\", "/")
        items.append({"name": name, "is_dir": is_dir, "rel_path": child_rel})
    except Exception:
      pass

    items.sort(key=lambda x: (not x["is_dir"], x["name"].lower()))
    parent_rel = (
        os.path.dirname(rel_path).replace("\\", "/") if rel_path else None
    )

    return render_template(
        "files.html",
        mode="dir",
        rel_path=rel_path,
        parent_rel=parent_rel,
        items=items,
    )


@app.route("/raw_file")
def raw_file():
  rel_path = request.args.get("path", "")
  target_file = os.path.normpath(os.path.join(ROOT_DIR, rel_path))
  if not target_file.startswith(ROOT_DIR) or not os.path.isfile(target_file):
    return abort(404)
  return send_file(target_file)


@app.route("/download")
def download_page():
  resources = [
      {
          "name": "算電阻.py",
          "path": "算電阻.py",
          "desc": "本機自動批次為 Excel 工作表插入 I2/I3 公式之 Python 腳本。",
      },
      {
          "name": "整理.py",
          "path": "整理.py",
          "desc": "包含 Tkinter GUI 之完整 Excel 數據整理與矩陣匯出工具。",
      },
  ]
  return render_template("download.html", resources=resources)


if __name__ == "__main__":
  app.run(debug=True, host="0.0.0.0", port=5000)