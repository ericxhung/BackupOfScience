from datetime import datetime
import os
import re
import string
from flask import Flask, render_template, jsonify
import openpyxl
import pandas as pd

app = Flask(
    __name__,
    template_folder="../templates",
    static_folder="../static",
)


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


def load_experiment_data():
  # 尋找專案根目錄下的 電阻 資料夾
  base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
  folder_path = os.path.join(base_dir, "電阻")

  file_list = []
  if os.path.exists(folder_path):
    for root, dirs, files in os.walk(folder_path):
      for file in files:
        if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
          f_path = os.path.join(root, file)
          folder_name = os.path.basename(root)

          # 解析日期
          date_str = parse_date_from_string(file)
          if not date_str:
            date_str = parse_date_from_string(folder_name)
          if not date_str:
            date_str = "2026-01-01"  # 預設保底日期

          file_list.append({"path": f_path, "filename": file, "date": date_str})

  if not file_list:
    return pd.DataFrame(), {}

  matrix_data = {}
  cols = list(string.ascii_uppercase[:16])
  rows = range(1, 31)
  pattern = re.compile(r"^=\d+\*")

  all_items_set = set()

  # 先收集所有出現過的工作表名稱作為項目
  for f in file_list:
    try:
      wb = openpyxl.load_workbook(f["path"], read_only=True, data_only=True)
      for sname in wb.sheetnames:
        if sname not in ["氣溫", "溫度"] and "plxdaq" not in sname.lower():
          all_items_set.add(sname)
      wb.close()
    except Exception:
      pass

  sorted_items = sorted(list(all_items_set))

  for f in file_list:
    f_path = f["path"]
    dt_obj = datetime.strptime(f["date"], "%Y-%m-%d")
    date_col_name = dt_obj.strftime("%Y/%m/%d")

    if date_col_name not in matrix_data:
      matrix_data[date_col_name] = {item: "" for item in sorted_items}

    try:
      wb_formula = openpyxl.load_workbook(f_path, data_only=False)
      wb_data = openpyxl.load_workbook(f_path, data_only=True)

      for item_name in sorted_items:
        if item_name in wb_formula.sheetnames and item_name in wb_data.sheetnames:
          sheet_f = wb_formula[item_name]
          sheet_d = wb_data[item_name]

          found_coord = None
          for col in cols:
            for row in rows:
              cell_ref = f"{col}{row}"
              val = sheet_f[cell_ref].value
              if val is not None and isinstance(val, str):
                if pattern.match(val):
                  found_coord = cell_ref
                  break
            if found_coord:
              break

          if found_coord:
            cell_val = sheet_d[found_coord].value
            if cell_val is not None:
              try:
                matrix_data[date_col_name][item_name] = float(cell_val)
              except ValueError:
                pass

      wb_formula.close()
      wb_data.close()
    except Exception as e:
      print(f"讀取檔案錯誤: {e}")

  df = pd.DataFrame(matrix_data)
  if not df.empty:
    df = df.reindex(sorted(df.columns), axis=1)

  return df, matrix_data


@app.route("/")
def index():
  df, _ = load_experiment_data()
  if df.empty:
    return render_template(
        "index.html",
        tables="<p>目前「電阻」資料夾中沒有找到有效的 Excel 檔案或數據。</p>",
        dates=[],
        items=[],
        chart_data={},
    )

  # 準備給前端 Chart.js 用的圖表數據
  dates = list(df.columns)
  items = list(df.index)

  chart_datasets = []
  colors = [
      "#3b82f6",
      "#ef4444",
      "#10b981",
      "#f59e0b",
      "#8b5cf6",
      "#ec4899",
      "#06b6d4",
  ]
  for idx, item in enumerate(items):
    values = [
        df.loc[item, d] if pd.notna(df.loc[item, d]) and df.loc[item, d] != "" else None
        for d in dates
    ]
    chart_datasets.append({
        "label": item,
        "data": values,
        "borderColor": colors[idx % len(colors)],
        "backgroundColor": colors[idx % len(colors)],
        "fill": False,
        "tension": 0.1,
    })

  chart_data = {"labels": dates, "datasets": chart_datasets}

  # 轉成 HTML 表格
  html_table = df.to_html(
      classes="table-auto w-full border-collapse border border-gray-300 text-center text-sm",
      na_rep="-",
  )

  return render_template(
      "index.html",
      tables=html_table,
      dates=dates,
      items=items,
      chart_data=chart_data,
  )


if __name__ == "__main__":
  app.run(debug=True)