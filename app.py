from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def clean_header(col):
    return col.replace("\n", " ").strip()


def safe_key(col):
    return clean_header(col).lower().replace(" ", "_")


def get_master_files():
    files = []
    for f in os.listdir(UPLOAD_FOLDER):
        if f.endswith("_master_inventory.xlsx"):
            path = os.path.join(UPLOAD_FOLDER, f)
            files.append({
                "name": f,
                "path": path,
                "time": os.path.getmtime(path)
            })
    return sorted(files, key=lambda x: x["time"], reverse=True)


@app.route("/", methods=["GET", "POST"])
def index():
    master_files = get_master_files()

    if request.method == "POST":
        file = request.files["file"]

        df = pd.read_excel(file).fillna("")
        df.columns = [clean_header(c) for c in df.columns]

        column_map = {c: safe_key(c) for c in df.columns}

        return render_template(
            "table.html",
            columns=df.columns.tolist(),
            column_map=column_map,
            data=df.to_dict(orient="records"),
            master_files=master_files
        )

    return render_template("index.html", master_files=master_files)


@app.route("/generate", methods=["POST"])
def generate():
    columns = request.form.getlist("columns")
    safe_columns = request.form.getlist("safe_columns")
    row_count = int(request.form["row_count"])
    action = request.form.get("action")

    kitchen_name = request.form.get("kitchen_name", "Kitchen").replace(" ", "_")
    today = datetime.now().strftime("%Y-%m-%d")

    rows = []
    for i in range(row_count):
        row = {}
        for col, safe in zip(columns, safe_columns):
            row[col] = request.form.get(f"{safe}_{i}", "").strip()
        rows.append(row)

    df = pd.DataFrame(rows)

    requested_col = next(c for c in df.columns if "requested" in c.lower())
    supplier_col = next(c for c in df.columns if "supplier" in c.lower())
    current_col = next(c for c in df.columns if "current" in c.lower())

    df[requested_col] = pd.to_numeric(df[requested_col], errors="coerce").fillna(0)

    # 🔥 SAVE MASTER INVENTORY (ALWAYS)
    master_path = os.path.join(
        UPLOAD_FOLDER,
        f"{kitchen_name}_{today}_master_inventory.xlsx"
    )
    df.to_excel(master_path, index=False)

    # Draft only
    if action == "draft":
        return "Draft saved successfully"

    # ORDER LIST
    order_df = df[df[requested_col] > 0].copy()

    if current_col in order_df.columns:
        order_df.drop(columns=[current_col], inplace=True)

    output_path = os.path.join(
        UPLOAD_FOLDER,
        f"{kitchen_name}_{today}_order_list.xlsx"
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for supplier, group in order_df.groupby(supplier_col):
            group.to_excel(writer, sheet_name=str(supplier)[:31], index=False)

    wb = load_workbook(output_path)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True)
    wb.save(output_path)

    return send_file(output_path, as_attachment=True)


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(UPLOAD_FOLDER, filename)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
