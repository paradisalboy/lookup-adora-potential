# -*- coding: utf-8 -*-
import os
import pandas as pd
from flask import Flask, render_template, request, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "data.xlsx")

SHEETS = {
    "تعدادی":      ["کد مشتری", "نام مشتری", "نام فروشنده", "روز ویزیت", "کد کالا", "نام کالا", "فرنچایز", "تامین کننده"],
    "تامین کننده": ["کد مشتری", "نام مشتری", "نام فروشنده", "روز ویزیت", "تامین کننده"],
    "فرنچایز":     ["کد مشتری", "نام مشتری", "نام فروشنده", "روز ویزیت", "فرنچایز"],
}

app = Flask(__name__)
dfs = {}


def load_all():
    global dfs
    dfs = {}
    for sheet in SHEETS:
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
            dfs[sheet] = df
        except Exception as e:
            print(f"[warn] شیت «{sheet}»: {e}")


def classify_columns(df, fixed_cols):
    grand = [c for c in df.columns if "grand" in c.lower() or "جمع" in c.lower()]
    monthly = [c for c in df.columns if c not in fixed_cols and c not in grand]
    return monthly, grand


load_all()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/reload", methods=["POST"])
def reload():
    try:
        load_all()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/sheets")
def api_sheets():
    return jsonify({"sheets": list(SHEETS.keys())})


@app.route("/api/suggest")
def api_suggest():
    """برای جستجوی زنده (autocomplete) روی یک ستون خاص."""
    sheet = request.args.get("sheet", "تعدادی").strip()
    col   = request.args.get("col", "").strip()
    q     = request.args.get("q", "").strip()

    df = dfs.get(sheet)
    if df is None:
        return jsonify({"items": []})
    if col not in df.columns:
        return jsonify({"items": []})

    series = df[col].dropna().astype(str).str.strip()
    if q:
        series = series[series.str.contains(q, case=False, na=False)]
    items = sorted(series.unique().tolist())
    items = [i for i in items if i and i.lower() != "nan"]
    return jsonify({"items": items[:80]})


@app.route("/api/data")
def api_data():
    sheet = request.args.get("sheet", "تعدادی").strip()
    df = dfs.get(sheet)
    if df is None:
        return jsonify({"error": f"شیت «{sheet}» بارگذاری نشده"}), 400

    fixed = SHEETS.get(sheet, [])
    sub = df.copy()

    # فیلترهای متنی (جستجوی جزئی، بدون حساسیت به حروف)
    filter_cols = {
        "customer":   "نام مشتری",
        "seller":     "نام فروشنده",
        "product":    "نام کالا",
        "product_code": "کد کالا",
        "supplier":   "تامین کننده",
        "franchise":  "فرنچایز",
    }
    for param, col in filter_cols.items():
        val = request.args.get(param, "").strip()
        if val and col in sub.columns:
            sub = sub[sub[col].astype(str).str.contains(val, case=False, na=False)]

    monthly, grand = classify_columns(sub, fixed)

    rows = []
    for _, r in sub.iterrows():
        row = {}
        for c in df.columns:
            v = r[c]
            if pd.isna(v):
                row[c] = ""
            elif isinstance(v, float) and v == int(v):
                row[c] = int(v)
            elif isinstance(v, (int, float)):
                row[c] = v
            else:
                row[c] = str(v).strip()
        rows.append(row)

    return jsonify({
        "columns": list(df.columns),
        "fixed_columns": fixed,
        "monthly_columns": monthly,
        "grand_total_columns": grand,
        "count": len(rows),
        "rows": rows,
    })


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
