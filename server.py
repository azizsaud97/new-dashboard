"""
Marafiq - Geographic Meter Management Dashboard
Flask backend serving both the HTML dashboard and the JSON/Excel APIs.

Run:  python server.py
Open: http://127.0.0.1:5000
"""

from flask import Flask, jsonify, send_file, request, abort
import pandas as pd
import numpy as np
import os
import io

# ════════════════════════════════════════════════════════════
EXCEL_FILE = "marafiq_real_data.xlsx"
HTML_FILE  = "dashboard.html"
# ════════════════════════════════════════════════════════════

app = Flask(__name__)

# ---------- helpers ----------
def _f(v):
    """Safe float — returns None for NaN / blanks / bad values."""
    try:
        if v is None:
            return None
        if isinstance(v, float) and np.isnan(v):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None

def _s(v):
    """Safe string — empty string for None/NaN, trimmed."""
    if v is None:
        return ""
    if isinstance(v, float) and np.isnan(v):
        return ""
    return str(v).strip()


# ---------- load + clean ----------
def load_data():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ '{EXCEL_FILE}' غير موجود في المجلد الحالي.")
        return pd.DataFrame()

    df = pd.read_excel(EXCEL_FILE)
    df = df.replace({np.nan: None})

    rows = []
    for _, row in df.iterrows():
        address = _s(row.get('Address'))
        # FIX: addresses look like 'A1-04-002-201' → split on '-' (not space)
        sector   = address.split('-')[0] if address else "غير محدد"
        district = sector[0] if sector and sector != "غير محدد" else "?"

        status = _s(row.get('Meter Status')) or "غير محدد"

        rows.append({
            "contract_account": row.get('Contract Account'),
            "name":             _s(row.get('Name')),
            "account_det":      _s(row.get('Account Determination')),
            "zone_type":        _s(row.get('Account Determination')),
            "authenticated":    _s(row.get('Meter Authenticated')),
            "phone":            _s(row.get('Tel No.')),
            "address":          address,
            "district":         district,
            "sector":           sector,
            "lat":              _f(row.get('Y')),   # Y → latitude
            "lng":              _f(row.get('X')),   # X → longitude
            "average_12m":      row.get('Average Last 12 Months'),
            "status":           status,
            "zero_consumption": _s(row.get('Zero Consumption')),
            "disc_consumption": _s(row.get('Disconnected and Consumption')),
        })

    return pd.DataFrame(rows)


global_df = load_data()
print(f"✅ تم تحميل {len(global_df)} مشترك من Excel")
if not global_df.empty:
    print(f"   ↳ المحلات: {sorted(global_df['sector'].unique().tolist())}")


# ---------- routes ----------
@app.route('/')
def home():
    if not os.path.exists(HTML_FILE):
        return f"❌ ملف {HTML_FILE} غير موجود!", 404
    return send_file(HTML_FILE)


@app.route('/api/health')
def health():
    return jsonify({
        "ok": True,
        "rows": int(len(global_df)),
        "sectors": int(global_df['sector'].nunique()) if not global_df.empty else 0,
    })


@app.route('/api/data')
def get_all():
    if global_df.empty:
        return jsonify([])
    valid = global_df.dropna(subset=['lat', 'lng'])
    return jsonify(valid.to_dict(orient='records'))


@app.route('/api/search/<query>')
def search(query):
    if global_df.empty:
        return jsonify({"error": "لا توجد بيانات"})
    q = str(query).strip()
    match = global_df[
        (global_df['contract_account'].astype(str) == q) |
        (global_df['phone'].astype(str) == q)
    ]
    if match.empty:
        return jsonify({"error": "لم يتم العثور على المشترك"})
    return jsonify(match.iloc[0].to_dict())


@app.route('/api/export')
def export():
    if global_df.empty:
        return "لا توجد بيانات", 404

    sector_filter = request.args.get('sector', 'all')
    zone_filter   = request.args.get('zone',   'all')

    out = global_df.copy()
    if sector_filter != 'all':
        out = out[out['sector'] == sector_filter]
    if zone_filter != 'all':
        out = out[out['zone_type'] == zone_filter]

    out = out.rename(columns={
        "contract_account": "Contract Account",
        "name":             "Name",
        "account_det":      "Account Determination",
        "authenticated":    "Meter Authenticated",
        "phone":            "Tel No.",
        "address":          "Address",
        "lng":              "X",
        "lat":              "Y",
        "average_12m":      "Average Last 12 Months",
        "status":           "Meter Status",
        "zero_consumption": "Zero Consumption",
        "disc_consumption": "Disconnected and Consumption",
    })
    cols = ["Contract Account", "Name", "Account Determination", "Meter Authenticated",
            "Tel No.", "Address", "X", "Y", "Average Last 12 Months",
            "Meter Status", "Zero Consumption", "Disconnected and Consumption"]
    out = out[[c for c in cols if c in out.columns]]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        out.to_excel(writer, index=False, sheet_name='Marafiq_Export')
    buf.seek(0)
    return send_file(
        buf,
        download_name="Marafiq_Export.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ════════════════════════════════════════════════════════════
if __name__ == '__main__':
    print("\n" + "═" * 50)
    print("🌐 السيرفر جاهز!")
    print("👉 افتح في المتصفح: http://127.0.0.1:5000")
    print("═" * 50 + "\n")
    app.run(debug=False, port=5000, host='127.0.0.1', use_reloader=False)
