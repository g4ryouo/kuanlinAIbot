"""寬林薪資計算網頁：上傳 Google Sheet 匯出檔 → 自動帶入天數/加班/請假 → 可手動修正 → 下載薪資條 xlsx。
啟動：python app.py  然後開 http://localhost:5000  （手機同網段也可用）"""
import io, os, tempfile, uuid
from flask import Flask, jsonify, request, send_file, send_from_directory
import engine

app = Flask(__name__, static_folder="static")
CACHE = {}  # token -> {"path", "ym", "derived", "leave"}
FIELDS = [("work_days", "工作天數"), ("lunch_days", "午餐天數"), ("ot_hours", "平日加班(時)"),
          ("holiday_ot_hours", "假日加班(時)"), ("sick_days", "病假(天)"), ("personal_days", "事假(天)"),
          ("dian_gong_days", "點工(天)"), ("absent_days", "缺勤扣薪天數")]


def _employee_view(name, emp, d):
    inputs = {k: d.get(k) or 0 for k, _ in FIELDS}
    if emp["type"] != "daily":
        inputs["work_days"] = None                      # 月薪者不用
    if not emp.get("weekend_offsets_absence"):
        inputs["absent_days"] = None
    if emp.get("lunch_mode") != "per_day":
        inputs["lunch_days"] = None
    return {"name": name, "title": emp["title"], "type": emp["type"], "inputs": inputs,
            "workdays_in_month": d.get("workdays_in_month", 22),
            "extras": dict(emp.get("default_extras", {})), "warnings": d.get("warnings", []),
            "ot_detail": d.get("ot_detail", []), "ot_already_marked": d.get("ot_already_marked", 0),
            "night_dates": d.get("night_dates", []), "off_dates": d.get("off_dates", []),
            "present_days": d.get("present_days"), "night_shifts": d.get("night_shifts")}


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.post("/api/load")
def load():
    f, ym = request.files["file"], request.form["month"]
    path = os.path.join(tempfile.gettempdir(), f"kuanlin_{uuid.uuid4().hex}.xlsx")
    f.save(path)
    try:
        derived = engine.derive_all(path, ym)
        year, month = map(int, ym.split("-"))
        data = engine.load_workbook_data(path, year)
    except Exception as e:  # 檔案格式不對等
        return jsonify(error=f"讀取失敗：{e}"), 400
    n2u = {n: u for u, n in data["ids"].items()}
    leave = {n: [x for x in data["leave"].get(n2u.get(n), []) if x[2].month <= month] for n in derived}
    token = uuid.uuid4().hex
    CACHE[token] = {"path": path, "ym": ym, "derived": derived, "leave": leave}
    cfg = engine.load_cfg()
    emps = [_employee_view(n, cfg["employees"][n], derived.get(n, {})) for n in cfg["employees"]]
    return jsonify(token=token, month=ym, fields=FIELDS, employees=emps)


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _calc_all(ym, rows):
    out = {}
    for row in rows:
        inp = {k: _num(v) for k, v in row["inputs"].items()}
        inp["workdays_in_month"] = _num(row.get("workdays_in_month")) or 22
        extras = {k: _num(v) or 0 for k, v in row.get("extras", {}).items() if k.strip()}
        out[row["name"]] = {"res": engine.calc(row["name"], inp, ym, extras), "inputs": inp}
    return out


@app.post("/api/calc")
def calc():
    j = request.get_json()
    return jsonify({n: v["res"] for n, v in _calc_all(j["month"], j["employees"]).items()})


@app.post("/api/export")
def export():
    j = request.get_json()
    c = CACHE.get(j["token"])
    if not c:
        return jsonify(error="資料已過期，請重新上傳"), 400
    results = _calc_all(j["month"], j["employees"])
    buf = io.BytesIO()
    tmp = os.path.join(tempfile.gettempdir(), f"out_{uuid.uuid4().hex}.xlsx")
    engine.export_xlsx(j["month"], results, c["derived"], tmp, c["leave"])
    buf.write(open(tmp, "rb").read()); os.remove(tmp); buf.seek(0)
    y, m = j["month"].split("-")
    return send_file(buf, as_attachment=True, download_name=f"{y}年{int(m)}月寬林薪資條.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
