"""寬林薪資計算：讀取打卡/加班/請假紀錄，依 config.json 規則算出薪資條。

用法:
  python payroll.py --xlsx 員工打卡紀錄.xlsx --month 2026-10 --out 2026_10_薪資.xlsx
(xlsx 為 Google Sheet「檔案 > 下載 > .xlsx」的匯出檔)
"""
import argparse, json, math, re, datetime as dt
from pathlib import Path

CFG = Path(__file__).with_name("config.json")


def r(x):  # 四捨五入到整數元
    return int(math.floor(x + 0.5))


# ---------- 計算規則 ----------
def hourly_rate(emp):
    """時薪 = 本薪/30/8（日薪者為 日薪/8），無條件進位到整數（對照楊東翰 $209）。"""
    base = emp["base"] / (8 if emp["type"] == "daily" else 240)
    return math.ceil(base - 1e-9)


def overtime_pay(emp, hours, holiday_hours=0.0):
    """平日前 N 小時 x1.33，之後 x1.66；假日加班時薪 = 點工日薪/8 x1.33。"""
    rate = hourly_rate(emp)
    tier = emp.get("ot_tier_hours", 2)
    first, rest = min(hours, tier), max(hours - tier, 0)
    pay = first * rate * 1.33 + rest * rate * 1.66
    if holiday_hours:
        pay += holiday_hours * emp.get("holiday_day_rate", 0) / 8 * 1.33
    return r(pay)


def leave_deduction(emp, sick_days=0.0, personal_days=0.0, workdays_in_month=22):
    """回傳 (請假扣薪, 全勤獎金扣除)。
    病假：半薪；事假：全薪；全勤獎金依 獎金/30 按天扣。特休不扣。
    leave_rule: base_30=本薪/30、workdays=本薪/當月工作天數、daily=不扣(日薪制按出勤天數計)。"""
    rule = emp.get("leave_rule", "base_30")
    if rule == "daily":
        return 0, 0
    per_day = emp["base"] / (workdays_in_month if rule == "workdays" else 30)
    ded = per_day * personal_days + per_day * 0.5 * sick_days
    bonus = emp.get("attendance_bonus", 0) / 30 * (personal_days + sick_days)
    return r(ded), r(bonus)


def compute(emp, *, work_days=None, ot_hours=0.0, holiday_ot_hours=0.0, lunch_days=0,
            sick_days=0.0, personal_days=0.0, dian_gong_days=0.0, workdays_in_month=22,
            lunch_per_day=150, lunch_total=None, ot_pay=None, lunch_cap=3000):
    ded, bonus_ded = leave_deduction(emp, sick_days, personal_days, workdays_in_month)
    lines = {}  # 應發
    if emp["type"] == "daily":
        lines["薪資(日薪x天數)"] = r(emp["base"] * (work_days or 0))
    else:
        lines["本薪"] = emp["base"]
    lines.update(emp.get("allowances", {}))
    if emp.get("attendance_bonus"):
        lines["全勤獎金"] = emp["attendance_bonus"]
    if emp.get("lunch_mode") == "per_day":
        total = (lunch_days * lunch_per_day + emp.get("lunch_fixed", 0)) if lunch_total is None else lunch_total
        lines["免稅午餐津貼"] = min(total, lunch_cap)      # 免稅上限 $3,000
        if total > lunch_cap:
            lines["應稅午餐津貼"] = total - lunch_cap
    if dian_gong_days:
        lines["點工"] = r(dian_gong_days * emp.get("holiday_day_rate", 0))
    lines["加班費"] = overtime_pay(emp, ot_hours, holiday_ot_hours) if ot_pay is None else ot_pay
    gross = sum(lines.values())
    deducts = {"勞保費": emp["labor_ins"], "健保費": emp["health_ins"]}
    deducts.update(emp.get("other_deductions", {}))
    if ded: deducts["請假扣薪"] = ded
    if bonus_ded: deducts["全勤獎金扣除"] = bonus_ded
    if emp.get("tax"): deducts["扣繳稅額"] = emp["tax"]
    total_ded = sum(deducts.values())
    return {"應發": lines, "應扣": deducts, "應發合計": gross,
            "應扣合計": total_ded, "實領": gross - total_ded}


# ---------- 讀取 Google Sheet 匯出檔 ----------
def _md(s, year):
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})", str(s or ""))
    return dt.date(year, int(m[1]), int(m[2])) if m else None


def _full(s):
    m = re.match(r"\s*(\d{4})/(\d{1,2})/(\d{1,2})", str(s or ""))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def load_month(xlsx, year, month):
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ids = {row[0]: row[1] for row in wb["員工清單"].iter_rows(min_row=2, values_only=True) if row[0]}
    punches = {}
    for row in wb[str(year)].iter_rows(min_row=2, values_only=True):
        d = _md(row[2], year)
        if not d or d.month != month or not row[3]:
            continue
        name = ids.get(row[0]) or row[1]
        punches.setdefault(name, {})[d] = {"in": row[3], "out": row[5], "hours": row[7], "note": row[8]}
    ot, leave = {}, {}
    for row in wb["加班申請表"].iter_rows(values_only=True):   # 第1列資料被蓋掉表頭，所以用 LINE_ID 開頭判斷
        if not (row[0] and str(row[0]).startswith("U")) or row[6] != "審核通過":
            continue
        d = _full(row[2])
        if d and (d.year, d.month) == (year, month):
            ot.setdefault(row[1], []).append((d, float(row[5] or 0)))
    for row in wb["請假申請表"].iter_rows(min_row=2, values_only=True):
        if row[7] != "審核通過" or not row[1]:
            continue
        d = _md(row[3], year)
        if d and d.month == month:
            leave.setdefault(row[1], []).append((row[2], float(row[5] or 0), d, _md(row[4], year) or d))
    return punches, ot, leave


def run(xlsx, ym, out, overrides_path=None):
    import openpyxl
    from openpyxl.styles import Font, Alignment
    year, month = map(int, ym.split("-"))
    cfg = json.loads(CFG.read_text(encoding="utf-8"))
    ovr = json.loads(Path(overrides_path).read_text(encoding="utf-8")) if overrides_path else {}
    punches, ot, leave = load_month(xlsx, year, month)
    # 當月工作天數（排除週末；國定假日請用 overrides 的 workdays_in_month 調整）
    hol = set(cfg.get("holidays", []))
    wd = sum(1 for d in range(1, 32)
             if _valid(year, month, d) and dt.date(year, month, d).weekday() < 5
             and dt.date(year, month, d).isoformat() not in hol)
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    warns = []
    for name, emp in cfg["employees"].items():
        o = ovr.get(name, {})
        days = punches.get(name, {})
        sick = sum(x[1] for x in leave.get(name, []) if x[0] == "病假")
        pers = sum(x[1] for x in leave.get(name, []) if x[0] == "事假")
        ot_h = sum(h for d, h in ot.get(name, []) if d.weekday() < 5)
        hol_h = sum(h for d, h in ot.get(name, []) if d.weekday() >= 5)
        leave_days = set()
        for _t, _n, st, en in leave.get(name, []):
            x = st
            while x <= en:
                if x.weekday() < 5: leave_days.add(x)
                x += dt.timedelta(days=1)
        no_lunch = {dt.date.fromisoformat(x) for x in cfg.get("no_lunch_dates", [])}
        auto_lunch = sum(1 for d in days if d not in no_lunch and d not in leave_days)
        missing = [d.strftime("%m/%d") for d, p in days.items() if not p["out"]]
        if missing:
            warns.append(f"{name} 缺下班卡: {', '.join(missing)}")
        res = compute(emp,
                      work_days=o.get("work_days", len(days) - sick - pers),
                      ot_hours=o.get("ot_hours", ot_h), holiday_ot_hours=o.get("holiday_ot_hours", hol_h),
                      lunch_days=o.get("lunch_days", auto_lunch), sick_days=sick, personal_days=pers,
                      dian_gong_days=o.get("dian_gong_days", 0),
                      workdays_in_month=o.get("workdays_in_month", wd), lunch_per_day=cfg["lunch_per_day"],
                      lunch_total=o.get("lunch_total"), ot_pay=o.get("ot_pay"),
                      lunch_cap=cfg.get("lunch_tax_free_cap", 3000))
        ws = wb.create_sheet(name)
        ws.sheet_properties.tabColor = None
        ws["A1"] = cfg["company"]; ws["A1"].font = Font(bold=True, size=14)
        ws["A2"] = f"{year} 年 {month} 月 薪資條"
        ws["A3"] = f"員工姓名：{name}　職稱：{emp['title']}"
        ws["A4"] = f"發薪日期：{year + month // 12}/{month % 12 + 1}/{cfg['pay_day']}"
        row = 6
        for title, key in (("應發金額", "應發"), ("應扣金額", "應扣")):
            ws.cell(row, 1, title).font = Font(bold=True); row += 1
            for k, v in res[key].items():
                ws.cell(row, 1, k); ws.cell(row, 2, v); row += 1
            ws.cell(row, 1, "合計"); ws.cell(row, 2, res[key + "合計"]); row += 2
        ws.cell(row, 1, "實領金額").font = Font(bold=True)
        ws.cell(row, 2, res["實領"]).font = Font(bold=True)
        ws.column_dimensions["A"].width = 24
        print(f"{name}: 應發 {res['應發合計']:,} 應扣 {res['應扣合計']:,} 實領 {res['實領']:,}")
    for w_ in wb.worksheets:
        for rw in w_.iter_rows():
            for cell in rw:
                f = cell.font
                cell.font = Font(name="標楷體", bold=f.bold, size=f.size or 11)
    wb.save(out)
    for w in warns:
        print("⚠", w)


def _valid(y, m, d):
    try:
        dt.date(y, m, d); return True
    except ValueError:
        return False


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True); ap.add_argument("--month", required=True)
    ap.add_argument("--out", default="薪資.xlsx"); ap.add_argument("--overrides")
    a = ap.parse_args()
    run(a.xlsx, a.month, a.out, a.overrides)
