"""寬林薪資計算引擎：讀 Google Sheet 匯出檔(打卡/加班/請假) → 每人薪資 → 依範本輸出薪資條 xlsx。"""
import json, math, re, copy, datetime as dt
from pathlib import Path

HERE = Path(__file__).parent
CFG = HERE / "config.json"
LEAVE_PAID = {"特休"}  # 特休不扣薪


def load_cfg():
    return json.loads(CFG.read_text(encoding="utf-8"))


def r(x):  # 四捨五入
    return int(math.floor(x + 0.5))


# ---------- 純計算規則 ----------
def hourly_rate(emp):
    """時薪=本薪/30/8（日薪者 日薪/8），無條件進位（楊東翰 50000/240=208.3 → 209）。"""
    return math.ceil(emp["base"] / (8 if emp["type"] == "daily" else 240) - 1e-9)


def overtime_pay(emp, hours, holiday_hours=0.0):
    """平日前N小時x1.33、其後x1.66。假日加班：有設定 holiday_ot_mult 者 = ceil(點工日薪/8) x 倍率 x 時數
    （楊東翰1.33、王/莊1.0）；未設定者（張欽煌）視同平日。"""
    rate, tier = hourly_rate(emp), emp.get("ot_tier_hours", 2)
    if holiday_hours and "holiday_ot_mult" not in emp:
        hours, holiday_hours = hours + holiday_hours, 0
    pay = min(hours, tier) * rate * 1.33 + max(hours - tier, 0) * rate * 1.66
    if holiday_hours:
        hr = math.ceil(emp["holiday_day_rate"] / 8 - 1e-9)
        hr = r(hr * emp["holiday_ot_mult"]) if emp["holiday_ot_mult"] != 1 else hr
        return r(min(hours, tier) * rate * 1.33 + max(hours - tier, 0) * rate * 1.66) + round(hr * holiday_hours)
    return r(pay)


def leave_deduction(emp, sick_days=0.0, personal_days=0.0, workdays_in_month=22, absent_days=None):
    """(請假扣薪, 全勤獎金扣除)。病假半薪、事假全薪、全勤獎金依 獎金/30 按天扣；特休不扣。
    leave_rule: base_30=本薪/30；workdays=本薪/當月工作天數(用 absent_days，已扣掉假日補班)；daily=日薪制不扣。"""
    rule = emp.get("leave_rule", "base_30")
    if rule == "daily":
        return 0, 0
    if rule == "workdays":
        return r(emp["base"] / workdays_in_month * (absent_days if absent_days is not None
                 else personal_days + sick_days)), 0
    per = emp["base"] / 30
    bonus = emp.get("attendance_bonus", 0) / 30 * (personal_days + sick_days)
    return r(per * personal_days + per * 0.5 * sick_days), r(bonus)


def compute(emp, *, month=None, work_days=0, ot_hours=0.0, holiday_ot_hours=0.0, lunch_days=0,
            sick_days=0.0, personal_days=0.0, dian_gong_days=0.0, workdays_in_month=22,
            absent_days=None, lunch_per_day=150, lunch_cap=3000, lunch_total=None, ot_pay=None,
            extras=None):
    """extras: {'浮動工作獎金': 3000, ...} 每月變動項目，直接加進應發。"""
    ded, bonus_ded = leave_deduction(emp, sick_days, personal_days, workdays_in_month, absent_days)
    lines = {}
    lines["薪資(日薪x天數)" if emp["type"] == "daily" else "本薪"] = (
        r(emp["base"] * work_days) if emp["type"] == "daily" else emp["base"])
    lines.update(emp.get("allowances", {}))
    if emp.get("attendance_bonus"):
        lines["全勤獎金"] = emp["attendance_bonus"]
    if emp.get("lunch_mode") == "per_day":
        tot = r(lunch_days * lunch_per_day + emp.get("lunch_fixed", 0)) if lunch_total is None else lunch_total
        lines["免稅午餐津貼"] = min(tot, lunch_cap)
        if tot > lunch_cap:
            lines["應稅午餐津貼"] = tot - lunch_cap
    if dian_gong_days:
        lines["點工"] = r(dian_gong_days * emp.get("holiday_day_rate", 0))
    hy = emp.get("half_year_bonus")
    if hy and month in hy["months"]:
        lines["半年獎金"] = hy["amount"]
    for k, v in (extras or {}).items():
        if v:
            lines[k] = v
    lines["加班費"] = overtime_pay(emp, ot_hours, holiday_ot_hours) if ot_pay is None else ot_pay
    deducts = {"勞保費": emp["labor_ins"], "健保費": emp["health_ins"]}
    deducts.update(emp.get("other_deductions", {}))
    if ded:
        deducts["請假扣薪"] = ded
    if bonus_ded:
        deducts["全勤獎金扣除"] = bonus_ded
    if emp.get("tax"):
        deducts["扣繳稅額"] = emp["tax"]
    g, d = sum(lines.values()), sum(deducts.values())
    return {"應發": lines, "應扣": deducts, "應發合計": g, "應扣合計": d, "實領": g - d}


# ---------- 讀取匯出檔 ----------
def _md(s, year):
    m = re.match(r"\s*(\d{1,2})/(\d{1,2})", str(s or ""))
    try:
        return dt.date(year, int(m[1]), int(m[2])) if m else None
    except ValueError:
        return None


def _full(s):
    if isinstance(s, (dt.datetime, dt.date)):
        return s.date() if isinstance(s, dt.datetime) else s
    m = re.match(r"\s*(\d{4})/(\d{1,2})/(\d{1,2})", str(s or ""))
    return dt.date(int(m[1]), int(m[2]), int(m[3])) if m else None


def _hm(s):
    m = re.match(r"\s*(\d{1,2}):(\d{2})", str(s or ""))
    return int(m[1]) * 60 + int(m[2]) if m else None


def load_workbook_data(xlsx, year):
    """回傳 dict：ids(uid→姓名), punches[uid][date]={in,out}, ot[uid]=[(date,hrs,已打V)], leave[uid]=[(類別,天數,起,迄)]。
    一律用 LINE ID 對應員工（LINE 顯示名稱會改，例如「紅色」）。"""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ids = {row[0]: row[1] for row in wb["員工清單"].iter_rows(min_row=2, values_only=True) if row[0]}
    punches, ot, leave = {}, {}, {}
    for row in wb[str(year)].iter_rows(min_row=2, values_only=True):
        d = _md(row[2], year)
        if d and row[3] and row[0]:
            punches.setdefault(row[0], {}).setdefault(d, {"in": row[3], "out": row[5]})
    for row in wb["加班申請表"].iter_rows(values_only=True):  # 第1列常被資料蓋掉表頭，用 LINE ID 判斷
        if row[0] and str(row[0]).startswith("U") and row[6] == "審核通過" and _full(row[2]):
            done = str(row[9] if len(row) > 9 else "").strip().upper() in ("V", "TRUE", "✓")
            ot.setdefault(row[0], []).append((_full(row[2]), float(row[5] or 0), done, row[4]))
    for row in wb["請假申請表"].iter_rows(min_row=2, values_only=True):
        if row[0] and row[7] == "審核通過":
            st = _md(row[3], year)
            if st:
                leave.setdefault(row[0], []).append((row[2], float(row[5] or 0), st, _md(row[4], year) or st))
    return {"ids": ids, "punches": punches, "ot": ot, "leave": leave}


# ---------- 由打卡紀錄推導當月輸入 ----------
def derive(data, uid, emp, year, month, cfg):
    hol = {dt.date.fromisoformat(x) for x in cfg.get("holidays", [])}
    ty = {dt.date.fromisoformat(x) for x in cfg.get("typhoon_dates", [])}
    pun = {d: p for d, p in data["punches"].get(uid, {}).items() if d.year == year and d.month == month}
    warn, cont, night = [], set(), set()
    # 夜工：前一天有上班卡但沒下班卡(晚間上班)，隔天凌晨系統會當成新的一天「上班」→ 其實是前一天的下班卡
    for d in sorted(pun):
        prev = d - dt.timedelta(days=1)
        if prev in pun and not pun[prev]["out"] and (_hm(pun[prev]["in"]) or 0) >= 15 * 60 \
                and (_hm(pun[d]["in"]) or 99999) < 12 * 60:
            cont.add(d); night.add(prev)
    present = [d for d in pun if d not in cont]                       # 實際上班日（夜工只算1天）
    is_off = lambda d: d.weekday() >= 5 or d in hol
    lunch_days = 0
    for d in present:
        if d in ty:
            continue
        if pun[d]["out"] or d in night:
            lunch_days += 1
        else:
            warn.append(f"{d.month}/{d.day} 缺下班卡，不給午餐津貼")
    leaves = [x for x in data["leave"].get(uid, []) if x[2].year == year and x[2].month == month]
    lv = {}  # 日期→(類別, 天數權重)
    for typ, days, st, en in leaves:
        dd = [st + dt.timedelta(days=i) for i in range((en - st).days + 1)]
        dd = [x for x in dd if x.weekday() < 5 and x not in hol] or [st]
        for x in dd:
            lv[x] = (typ, days if len(dd) == 1 else 1.0)
    sick = sum(w for t, w in lv.values() if t == "病假")
    pers = sum(w for t, w in lv.values() if t == "事假")
    ot = [x for x in data["ot"].get(uid, []) if x[0].year == year and x[0].month == month]
    ot_wd = sum(h for d, h, *_ in ot if not is_off(d))
    ot_hd = sum(h for d, h, *_ in ot if is_off(d))
    off_days = [d for d in present if is_off(d)]                       # 假日/國定假日出勤
    wd_days = [d for d in (dt.date(year, month, i) for i in range(1, 32) if _ok(year, month, i))
               if d.weekday() < 5 and d not in hol]
    out = {"present_days": len(present), "night_shifts": len(night), "off_days": len(off_days),
           "lunch_days": lunch_days, "sick_days": sick, "personal_days": pers,
           "ot_hours": ot_wd, "holiday_ot_hours": ot_hd, "workdays_in_month": len(wd_days),
           "ot_detail": [(d.isoformat(), h, done, why) for d, h, done, why in ot],
           "ot_already_marked": sum(1 for x in ot if x[2]), "warnings": warn,
           "night_dates": sorted(d.isoformat() for d in night), "off_dates": sorted(d.isoformat() for d in off_days)}
    if emp["type"] == "daily":
        paid_leave = sum(w for t, w in lv.values() if t in LEAVE_PAID)
        out["work_days"] = len(present) + len(night) - sick - pers + 0 * paid_leave
        out["dian_gong_days"] = 0
    else:
        out["work_days"] = 0
        out["dian_gong_days"] = len(off_days)
    if emp.get("weekend_offsets_absence"):                              # 楊東翰：平日缺勤用假日出勤 1:1 補
        absent = {d: 1.0 for d in wd_days if d not in pun and d not in ty and d not in cont
                  and lv.get(d, ("", 0))[0] not in LEAVE_PAID}
        for d, (t, w) in lv.items():
            if t not in LEAVE_PAID:
                absent[d] = w
        a = sum(absent.values()); off = len(off_days)
        out["absent_days"] = max(a - off, 0)
        out["dian_gong_days"] = max(off - a, 0)
        out["absent_dates"] = sorted(d.isoformat() for d in absent)
    return out


def _ok(y, m, d):
    try:
        dt.date(y, m, d); return True
    except ValueError:
        return False


def derive_all(xlsx, ym):
    year, month = map(int, ym.split("-"))
    cfg = load_cfg(); data = load_workbook_data(xlsx, year)
    name2uid = {n: u for u, n in data["ids"].items()}
    res = {}
    for name, emp in cfg["employees"].items():
        uid = name2uid.get(name)
        if not uid:
            res[name] = {"warnings": [f"員工清單找不到「{name}」"], "inputs": {}}
            continue
        res[name] = derive(data, uid, emp, year, month, cfg)
    return res


INPUT_KEYS = ["work_days", "lunch_days", "ot_hours", "holiday_ot_hours", "sick_days", "personal_days",
              "dian_gong_days", "absent_days"]


def calc(name, inputs, ym, extras=None):
    cfg = load_cfg(); emp = cfg["employees"][name]; year, month = map(int, ym.split("-"))
    kw = {k: inputs[k] for k in INPUT_KEYS if inputs.get(k) not in (None, "")}
    return compute(emp, month=month, workdays_in_month=inputs.get("workdays_in_month", 22),
                   lunch_per_day=cfg["lunch_per_day"], lunch_cap=cfg.get("lunch_tax_free_cap", 3000),
                   extras=extras, **kw)


# ---------- 輸出：依「薪資條範本.xlsx」格式填入 ----------
_NORM = lambda s: re.sub(r"\s+", "", str(s or ""))
STATIC_NOTE = ("勞健保保費", "員工團體保險", "加班費計算", "到職", "排休", "薪資資訊屬個人隱私", "本項勞務費",
               "本月勞保", "假日加班以點工", "假日/專案加班")
SKIP_NOTE = ("已於", "目前尚未擁有")


def _label_values(emp, res, inp):
    A, D = res["應發"], res["應扣"]
    leave_amt = D.get("請假扣薪", 0)
    v = {"本薪": A.get("本薪", 0), "日薪": emp["base"], "工作天數": inp.get("work_days", 0),
         "免稅午餐津貼": A.get("免稅午餐津貼", 0) + sum(x for k, x in A.items() if k == "免稅午餐津貼" and False),
         "應稅午餐津貼": A.get("應稅午餐津貼", 0), "餐費補貼": A.get("免稅午餐津貼", 0),
         "全勤獎金": A.get("全勤獎金", 0), "加班費": A.get("加班費", 0),
         "勞保費": D.get("勞保費", 0), "健保費": D.get("健保費", 0), "薪資所得扣繳稅額": D.get("扣繳稅額", 0),
         "事假(按比例扣除)": leave_amt, "病假(按比例扣除)": leave_amt, "病假（按比例扣除)": leave_amt,
         "全勤獎金(按比例扣除)": D.get("全勤獎金扣除", 0), "行政作業分攤金": D.get("行政作業分攤金", 0),
         "點工": A.get("點工", 0), "點工加班": A.get("點工", 0), "半年獎金": A.get("半年獎金", 0)}
    for k, x in A.items():
        v.setdefault(k, x)
    if emp.get("lunch_mode") == "fixed":                                  # 駱心愉：固定午餐津貼
        v["免稅午餐津貼"] = A.get("免稅午餐津貼", 0)
    return {_NORM(k): x for k, x in v.items()}


def _notes(cfg, emp, name, res, inp, info, old_notes, year, month, all_leave):
    keep = [n for n in old_notes if any(k in n for k in STATIC_NOTE) and not any(k in n for k in SKIP_NOTE)]
    keep = [re.sub(r"^\d+\.\s*", "", n) for n in keep]
    dyn = []
    fmt = lambda iso: f"{int(iso[5:7])}/{int(iso[8:10])}"
    if info.get("ot_detail"):
        hrs = sum(h for _, h, *_ in info["ot_detail"])
        dyn.append("於 " + "、".join(fmt(d) for d, *_ in info["ot_detail"]) +
                   f" 申請加班共 {hrs:g} 小時，加班費為 ${res['應發'].get('加班費', 0):,} 元。")
    spec = []
    if info.get("off_dates"):
        spec.append("、".join(fmt(d) for d in info["off_dates"]) + " 假日出勤（點工）")
    if info.get("night_dates"):
        spec.append("、".join(fmt(d) for d in info["night_dates"]) + " 夜工計 2 天班")
    if spec:
        dyn.append("；".join(spec) + (f"，出勤天數為 {(inp.get('work_days') or 0):g} 天。" if emp["type"] == "daily" else "。"))
    if res["應扣"].get("請假扣薪"):
        dyn.append(f"本月請假扣除 ${res['應扣']['請假扣薪']:,} 元" +
                   (f"、全勤獎金 ${res['應扣']['全勤獎金扣除']:,} 元。" if res["應扣"].get("全勤獎金扣除") else "。"))
    if emp.get("lunch_mode") == "per_day":
        tot = res["應發"].get("免稅午餐津貼", 0) + res["應發"].get("應稅午餐津貼", 0)
        dyn.append(f"午餐津貼採純打卡實發制（跨日夜工計 1 次供餐、請假/缺卡/颱風假當日不發），"
                   f"合格供餐天數共 {(inp.get('lunch_days') or 0):g} 天，發放 ${tot:,} 元整。")
    sl = [x for x in all_leave if x[0] == "特休"]
    if sl:
        tot = next((int(m[1]) for n in old_notes for m in [re.search(r"特休共\s*(\d+)\s*天", n)] if m), None)
        used = sum(x[1] for x in sl)
        rng = "、".join(f"{x[2].month}/{x[2].day}" + (f" - {x[3].month}/{x[3].day}" if x[3] != x[2] else "") for x in sl)
        dyn.append(f"已於 {year}/{rng} 使用特休共 {used:g} 天" + (f"，剩餘特休共 {tot - used:g} 天。" if tot else "。"))
    # 靜態說明放前面，動態說明中段，保密條款放最後
    secret = [n for n in keep if "隱私" in n]
    lines = [n for n in keep if "隱私" not in n] + dyn + secret
    return [f"{i}. {t}" for i, t in enumerate(lines, 1)]


def export_xlsx(ym, results, derived, out_path, all_leave=None):
    """results[name] = {'res':compute結果, 'inputs':{...}}；以範本檔為底，保留所有格式/字型/版面。"""
    import openpyxl
    cfg = load_cfg(); year, month = map(int, ym.split("-"))
    wb = openpyxl.load_workbook(HERE / cfg["template"])
    pay = dt.datetime(year + month // 12, month % 12 + 1, cfg["pay_day"])
    last = (dt.date(year + month // 12, month % 12 + 1, 1) - dt.timedelta(days=1)).day
    for ws in list(wb.worksheets):
        name = ws.title
        if name not in results:
            wb.remove(ws); continue
    for name, ws in ((n, wb[n]) for n in [w.title for w in wb.worksheets]):
        emp = cfg["employees"][name]; res, inp = results[name]["res"], results[name]["inputs"]
        vals = _label_values(emp, res, inp)
        info = derived.get(name, {})
        heads = [row for row in range(1, ws.max_row + 1) if ws.cell(row, 2).value in ("本薪", "日薪")]
        blocks = []
        for h in heads:
            for c in range(2, ws.max_column + 1):
                lab = _NORM(ws.cell(h, c).value)
                if lab:
                    x = vals.get(lab, 0)
                    ws.cell(h + 2, c).value = x
        for row in range(1, ws.max_row + 1):
            b, a = ws.cell(row, 2).value, ws.cell(row, 1).value
            if isinstance(a, str) and "薪資條" in a:
                ws.cell(row, 1).value = f"{year} 年 {month} 月 薪資條"
            tgt = None
            for c in range(3, 8):
                if ws.cell(row, c).value not in (None, "：") :
                    tgt = ws.cell(row, c); break
            if b == "員工姓名" and tgt: tgt.value = name
            elif b == "員工職稱" and tgt: tgt.value = emp["title"]
            elif b == "發薪日期" and tgt: tgt.value = pay
            elif b == "計薪期間" and tgt: tgt.value = f"{year}/{month}/1 - {year}/{month}/{last}"
            elif b == "備註：": blocks.append(row)
        for i, start in enumerate(blocks):
            end = blocks[i + 1] - 1 if i + 1 < len(blocks) else ws.max_row
            nxt = [rw for rw in range(start + 1, end + 1) if isinstance(ws.cell(rw, 1).value, str)]
            end = (nxt[0] - 1) if nxt else end
            old = [ws.cell(rw, 2).value for rw in range(start + 1, end + 1) if isinstance(ws.cell(rw, 2).value, str)]
            old = [s for s in old if "www" not in s and s.strip()]
            lines = _notes(cfg, emp, name, res, inp, info, old, year, month, all_leave.get(name, []) if all_leave else [])
            for k, rw in enumerate(range(start + 1, end + 1)):
                cell = ws.cell(rw, 2)
                try:
                    cell.value = lines[k] if k < len(lines) else None
                except AttributeError:
                    pass
    wb.save(out_path)
    return out_path
