"""端到端：用 10 月張欽煌/楊東翰的真實打卡型態造匯出檔，驗證天數推導與輸出薪資條。"""
import io, openpyxl
import pytest
import app as A


def make_xlsx(path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "2026"
    ws.append(["id", "name", "date", "in", "loc", "out", "loc", "h", "note"])
    zhang = [("10/1", "08:25", "17:30"), ("10/2", "08:19", "17:30"), ("10/3 (Sa)", "08:12", "17:30"),
             ("10/5", "08:23", "17:30"), ("10/6", "08:18", "17:30"), ("10/7", "19:16", ""), ("10/8", "04:48", ""),
             ("10/9", "08:09", "17:30"), ("10/12", "21:08", ""), ("10/13", "03:52", ""), ("10/14", "08:25", "17:30"),
             ("10/15", "08:22", "13:05"), ("10/16", "08:25", "17:30"), ("10/19", "08:15", "17:30"),
             ("10/20", "08:28", "17:30"), ("10/21", "08:22", "17:30"), ("10/22", "08:22", "17:30"),
             ("10/27", "08:12", "17:30"), ("10/28", "08:22", "17:30"), ("10/29", "08:19", "17:30"),
             ("10/30", "08:22", "17:30")]
    for d, i, o in zhang:
        ws.append(["Uz", "張欽煌", d, i, "x", o, "", "", ""])
    yang = [("10/1", "08:27", "17:30"), ("10/2", "08:30", "17:30"), ("10/5", "08:22", "17:30"),
            ("10/6", "08:19", "17:30"), ("10/7", "19:47", ""), ("10/8", "04:48", ""), ("10/12", "21:20", ""),
            ("10/13", "03:30", ""), ("10/14", "08:29", "17:30"), ("10/15", "08:30", "17:30"),
            ("10/16", "08:26", "17:30"), ("10/19", "08:30", "17:30"), ("10/20", "08:29", "17:30"),
            ("10/21", "08:20", "17:30"), ("10/27", "08:27", "17:30"), ("10/28", "08:31", "17:30"),
            ("10/29", "08:23", "17:30"), ("10/30", "08:23", "17:30")]
    for d, i, o in yang:
        ws.append(["Uy", "楊東翰", d, i, "x", o, "", "", ""])
    ws.append(["Uk", "紅色", "10/5", "08:20", "x", "" , "", "", ""])      # 缺下班卡 + LINE 名稱不同
    e = wb.create_sheet("員工清單"); e.append(["id", "n"])
    for u, n in (("Uz", "張欽煌"), ("Uy", "楊東翰"), ("Uk", "李弘盛")):
        e.append([u, n])
    o = wb.create_sheet("加班申請表")
    o.append(["Uz", "張欽煌", "2026/10/9", "", "加班", "1", "審核通過", "姑姑", "", ""])
    l = wb.create_sheet("請假申請表")
    l.append(["id", "n", "類別", "s", "e", "d", "r", "審核通過"])
    l.append(["Uz", "張欽煌", "事假", "10/15", "10/15", "0.5", "x", "審核通過"])
    l.append(["Uy", "楊東翰", "病假", "10/22", "10/22", "1", "x", "審核通過"])
    wb.save(path)


@pytest.fixture
def client(tmp_path):
    p = tmp_path / "in.xlsx"; make_xlsx(p)
    c = A.app.test_client()
    with open(p, "rb") as f:
        j = c.post("/api/load", data={"file": (f, "in.xlsx"), "month": "2026-10"}).get_json()
    return c, j


def emp(j, name):
    return next(e for e in j["employees"] if e["name"] == name)


def test_張欽煌_天數(client):
    _, j = client
    e = emp(j, "張欽煌")["inputs"]
    assert e["work_days"] == 20.5 and e["lunch_days"] == 19
    assert e["ot_hours"] + e["holiday_ot_hours"] == 1          # 10/9 是補假日，算假日加班(張欽煌視同平日費率)


def test_楊東翰_缺勤與補班(client):
    _, j = client
    e = emp(j, "楊東翰")["inputs"]
    assert e["absent_days"] == 1 and e["dian_gong_days"] == 0      # 10/22 病假；10/23 颱風不算


def test_李弘盛_缺下班卡不給餐(client):
    _, j = client
    v = emp(j, "李弘盛")
    assert v["inputs"]["lunch_days"] == 0 and any("缺下班卡" in w for w in v["warnings"])


def test_匯出薪資條(client):
    c, j = client
    rows = [{"name": e["name"], "inputs": e["inputs"], "extras": e["extras"],
             "workdays_in_month": e["workdays_in_month"]} for e in j["employees"]]
    calc = c.post("/api/calc", json={"month": "2026-10", "employees": rows}).get_json()
    assert calc["張欽煌"]["實領"] == 57831                       # 與 10 月 v5 實發薪資條一致
    r = c.post("/api/export", json={"token": j["token"], "month": "2026-10", "employees": rows})
    wb = openpyxl.load_workbook(io.BytesIO(r.data))
    ws = wb["張欽煌"]
    assert ws["D16"].value == 20.5 and ws["E16"].value == 499 and ws["G16"].value == 2850
    assert ws["D8"].value == "張欽煌" and "2026 年 10 月" in ws["A6"].value
    assert ws["D44"].value == 20.5                                # 第二份(員工聯)同步填入
