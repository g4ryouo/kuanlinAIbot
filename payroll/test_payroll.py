"""用 2026/10 月實際發出的薪資條(v5)驗算規則。 python -m pytest -q"""
import json
from pathlib import Path
from payroll import compute, overtime_pay, leave_deduction

E = json.loads((Path(__file__).with_name("config.json")).read_text(encoding="utf-8"))["employees"]


def test_駱心愉():
    assert compute(E["駱心愉"])["實領"] == 28651


def test_李弘盛():  # 午餐18天2700；病假1天 583、全勤扣33
    assert leave_deduction(E["李弘盛"], sick_days=1) == (583, 33)
    assert compute(E["李弘盛"], lunch_days=18, sick_days=1)["實領"] == 36618


def test_王永杰():  # 午餐 4700 = 免稅3000+應稅1700
    res = compute(E["王永杰"], lunch_days=18)  # 18天x150+固定2000
    assert res["應發"]["應稅午餐津貼"] == 1700 and res["實領"] == 57845


def test_莊岳倫():  # 點工1天2500、加班 313、午餐 5000
    res = compute(E["莊岳倫"], lunch_days=20, dian_gong_days=1, ot_pay=313)
    assert res["實領"] == 49266


def test_楊東翰():  # 當月工作天數20，請假1天扣 2500；午餐18天
    assert leave_deduction(E["楊東翰"], sick_days=0, personal_days=1, workdays_in_month=20) == (2500, 0)
    res = compute(E["楊東翰"], lunch_days=18, personal_days=1, workdays_in_month=20)
    assert res["實領"] == 50200


def test_張欽煌():  # 20.5天、午餐19天2850、加班1h 499
    assert overtime_pay(E["張欽煌"], 1) == 499
    assert compute(E["張欽煌"], work_days=20.5, ot_hours=1, lunch_days=19)["實領"] == 57831


def test_楊東翰_加班():  # 平日前3h $278；假日每小時 $466
    assert overtime_pay(E["楊東翰"], 1) == 278
    assert overtime_pay(E["楊東翰"], 0, holiday_hours=0.5) == 233
