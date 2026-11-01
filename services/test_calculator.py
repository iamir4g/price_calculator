"""
Quick regression checks for the pricing engine.

Run from the project root (venv active):
    python tests/test_calculator.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import io  # noqa: E402

import openpyxl  # noqa: E402

from services.calculator import (  # noqa: E402
    calculate_bundle,
    calculate_competitive_price,
    calculate_price,
    list_price_for_discount,
)
from services.report import add_profit, compare_with_settings, parse_settlement_report  # noqa: E402
from services.batch import process_dataframe  # noqa: E402
import pandas as pd  # noqa: E402

SNAPP = {
    "platform_fee_pct": 9,
    "infra_fee_pct": 0.6,
    "vat_pct": 9,
    "vat_on_infra": True,
    "profit_pct": 10,
    "partner_discount_pct": 10,
    "shipping_pct": 7,
    "shipping_floor_toman": 30000,
    "shipping_ceiling_toman": 190000,
    "rounding_toman": 1000,
    "default_currency": "toman",
}


def test_lamp_20w_matches_excel():
    r = calculate_price(509_000, SNAPP, "toman")  # list price 5,090,000 rial
    assert r.effective_base == 458_100
    assert r.sell_price == 611_000
    assert round(r.shipping) == 42_770
    assert r.shipping_badge == "۷٪"
    assert 10.0 <= r.actual_profit_pct < 10.2


def test_rial_input_gives_same_price():
    r = calculate_price(5_090_000, SNAPP, "rial")
    assert r.sell_price == 6_110_000


def test_shipping_floor_and_ceiling_regimes():
    cheap = calculate_price(100_000, SNAPP, "toman")
    assert cheap.shipping == 30_000 and cheap.shipping_badge == "حداقل"
    pricey = calculate_price(5_000_000, SNAPP, "toman")
    assert pricey.shipping == 190_000 and pricey.shipping_badge == "حداکثر"


def test_price_is_minimal_and_meets_target():
    for base in (50_000, 144_000, 509_000, 1_234_567, 9_999_999):
        r = calculate_price(base, SNAPP, "toman")
        assert r.profit >= r.target_net - r.effective_base - 1e-6
        # one rounding step lower must miss the target
        lower = r.sell_price - 1000
        k = 1 - 0.09 * 1.09 - 0.006 * 1.09
        ship = min(max(lower * 0.07, 30_000), 190_000)
        assert lower * k - ship < r.target_net


def test_batch_with_row_overrides():
    df = pd.DataFrame(
        {"قیمت لیست": ["509,000", "۲۰۰۰۰۰"], "کمیسیون (%)": [9, "11%"], "هزینه زیرساخت (%)": [0.6, None]}
    )
    out = process_dataframe(df, SNAPP, base_price_column="قیمت لیست", currency="toman")
    assert out.loc[0, "قیمت فروش در پلتفرم"] == 611_000
    assert out.loc[1, "درصد کمیسیون"] == 11
    assert out.loc[1, "درصد زیرساخت"] == 0.6  # empty cell -> platform default


def test_bundle_matches_excel_row_20():
    b = calculate_bundle(509_000, SNAPP, 3, "toman")
    assert b.bundle_price == 1_832_000
    assert round(b.unit_price_in_bundle, 2) == 610_666.67
    assert b.min_qty_for_max_shipping == 5
    assert b.bundle.profit >= 3 * 458_100 * 0.10 - 1e-6


def test_bundle_of_one_equals_single():
    b = calculate_bundle(509_000, SNAPP, 1, "toman")
    assert b.bundle_price == b.single.sell_price == 611_000
    assert abs(b.saving_pct) < 1e-9


def test_bundle_batch_columns():
    df = pd.DataFrame({"قیمت لیست": [509000]})
    out = process_dataframe(df, SNAPP, base_price_column="قیمت لیست", currency="toman", bundle_quantity=3)
    assert out.loc[0, "قیمت کل باندل"] == 1_832_000
    assert out.loc[0, "حداقل تعداد برای سقف ارسال"] == 5


def test_competitive_follow_floor_and_none():
    cfg = {**SNAPP, "floor_profit_pct": 3, "competitor_step_toman": 1000}
    q = calculate_competitive_price(509_000, cfg, 580_000, currency="toman")
    assert q.status == "follow" and q.recommended_price == 579_000
    assert q.gap_to_competitor == -1_000
    assert q.recommended.actual_profit_pct >= 3.0
    q = calculate_competitive_price(509_000, cfg, 570_000, currency="toman")
    assert q.status == "floor" and q.recommended_price == q.floor.sell_price == 572_000
    q = calculate_competitive_price(509_000, cfg, None, currency="toman")
    assert q.status == "no_competitor" and q.recommended_price == 611_000


def test_competitive_rial_matches_toman():
    cfg = {**SNAPP, "floor_profit_pct": 3, "competitor_step_toman": 1000}
    t = calculate_competitive_price(509_000, cfg, 580_000, currency="toman")
    r = calculate_competitive_price(5_090_000, cfg, 5_800_000, currency="rial")
    assert r.recommended_price == t.recommended_price * 10


def test_batch_competitor_and_floor_columns():
    df = pd.DataFrame({"قیمت لیست": [509000, 509000], "قیمت رقیب": [580000, 570000], "کف حاشیه (%)": [3, 3]})
    out = process_dataframe(df, SNAPP, base_price_column="قیمت لیست", currency="toman")
    assert list(out["قیمت پیشنهادی"]) == [579_000, 572_000]
    assert out.loc[1, "وضعیت رقابت"].startswith("رقیب زیر کف")
    plain = process_dataframe(df[["قیمت لیست"]], SNAPP, base_price_column="قیمت لیست", currency="toman")
    assert "قیمت پیشنهادی" not in plain.columns and "قیمت کف" in plain.columns


def test_vat_10_is_the_default_and_matches_report_rates():
    cfg = {k: v for k, v in SNAPP.items() if k != "vat_pct"}  # default VAT
    r = calculate_price(509_000, cfg, "toman")
    assert r.vat_pct == 10
    assert r.sell_price == 612_000
    assert r.actual_profit_pct >= 10.0


def test_profit_at_fixed_price_depends_on_order_quantity():
    cfg = {**SNAPP, "vat_pct": 10}
    price = 510_000  # the real listing price of the 20W lamp
    one = calculate_price(458_100, cfg, "toman", fixed_sell_price=price)
    ten = calculate_price(458_100 * 10, cfg, "toman", fixed_sell_price=price * 10)
    assert one.shipping == 35_700  # 7%
    assert ten.shipping == 190_000  # ceiling: far cheaper per unit
    assert ten.profit / 10 > one.profit  # same price, better margin in bulk


def test_buy_price_column_overrides_list_price():
    cfg = {**SNAPP, "vat_pct": 10}
    df = pd.DataFrame({"قیمت لیست": [509_000, 509_000], "قیمت خرید": [None, 420_000]})
    out = process_dataframe(df, cfg, base_price_column="قیمت لیست", currency="toman")
    assert out.loc[0, "قیمت خرید مؤثر"] == 458_100  # list x (1 - 10%)
    assert out.loc[1, "قیمت خرید مؤثر"] == 420_000  # real buy price, no extra discount


def test_seller_discount_helper():
    cfg = {**SNAPP}
    list_p, final_p = list_price_for_discount(579_000, 5, cfg, "toman")
    assert list_p % 1000 == 0 and final_p >= 579_000
    assert list_p - 1000 < 579_000 / 0.95 <= list_p


def _fake_report() -> bytes:
    """A tiny workbook in the Snappshop settlement format (VAT 10%)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "عملکرد کلی"
    ws.append(["خلاصه عملکرد طی دوره مالی:", "از 1405/07/01 تا 1405/07/30"])
    ws.append(["مانده طی دوره", None, "-", 0])  # filled below
    sales_rows = [  # order, title, qty, unit price
        (111, "کالای الف", 10, 220_000),
        (222, "کالای الف", 1, 220_000),
        (222, "کالای ب", 2, 200_000),
    ]
    sh = {111: 154_000, 222: 43_400}
    s = wb.create_sheet("فروش")
    s.append(["عنوان کالا", "دسته بندی کالا", "تاریخ سفارش", "نوع سفارش", "تاریخ مبنای محاسبه عملکرد", "شماره سفارش",
              "قیمت واحد (تومان)", "تخفیف فروشنده (تومان)", "قیمت فروش با ارزش افزوده (تومان)",
              "قیمت فروش بدون ارزش افزوده (تومان)", "تعداد", "جمع فروش (تومان)"])
    c = wb.create_sheet("کمیسیون فروش")
    c.append(["عنوان کالا", "دسته بندی کالا", "تاریخ سفارش", "نوع سفارش", "تاریخ مبنای محاسبه عملکرد", "شماره سفارش",
              "جمع فروش (تومان)", "درصد کمیسیون", "کمیسیون(تومان)", "مالیات بر ارزش افزوده (تومان)", "جمع کمیسیون (تومان)"])
    i = wb.create_sheet("هزینه توسعه زیرساخت")
    i.append(["عنوان کالا", "دسته بندی کالا", "تاریخ سفارش", "نوع سفارش", "تاریخ مبنای محاسبه عملکرد", "شماره سفارش",
              "جمع فروش (تومان)", "درصد هزینه توسعه زیرساخت", "هزینه توسعه زیرساخت (تومان)",
              "مالیات بر ارزش افزوده (تومان)", "جمع هزینه توسعه زیرساخت (تومان)"])
    total_net = 0.0
    for order, title, qty, unit in sales_rows:
        tot = qty * unit
        s.append([title, "لامپ", "1405-07-20", "اکسپرس", "1405-07-21", order, unit, 0, unit, unit, qty, tot])
        cb, ib = tot * 0.09, tot * 0.006
        c.append([title, "لامپ", "1405-07-20", "اکسپرس", "1405-07-21", order, tot, 9, cb, cb * 0.1, cb * 1.1])
        i.append([title, "لامپ", "1405-07-20", "اکسپرس", "1405-07-21", order, tot, 0.6, ib, ib * 0.1, ib * 1.1])
        total_net += tot - cb * 1.1 - ib * 1.1
    total_net -= sum(sh.values())
    e = wb.create_sheet("هزینه ارسال اکسپرس")
    e.append(["تاریخ سفارش", "شماره سفارش", "بدهکار(تومان)", "بستانکار(تومان)", "تاریخ مبنای محاسبه"])
    for order, amount in sh.items():
        e.append(["1405-07-20", order, amount, 0, "1405-07-21"])
    ws["D2"] = round(total_net)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_settlement_report_analysis():
    rep = parse_settlement_report(_fake_report())
    assert rep.totals["orders"] == 2 and rep.totals["units"] == 13
    assert abs(rep.totals["net"] - rep.payout_reported) < 2  # reproduces the payout
    assert abs(rep.vat_commission_pct - 10) < 1e-6 and abs(rep.vat_infra_pct - 10) < 1e-6
    a = rep.products[rep.products["title"] == "کالای الف"].iloc[0]
    assert a["units"] == 11 and a["orders"] == 2
    # order 222's shipping (43,400) is split by line sales: 220k of 620k -> 15,400
    order_222 = rep.orders[rep.orders["order"] == "222"].iloc[0]
    assert round(order_222["shipping"]) == 43_400
    prof = add_profit(rep.products, {"کالای الف": 150_000})
    row = prof[prof["title"] == "کالای الف"].iloc[0]
    assert round(row["profit"]) == round(row["net"] - 150_000 * 11)
    assert prof[prof["title"] == "کالای ب"].iloc[0]["profit"] != prof[prof["title"] == "کالای ب"].iloc[0]["profit"]  # NaN
    warns = compare_with_settings(rep, {"vat_pct": 9, "platform_fee_pct": 9, "infra_fee_pct": 0.6})
    assert any("مالیات واقعی" in w for w in warns)
    assert not compare_with_settings(rep, {**SNAPP, "vat_pct": 10})


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
