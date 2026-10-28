"""
Quick regression checks for the pricing engine.

Run from the project root (venv active):
    python tests/test_calculator.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from services.calculator import calculate_bundle, calculate_price  # noqa: E402
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


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
