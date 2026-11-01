"""
Analysis of the Snappshop settlement report (``reports-summary-...xlsx``).

The report has one sheet per cost type. This module joins the sales lines with
their commission / infra-fee lines (by order number + product + line total) and
spreads each order's express-shipping cost over its lines, so you get the real
net per product and per order. With your buy prices it also gives real profit.

All amounts in the report are in toman.
"""

from __future__ import annotations

import io
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_BUY_PRICES_PATH = PROJECT_ROOT / "data" / "buy_prices.json"

SHEET_SUMMARY = "عملکرد کلی"
SHEET_SALES = "فروش"
SHEET_COMMISSION = "کمیسیون فروش"
SHEET_INFRA = "هزینه توسعه زیرساخت"
SHEET_SHIPPING = "هزینه ارسال اکسپرس"

REQUIRED_SHEETS = (SHEET_SALES, SHEET_COMMISSION, SHEET_INFRA, SHEET_SHIPPING)

# Sheets that are empty for a seller who only uses express shipping and has no
# cancellations. If they contain rows we show them (they change your payout).
# name -> (kind, columns). kind "debit_credit" = sum(debit) - sum(credit).
EXTRA_SHEETS: dict[str, tuple[str, tuple[str, ...]]] = {
    "کنسلی طی دوره": ("amount", ("جمع برگشت از فروش (تومان)",)),
    "برگشت از فروش دوره قبلی": ("amount", ("جمع برگشت از فروش (تومان)",)),
    "کمیسیون برگشت از فروش دوره قبلی": ("amount", ("جمع کمیسیون (تومان)",)),
    "برگشت از هزینه توسعه زیرساخت": ("amount", ("جمع هزینه توسعه زیرساخت (تومان)",)),
    "سایر": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
    "برگشت از هزینه ارسال اکسپرس": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
    "هزینه ارسال به انبار و پردازش": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
    "هزینه ارسال به مشتری": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
    "برگشت از هزینه ارسال به مشتری": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
    "هزینه ارسال سریع به انبار": ("debit_credit", ("بدهکار(تومان)", "بستانکار(تومان)")),
}

PRODUCT_COLUMNS = {
    "title": "کالا",
    "category": "دسته",
    "orders": "تعداد سفارش",
    "units": "تعداد فروش",
    "sales": "جمع فروش",
    "seller_discount": "جمع تخفیف فروشنده",
    "commission_pct": "کمیسیون ٪",
    "infra_pct": "زیرساخت ٪",
    "commission": "کمیسیون با مالیات",
    "infra": "زیرساخت با مالیات",
    "shipping": "ارسال",
    "net": "خالص دریافتی",
    "deduction_pct": "کسورات از فروش ٪",
    "avg_price": "میانگین قیمت فروش هر عدد",
    "breakeven_buy": "حداکثر قیمت خرید برای سر‌به‌سر (هر عدد)",
    "buy_price": "قیمت خرید هر عدد",
    "cost": "هزینه خرید",
    "profit": "سود خالص",
    "profit_pct": "سود از هزینه خرید ٪",
}

ORDER_COLUMNS = {
    "order": "شماره سفارش",
    "order_date": "تاریخ سفارش",
    "basis_date": "تاریخ نقطه فروش",
    "units": "تعداد",
    "sales": "جمع فروش",
    "commission": "کمیسیون با مالیات",
    "infra": "زیرساخت با مالیات",
    "shipping": "ارسال",
    "shipping_pct": "ارسال ٪ از فروش",
    "net": "خالص",
}


@dataclass(frozen=True)
class SettlementReport:
    period: str
    lines: pd.DataFrame
    orders: pd.DataFrame
    products: pd.DataFrame
    totals: dict[str, float]
    payout_reported: float | None
    vat_commission_pct: float | None
    vat_infra_pct: float | None
    adjustments: list[tuple[str, int, float]]  # (sheet, rows, amount)


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _num(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    cleaned = cleaned.replace({"-": "0", "nan": "0", "None": "0", "": "0"})
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def _order_str(value: Any) -> str:
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value)


def _keyed(df: pd.DataFrame, total_col: str) -> pd.DataFrame:
    """Add a unique join key: order | title | line total | n-th duplicate."""
    d = df.copy()
    d["order"] = d["شماره سفارش"].map(_order_str)
    total = _num(d[total_col]).round().astype("int64").astype(str)
    base = d["order"] + "|" + d["عنوان کالا"].astype(str) + "|" + total
    d["_key"] = base + "|" + base.groupby(base).cumcount().astype(str)
    return d


def _require(sheets: Mapping[str, pd.DataFrame]) -> None:
    missing = [s for s in REQUIRED_SHEETS if s not in sheets]
    if missing:
        raise ValueError(
            "این فایل شبیه گزارش تسویه اسنپ‌شاپ نیست. شیت‌های لازم پیدا نشد: " + "، ".join(missing)
        )


def _read_summary(raw: pd.DataFrame) -> tuple[str, float | None]:
    period = ""
    try:
        period = str(raw.iloc[0, 1])
    except (IndexError, KeyError):
        pass
    payout = None
    for _, row in raw.iterrows():
        cells = [str(c) for c in row.tolist()]
        if any("مانده طی دوره" in c for c in cells):
            for c in reversed(row.tolist()):
                try:
                    payout = float(str(c).replace(",", ""))
                    break
                except ValueError:
                    continue
            break
    return period, payout


# --------------------------------------------------------------------------- #
# parsing
# --------------------------------------------------------------------------- #
def parse_settlement_report(file_bytes: bytes) -> SettlementReport:
    """Parse the settlement workbook into per-line, per-order and per-product tables."""
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=0)
    _require(sheets)

    raw_summary = pd.read_excel(io.BytesIO(file_bytes), sheet_name=SHEET_SUMMARY, header=None) if SHEET_SUMMARY in sheets else pd.DataFrame()
    period, payout = _read_summary(raw_summary) if not raw_summary.empty else ("", None)

    sales = _keyed(sheets[SHEET_SALES], "جمع فروش (تومان)")
    comm = _keyed(sheets[SHEET_COMMISSION], "جمع فروش (تومان)")
    infra = _keyed(sheets[SHEET_INFRA], "جمع فروش (تومان)")

    lines = pd.DataFrame(
        {
            "_key": sales["_key"],
            "order": sales["order"],
            "title": sales["عنوان کالا"],
            "category": sales["دسته بندی کالا"],
            "order_date": sales["تاریخ سفارش"].astype(str),
            "basis_date": sales["تاریخ مبنای محاسبه عملکرد"].astype(str),
            "unit_price": _num(sales["قیمت واحد (تومان)"]),
            "seller_discount_unit": _num(sales["تخفیف فروشنده (تومان)"]),
            "qty": _num(sales["تعداد"]),
            "sales": _num(sales["جمع فروش (تومان)"]),
        }
    )
    lines["seller_discount"] = lines["seller_discount_unit"] * lines["qty"]

    c = pd.DataFrame(
        {
            "_key": comm["_key"],
            "commission_pct": _num(comm["درصد کمیسیون"]),
            "commission_base": _num(comm["کمیسیون(تومان)"]),
            "commission_vat": _num(comm["مالیات بر ارزش افزوده (تومان)"]),
            "commission": _num(comm["جمع کمیسیون (تومان)"]),
        }
    )
    i = pd.DataFrame(
        {
            "_key": infra["_key"],
            "infra_pct": _num(infra["درصد هزینه توسعه زیرساخت"]),
            "infra_base": _num(infra["هزینه توسعه زیرساخت (تومان)"]),
            "infra_vat": _num(infra["مالیات بر ارزش افزوده (تومان)"]),
            "infra": _num(infra["جمع هزینه توسعه زیرساخت (تومان)"]),
        }
    )
    lines = lines.merge(c, on="_key", how="left").merge(i, on="_key", how="left").fillna(
        {"commission": 0.0, "infra": 0.0, "commission_pct": 0.0, "infra_pct": 0.0,
         "commission_base": 0.0, "commission_vat": 0.0, "infra_base": 0.0, "infra_vat": 0.0}
    )

    # express shipping: one row per order
    ship_sheet = sheets[SHEET_SHIPPING].copy()
    ship_sheet["order"] = ship_sheet["شماره سفارش"].map(_order_str)
    ship_sheet["amount"] = _num(ship_sheet["بدهکار(تومان)"]) - _num(ship_sheet["بستانکار(تومان)"])
    ship_by_order = ship_sheet.groupby("order")["amount"].sum()

    order_sales = lines.groupby("order")["sales"].transform("sum")
    order_ship = lines["order"].map(ship_by_order).fillna(0.0)
    lines["shipping"] = (order_ship * lines["sales"] / order_sales.where(order_sales > 0, 1.0)).where(order_sales > 0, 0.0)
    lines["net"] = lines["sales"] - lines["commission"] - lines["infra"] - lines["shipping"]

    # ---- orders -------------------------------------------------------------
    g = lines.groupby("order")
    orders = pd.DataFrame(
        {
            "order_date": g["order_date"].min(),
            "basis_date": g["basis_date"].max(),
            "units": g["qty"].sum(),
            "sales": g["sales"].sum(),
            "commission": g["commission"].sum(),
            "infra": g["infra"].sum(),
        }
    )
    orders = orders.join(ship_by_order.rename("shipping"), how="outer")
    orders[["units", "sales", "commission", "infra", "shipping"]] = orders[
        ["units", "sales", "commission", "infra", "shipping"]
    ].fillna(0.0)
    orders["shipping_pct"] = (orders["shipping"] / orders["sales"].where(orders["sales"] > 0)) * 100
    orders["net"] = orders["sales"] - orders["commission"] - orders["infra"] - orders["shipping"]
    orders = orders.reset_index().rename(columns={"index": "order"})
    orders = orders.sort_values(["order_date", "order"], na_position="last").reset_index(drop=True)

    # ---- products -----------------------------------------------------------
    rows = []
    for title, d in lines.groupby("title", sort=False):
        sales_sum = d["sales"].sum()
        units = d["qty"].sum()
        rows.append(
            {
                "title": title,
                "category": d["category"].iloc[0],
                "orders": d["order"].nunique(),
                "units": units,
                "sales": sales_sum,
                "seller_discount": d["seller_discount"].sum(),
                "commission_pct": (d["commission_pct"] * d["sales"]).sum() / sales_sum if sales_sum else 0.0,
                "infra_pct": (d["infra_pct"] * d["sales"]).sum() / sales_sum if sales_sum else 0.0,
                "commission": d["commission"].sum(),
                "infra": d["infra"].sum(),
                "shipping": d["shipping"].sum(),
                "net": d["net"].sum(),
                "deduction_pct": (1 - d["net"].sum() / sales_sum) * 100 if sales_sum else 0.0,
                "avg_price": sales_sum / units if units else 0.0,
                "breakeven_buy": d["net"].sum() / units if units else 0.0,
            }
        )
    products = pd.DataFrame(rows).sort_values("sales", ascending=False).reset_index(drop=True)

    # ---- totals / VAT actually charged ---------------------------------------
    cb, cv = lines["commission_base"].sum(), lines["commission_vat"].sum()
    ib, iv = lines["infra_base"].sum(), lines["infra_vat"].sum()
    totals = {
        "sales": float(lines["sales"].sum()),
        "commission": float(lines["commission"].sum()),
        "infra": float(lines["infra"].sum()),
        "shipping": float(ship_by_order.sum()),
        "units": float(lines["qty"].sum()),
        "orders": float(lines["order"].nunique()),
        "seller_discount": float(lines["seller_discount"].sum()),
    }
    totals["net"] = totals["sales"] - totals["commission"] - totals["infra"] - totals["shipping"]

    adjustments: list[tuple[str, int, float]] = []
    for name, (kind, cols) in EXTRA_SHEETS.items():
        df = sheets.get(name)
        if df is None or df.empty or any(col not in df.columns for col in cols):
            continue
        if kind == "amount":
            amount = float(_num(df[cols[0]]).sum())
        else:
            amount = float(_num(df[cols[0]]).sum() - _num(df[cols[1]]).sum())
        adjustments.append((name, int(len(df)), amount))

    return SettlementReport(
        period=period,
        lines=lines.drop(columns=["_key"]),
        orders=orders,
        products=products,
        totals=totals,
        payout_reported=payout,
        vat_commission_pct=(cv / cb * 100) if cb else None,
        vat_infra_pct=(iv / ib * 100) if ib else None,
        adjustments=adjustments,
    )


# --------------------------------------------------------------------------- #
# profit & checks
# --------------------------------------------------------------------------- #
def add_profit(products: pd.DataFrame, buy_prices: Mapping[str, float]) -> pd.DataFrame:
    """Add buy price / cost / profit columns. Missing buy prices stay empty (NaN)."""
    out = products.copy()
    out["buy_price"] = out["title"].map(lambda t: buy_prices.get(t)).astype("float64")
    out["cost"] = out["buy_price"] * out["units"]
    out["profit"] = out["net"] - out["cost"]
    out["profit_pct"] = out["profit"] / out["cost"] * 100
    return out


def _model_shipping(order_sales: float, pct: float, floor: float, ceiling: float) -> float:
    return min(max(order_sales * pct / 100.0, floor), ceiling)


def compare_with_settings(report: SettlementReport, settings: Mapping[str, Any]) -> list[str]:
    """Human-readable warnings where the report disagrees with the platform settings."""
    warnings: list[str] = []
    vat = float(settings.get("vat_pct", 10))
    if report.vat_commission_pct is not None and abs(report.vat_commission_pct - vat) > 0.05:
        warnings.append(
            f"مالیات واقعی روی کمیسیون در گزارش {report.vat_commission_pct:.1f}٪ است ولی در تنظیمات {vat:g}٪ گذاشته‌اید."
        )
    if report.vat_infra_pct is not None:
        if not settings.get("vat_on_infra", True):
            warnings.append("در گزارش روی هزینه زیرساخت هم مالیات گرفته شده، ولی در تنظیمات «مالیات روی زیرساخت» خاموش است.")
        elif abs(report.vat_infra_pct - vat) > 0.05:
            warnings.append(
                f"مالیات واقعی روی زیرساخت در گزارش {report.vat_infra_pct:.1f}٪ است ولی در تنظیمات {vat:g}٪ است."
            )

    fee = float(settings.get("platform_fee_pct", 9))
    infra = float(settings.get("infra_fee_pct", 0.6))
    for _, r in report.products.iterrows():
        if abs(r["commission_pct"] - fee) > 0.01:
            warnings.append(f"کمیسیون «{r['title']}» در گزارش {r['commission_pct']:.2f}٪ است (پیش‌فرض پلتفرم {fee:g}٪).")
        if abs(r["infra_pct"] - infra) > 0.01:
            warnings.append(f"هزینه زیرساخت «{r['title']}» در گزارش {r['infra_pct']:.2f}٪ است (پیش‌فرض پلتفرم {infra:g}٪).")

    pct = float(settings.get("shipping_pct", 7))
    floor = float(settings.get("shipping_floor_toman", 30000))
    ceiling = float(settings.get("shipping_ceiling_toman", 190000))
    bad = 0
    for _, o in report.orders.iterrows():
        if o["sales"] > 0 and abs(_model_shipping(o["sales"], pct, floor, ceiling) - o["shipping"]) > 1.5:
            bad += 1
    if bad:
        warnings.append(f"هزینه ارسال واقعی {bad} سفارش با فرمول تنظیمات ({pct:g}٪، کف و سقف) نمی‌خواند.")
    return warnings


def analysis_to_xlsx(products: pd.DataFrame, orders: pd.DataFrame) -> bytes:
    """Excel export of the product and order tables with Persian headers."""
    p = products[[c for c in PRODUCT_COLUMNS if c in products.columns]].rename(columns=PRODUCT_COLUMNS)
    o = orders[[c for c in ORDER_COLUMNS if c in orders.columns]].rename(columns=ORDER_COLUMNS)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        p.round(2).to_excel(writer, index=False, sheet_name="محصولات")
        o.round(2).to_excel(writer, index=False, sheet_name="سفارش‌ها")
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# buy price persistence
# --------------------------------------------------------------------------- #
class BuyPriceStore:
    """Remembers the buy price per product title between runs (data/buy_prices.json)."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = Path(path) if path else DEFAULT_BUY_PRICES_PATH
        self._data: dict[str, float] = {}
        if self.path.exists():
            try:
                self._data = {str(k): float(v) for k, v in json.loads(self.path.read_text(encoding="utf-8")).items()}
            except (ValueError, OSError):
                self._data = {}

    def all(self) -> dict[str, float]:
        return dict(self._data)

    def update(self, prices: Mapping[str, float | None]) -> None:
        for title, value in prices.items():
            if value is None or pd.isna(value) or float(value) <= 0:
                self._data.pop(title, None)
            else:
                self._data[title] = float(value)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self._data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
