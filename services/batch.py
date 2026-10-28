"""
Batch pricing: read CSV/Excel, append calculated columns, export XLSX.
"""

from __future__ import annotations

import io
from typing import Any, Mapping

import pandas as pd

from services.calculator import Currency, calculate_bundle, calculate_price

# Common column name candidates for base (manufacturer list) price
BASE_PRICE_CANDIDATES = (
    "قیمت پایه",
    "قیمت_پایه",
    "قیمتپایه",
    "قیمت لیست",
    "base_price",
    "baseprice",
    "list_price",
    "price",
    "قیمت",
    "مبلغ",
)

# Optional per-row overrides (commission / infra fee differ between categories).
# Values are percent numbers, e.g. 9 or 0.6 (a trailing % sign is accepted).
FEE_COLUMN_HINTS = ("کمیسیون", "commission")
INFRA_COLUMN_HINTS = ("زیرساخت", "infra")

OUTPUT_COLUMNS = {
    "partner_discount": "تخفیف خرید",
    "effective_base": "قیمت خرید",
    "fee_pct": "درصد کمیسیون",
    "infra_pct": "درصد زیرساخت",
    "platform_fee": "کمیسیون پلتفرم",
    "infra_fee": "هزینه زیرساخت",
    "vat": "مالیات",
    "shipping": "هزینه ارسال",
    "shipping_badge": "وضعیت ارسال",
    "sell_price": "قیمت فروش در پلتفرم",
    "net_received": "مبلغ خالص دریافتی",
    "profit": "سود خالص",
    "profit_pct": "درصد سود از قیمت خرید",
    "site_price": "قیمت فروش سایت",
    "bundle_qty": "تعداد در باندل",
    "bundle_price": "قیمت کل باندل",
    "bundle_unit": "قیمت هر عدد در باندل",
    "bundle_saving": "کاهش قیمت هر عدد نسبت به تکی (٪)",
    "bundle_shipping": "هزینه ارسال باندل",
    "bundle_profit": "سود خالص باندل",
    "min_qty_ship": "حداقل تعداد برای سقف ارسال",
}

_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def _parse_number(raw: Any) -> float:
    """Parse numbers with Persian/Arabic digits, thousand separators and % signs."""
    if raw is None or pd.isna(raw):
        raise ValueError("empty")
    text = str(raw).strip().translate(_DIGITS)
    for ch in (",", "٬", " ", "%", "٪"):
        text = text.replace(ch, "")
    text = text.replace("٫", ".")
    return float(text)


parse_number = _parse_number  # public alias (used by the UI)


def _find_column(columns: list[str], hints: tuple[str, ...]) -> str | None:
    for col in columns:
        low = str(col).strip().lower()
        if any(h in low for h in hints):
            return col
    return None


def detect_base_price_column(columns: list[str]) -> str | None:
    """Return the best matching base-price column name, or None."""
    normalized = {str(c).strip(): c for c in columns}
    lower_map = {str(c).strip().lower().replace(" ", "").replace("_", ""): c for c in columns}

    for candidate in BASE_PRICE_CANDIDATES:
        if candidate in normalized:
            return normalized[candidate]
        key = candidate.lower().replace(" ", "").replace("_", "")
        if key in lower_map:
            return lower_map[key]
    return None


def read_upload(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Load an uploaded CSV or Excel file into a DataFrame."""
    name = filename.lower()
    buffer = io.BytesIO(file_bytes)
    if name.endswith(".csv"):
        # Try utf-8-sig first (Excel on Windows often adds BOM)
        try:
            return pd.read_csv(buffer, encoding="utf-8-sig")
        except UnicodeDecodeError:
            buffer.seek(0)
            return pd.read_csv(buffer, encoding="cp1256")
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(buffer)
    raise ValueError("فرمت فایل پشتیبانی نمی‌شود. لطفاً CSV یا Excel آپلود کنید.")


def process_dataframe(
    df: pd.DataFrame,
    settings: Mapping[str, Any],
    *,
    base_price_column: str,
    currency: Currency,
    bundle_quantity: int | None = None,
) -> pd.DataFrame:
    """
    Append calculated pricing columns to a copy of ``df``.

    Non-numeric / empty base prices produce empty calculated cells.
    If the file has a commission and/or infra-fee column (names containing
    ``کمیسیون`` / ``زیرساخت``), those percentages override the platform
    defaults row by row. ``bundle_quantity`` defaults to the platform setting.
    """
    if base_price_column not in df.columns:
        raise ValueError(f"ستون قیمت پایه یافت نشد: {base_price_column}")

    result = df.copy()
    columns = [str(c) for c in result.columns]
    fee_col = _find_column([c for c in columns if c != base_price_column], FEE_COLUMN_HINTS)
    infra_col = _find_column([c for c in columns if c != base_price_column], INFRA_COLUMN_HINTS)

    keys = (
        "partner_discount",
        "effective_base",
        "fee_pct",
        "infra_pct",
        "platform_fee",
        "infra_fee",
        "vat",
        "shipping",
        "shipping_badge",
        "sell_price",
        "net_received",
        "profit",
        "profit_pct",
        "site_price",
        "bundle_qty",
        "bundle_price",
        "bundle_unit",
        "bundle_saving",
        "bundle_shipping",
        "bundle_profit",
        "min_qty_ship",
    )
    out: dict[str, list[Any]] = {k: [] for k in keys}

    for idx in range(len(result)):
        row_settings = dict(settings)
        try:
            if fee_col is not None:
                try:
                    row_settings["platform_fee_pct"] = _parse_number(result.iloc[idx][fee_col])
                except ValueError:
                    pass  # empty/invalid -> platform default
            if infra_col is not None:
                try:
                    row_settings["infra_fee_pct"] = _parse_number(result.iloc[idx][infra_col])
                except ValueError:
                    pass
            price = _parse_number(result.iloc[idx][base_price_column])
            bun = calculate_bundle(price, row_settings, bundle_quantity, currency=currency)
            b = bun.single
            values = {
                "partner_discount": round(b.partner_discount_amount),
                "effective_base": round(b.effective_base),
                "fee_pct": b.platform_fee_pct,
                "infra_pct": b.infra_fee_pct,
                "platform_fee": round(b.platform_fee),
                "infra_fee": round(b.infra_fee),
                "vat": round(b.vat_amount),
                "shipping": round(b.shipping),
                "shipping_badge": b.shipping_badge,
                "sell_price": round(b.sell_price),
                "net_received": round(b.net_received),
                "profit": round(b.profit),
                "profit_pct": round(b.actual_profit_pct, 2),
                "site_price": round(b.site_price),
                "bundle_qty": bun.quantity,
                "bundle_price": round(bun.bundle.sell_price),
                "bundle_unit": round(bun.unit_price_in_bundle),
                "bundle_saving": round(bun.saving_pct, 2),
                "bundle_shipping": round(bun.bundle.shipping),
                "bundle_profit": round(bun.bundle.profit),
                "min_qty_ship": bun.min_qty_for_max_shipping,
            }
        except (ValueError, TypeError):
            values = {k: None for k in keys}
        for k in keys:
            out[k].append(values[k])

    for k in keys:
        result[OUTPUT_COLUMNS[k]] = out[k]
    return result


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Serialize a DataFrame to an in-memory .xlsx file."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="قیمت‌گذاری")
    buffer.seek(0)
    return buffer.getvalue()
