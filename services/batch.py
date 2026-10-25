"""
Batch pricing: read CSV/Excel, append calculated columns, export XLSX.
"""

from __future__ import annotations

import io
from typing import Any, Mapping

import pandas as pd

from services.calculator import Currency, calculate_price

# Common column name candidates for base price (Persian + English)
BASE_PRICE_CANDIDATES = (
    "قیمت پایه",
    "قیمت_پایه",
    "قیمتپایه",
    "base_price",
    "baseprice",
    "price",
    "قیمت",
    "مبلغ",
)

OUTPUT_COLUMNS = {
    "partner_discount": "تخفیف همکار",
    "effective_base": "قیمت پایه مؤثر",
    "platform_fee": "کارمزد پلتفرم",
    "profit": "سود",
    "shipping": "هزینه ارسال",
    "shipping_badge": "وضعیت ارسال",
    "total_cost": "هزینه تمام شده",
    "site_price": "قیمت فروش سایت",
}


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
) -> pd.DataFrame:
    """
    Append calculated pricing columns to a copy of ``df``.

    Non-numeric / empty base prices produce empty calculated cells.
    """
    if base_price_column not in df.columns:
        raise ValueError(f"ستون قیمت پایه یافت نشد: {base_price_column}")

    result = df.copy()
    fee_label = f"کارمزد پلتفرم {settings.get('platform_fee_pct', 10):g}٪"
    profit_label = f"سود {settings.get('profit_pct', 10):g}٪"

    partner_vals: list[float | None] = []
    effective_vals: list[float | None] = []
    fee_vals: list[float | None] = []
    profit_vals: list[float | None] = []
    shipping_vals: list[float | None] = []
    badge_vals: list[str | None] = []
    total_vals: list[float | None] = []
    site_vals: list[float | None] = []

    for raw in result[base_price_column]:
        try:
            if pd.isna(raw):
                raise ValueError("empty")
            # Handle Persian/Arabic digits and thousand separators
            text = str(raw).strip()
            text = text.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789"))
            text = text.replace(",", "").replace("٬", "").replace(" ", "")
            price = float(text)
            breakdown = calculate_price(price, settings, currency=currency)
            partner_vals.append(round(breakdown.partner_discount_amount))
            effective_vals.append(round(breakdown.effective_base))
            fee_vals.append(round(breakdown.platform_fee))
            profit_vals.append(round(breakdown.profit))
            shipping_vals.append(round(breakdown.shipping))
            badge_vals.append(breakdown.shipping_badge)
            total_vals.append(round(breakdown.total_cost))
            site_vals.append(round(breakdown.site_price))
        except (ValueError, TypeError):
            partner_vals.append(None)
            effective_vals.append(None)
            fee_vals.append(None)
            profit_vals.append(None)
            shipping_vals.append(None)
            badge_vals.append(None)
            total_vals.append(None)
            site_vals.append(None)

    result[OUTPUT_COLUMNS["partner_discount"]] = partner_vals
    result[OUTPUT_COLUMNS["effective_base"]] = effective_vals
    result[fee_label] = fee_vals
    result[profit_label] = profit_vals
    result[OUTPUT_COLUMNS["shipping"]] = shipping_vals
    result[OUTPUT_COLUMNS["shipping_badge"]] = badge_vals
    result[OUTPUT_COLUMNS["total_cost"]] = total_vals
    result[OUTPUT_COLUMNS["site_price"]] = site_vals
    return result


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    """Serialize a DataFrame to an in-memory .xlsx file."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="قیمت‌گذاری")
    buffer.seek(0)
    return buffer.getvalue()
