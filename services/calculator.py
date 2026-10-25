"""
Core pricing calculation engine.

All monetary math is done in the caller's input currency (rial or toman).
Shipping floor/ceiling are stored in toman in platform settings and converted
when the input currency is rial (1 toman = 10 rials).
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Literal, Mapping

Currency = Literal["rial", "toman"]

RIAL_PER_TOMAN = 10


@dataclass(frozen=True)
class PriceBreakdown:
    """Full breakdown of a single product price calculation."""

    currency: Currency
    base_price: float
    partner_discount_pct: float
    partner_discount_amount: float
    effective_base: float
    platform_fee_pct: float
    platform_fee: float
    profit_pct: float
    profit: float
    shipping_pct: float
    shipping: float
    shipping_badge: str  # "حداقل" | "حداکثر" | e.g. "۷٪"
    total_cost: float
    site_markup_pct: float
    site_price: float

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def to_toman(amount: float, currency: Currency) -> float:
    """Convert an amount to toman."""
    if currency == "toman":
        return amount
    return amount / RIAL_PER_TOMAN


def to_rial(amount: float, currency: Currency) -> float:
    """Convert an amount to rial."""
    if currency == "rial":
        return amount
    return amount * RIAL_PER_TOMAN


def convert_amount(amount: float, from_currency: Currency, to_currency: Currency) -> float:
    """Convert between rial and toman."""
    if from_currency == to_currency:
        return amount
    if from_currency == "toman" and to_currency == "rial":
        return amount * RIAL_PER_TOMAN
    return amount / RIAL_PER_TOMAN


def format_amount(amount: float, currency: Currency, *, decimals: int = 0) -> str:
    """
    Format an amount with thousand separators and Persian currency label.

    Example: ``5,829,300 ریال`` / ``582,930 تومان``
    """
    rounded = round(amount, decimals)
    if decimals == 0:
        formatted = f"{int(rounded):,}"
    else:
        formatted = f"{rounded:,.{decimals}f}"
    label = "ریال" if currency == "rial" else "تومان"
    return f"{formatted} {label}"


def format_dual(amount: float, currency: Currency, *, decimals: int = 0) -> str:
    """Show the same amount in both rials and tomans."""
    rial_value = to_rial(amount, currency)
    toman_value = to_toman(amount, currency)
    return (
        f"{format_amount(rial_value, 'rial', decimals=decimals)} / "
        f"{format_amount(toman_value, 'toman', decimals=decimals)}"
    )


def _shipping_bounds(settings: Mapping[str, Any], currency: Currency) -> tuple[float, float]:
    """Return (floor, ceiling) in the calculation currency."""
    floor_toman = float(settings["shipping_floor_toman"])
    ceiling_toman = float(settings["shipping_ceiling_toman"])
    if currency == "toman":
        return floor_toman, ceiling_toman
    return floor_toman * RIAL_PER_TOMAN, ceiling_toman * RIAL_PER_TOMAN


def _pct_label(pct: float) -> str:
    """Human-readable percent label, preferring Persian digits for whole numbers."""
    if abs(pct - round(pct)) < 1e-9:
        western = str(int(round(pct)))
    else:
        western = f"{pct:g}"
    persian_digits = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
    return western.translate(persian_digits) + "٪"


def calculate_price(
    base_price: float,
    settings: Mapping[str, Any],
    currency: Currency | None = None,
) -> PriceBreakdown:
    """
    Calculate the full price breakdown for a product.

    Parameters
    ----------
    base_price:
        Base list price in ``currency`` units.
    settings:
        Platform profile dict (percentages + shipping floor/ceiling in toman).
    currency:
        Input/output currency. Defaults to ``settings['default_currency']``.
    """
    if base_price < 0:
        raise ValueError("قیمت پایه نمی‌تواند منفی باشد.")

    cur: Currency = currency or settings.get("default_currency", "toman")  # type: ignore[assignment]
    if cur not in ("rial", "toman"):
        raise ValueError(f"واحد پول نامعتبر: {cur}")

    partner_discount_pct = float(settings.get("partner_discount_pct", 0))
    platform_fee_pct = float(settings.get("platform_fee_pct", 10))
    profit_pct = float(settings.get("profit_pct", 10))
    shipping_pct = float(settings.get("shipping_pct", 7))
    site_markup_pct = float(settings.get("site_markup_pct", 10))

    partner_discount_amount = base_price * (partner_discount_pct / 100.0)
    effective_base = base_price - partner_discount_amount

    platform_fee = effective_base * (platform_fee_pct / 100.0)
    profit = effective_base * (profit_pct / 100.0)

    floor, ceiling = _shipping_bounds(settings, cur)
    raw_shipping = effective_base * (shipping_pct / 100.0)
    shipping = min(max(raw_shipping, floor), ceiling)

    if shipping <= floor + 1e-9 and raw_shipping < floor:
        shipping_badge = "حداقل"
    elif shipping >= ceiling - 1e-9 and raw_shipping > ceiling:
        shipping_badge = "حداکثر"
    else:
        shipping_badge = _pct_label(shipping_pct)

    total_cost = effective_base + platform_fee + profit + shipping
    site_price = effective_base * (1.0 + site_markup_pct / 100.0)

    return PriceBreakdown(
        currency=cur,
        base_price=base_price,
        partner_discount_pct=partner_discount_pct,
        partner_discount_amount=partner_discount_amount,
        effective_base=effective_base,
        platform_fee_pct=platform_fee_pct,
        platform_fee=platform_fee,
        profit_pct=profit_pct,
        profit=profit,
        shipping_pct=shipping_pct,
        shipping=shipping,
        shipping_badge=shipping_badge,
        total_cost=total_cost,
        site_markup_pct=site_markup_pct,
        site_price=site_price,
    )
