"""
Core pricing calculation engine.

The model mirrors the Snappshop Excel sheet (``snappshop_pricing.xlsx``):

* ``buy_price   = list_price * (1 - partner_discount%)``
* ``target_net  = buy_price * (1 + profit%)``           (profit is on the BUY price)
* The platform deducts, **from the selling price p**:
      commission%  * (1 + VAT)
    + infra_fee%   * (1 + VAT)   (VAT on infra is optional)
    + shipping     = clamp(shipping% * p, floor, ceiling)
* We solve for the smallest ``p`` such that ``p - deductions >= target_net``
  and round it UP to ``rounding_toman``.

Because shipping is a percentage of ``p`` (with a floor and a ceiling) the
equation is piece-wise linear, so it is solved exactly (not approximated):

    k = 1 - fee%*(1+VAT) - infra%*(1+VAT*flag)

    1. shipping = floor    ->  p = (target + floor) / k          if shipping% * p <= floor
    2. shipping = shipping%->  p = target / (k - shipping%)      if shipping% * p <= ceiling
    3. shipping = ceiling  ->  p = (target + ceiling) / k

All monetary math is done in the caller's input currency (rial or toman).
Shipping floor/ceiling and the rounding step are stored in toman in platform
settings and converted when the input currency is rial (1 toman = 10 rials).
"""

from __future__ import annotations

import math
from dataclasses import asdict, dataclass
from typing import Any, Literal, Mapping

Currency = Literal["rial", "toman"]

RIAL_PER_TOMAN = 10


@dataclass(frozen=True)
class PriceBreakdown:
    """Full breakdown of a single product price calculation."""

    currency: Currency
    # --- purchase side -----------------------------------------------------
    base_price: float  # manufacturer list price
    partner_discount_pct: float
    partner_discount_amount: float
    effective_base: float  # your buy price (list price after discount)
    profit_pct: float  # target profit, relative to buy price
    target_net: float  # buy price + target profit
    # --- platform side (all computed on the selling price) -----------------
    sell_price: float  # price to list on the platform
    platform_fee_pct: float
    platform_fee: float  # commission amount
    infra_fee_pct: float
    infra_fee: float
    vat_pct: float
    vat_amount: float  # VAT on commission (+ infra if enabled)
    shipping_pct: float
    shipping: float
    shipping_badge: str  # "حداقل" | "حداکثر" | e.g. "۷٪"
    # --- result ------------------------------------------------------------
    net_received: float  # what the platform pays out after all deductions
    profit: float  # actual profit = net_received - effective_base
    actual_profit_pct: float  # actual profit relative to buy price
    # --- own website -------------------------------------------------------
    site_markup_pct: float
    site_price: float

    @property
    def total_cost(self) -> float:
        """Backward-compatible alias: the price to list on the platform."""
        return self.sell_price

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


def _from_toman(amount_toman: float, currency: Currency) -> float:
    return amount_toman if currency == "toman" else amount_toman * RIAL_PER_TOMAN


def _shipping_bounds(settings: Mapping[str, Any], currency: Currency) -> tuple[float, float]:
    """Return (floor, ceiling) in the calculation currency."""
    floor_toman = float(settings.get("shipping_floor_toman", 30000))
    ceiling_toman = float(settings.get("shipping_ceiling_toman", 190000))
    return _from_toman(floor_toman, currency), _from_toman(ceiling_toman, currency)


def _pct_label(pct: float) -> str:
    """Human-readable percent label, preferring Persian digits for whole numbers."""
    if abs(pct - round(pct)) < 1e-9:
        western = str(int(round(pct)))
    else:
        western = f"{pct:g}"
    persian_digits = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
    return western.translate(persian_digits) + "٪"


def solve_sell_price(
    target_net: float,
    k: float,
    shipping_rate: float,
    floor: float,
    ceiling: float,
) -> float:
    """
    Smallest selling price whose net (after fees and shipping) equals ``target_net``.

    ``k`` is the net factor after commission, infra fee and VAT (see module doc).
    """
    if k <= 0:
        raise ValueError("مجموع کسورات پلتفرم نباید ۱۰۰٪ یا بیشتر باشد.")

    if shipping_rate <= 0:
        # Shipping never grows with price: it is always the floor.
        return (target_net + floor) / k

    # Case 1: shipping sits on the floor.
    p_floor = (target_net + floor) / k
    if p_floor <= floor / shipping_rate:
        return p_floor

    # Case 2: shipping is the plain percentage.
    if k - shipping_rate > 0:
        p_pct = target_net / (k - shipping_rate)
        if p_pct <= ceiling / shipping_rate:
            return p_pct

    # Case 3: shipping sits on the ceiling.
    return (target_net + ceiling) / k


def _round_up(value: float, step: float) -> float:
    if step <= 0:
        return value
    # round() guards against float noise such as 611000.0000000001
    return math.ceil(round(value / step, 9)) * step


def calculate_price(
    base_price: float,
    settings: Mapping[str, Any],
    currency: Currency | None = None,
) -> PriceBreakdown:
    """
    Calculate the selling price and full breakdown for a product.

    Parameters
    ----------
    base_price:
        Manufacturer list price in ``currency`` units.
    settings:
        Platform profile dict (percentages, shipping floor/ceiling and rounding
        step in toman).
    currency:
        Input/output currency. Defaults to ``settings['default_currency']``.
    """
    if base_price < 0:
        raise ValueError("قیمت پایه نمی‌تواند منفی باشد.")

    cur: Currency = currency or settings.get("default_currency", "toman")  # type: ignore[assignment]
    if cur not in ("rial", "toman"):
        raise ValueError(f"واحد پول نامعتبر: {cur}")

    partner_discount_pct = float(settings.get("partner_discount_pct", 0))
    platform_fee_pct = float(settings.get("platform_fee_pct", 9))
    infra_fee_pct = float(settings.get("infra_fee_pct", 0.6))
    vat_pct = float(settings.get("vat_pct", 9))
    vat_on_infra = bool(settings.get("vat_on_infra", True))
    profit_pct = float(settings.get("profit_pct", 10))
    shipping_pct = float(settings.get("shipping_pct", 7))
    site_markup_pct = float(settings.get("site_markup_pct", 10))
    rounding_toman = float(settings.get("rounding_toman", 1000))

    fee_rate = platform_fee_pct / 100.0
    infra_rate = infra_fee_pct / 100.0
    vat_rate = vat_pct / 100.0
    ship_rate = shipping_pct / 100.0

    partner_discount_amount = base_price * (partner_discount_pct / 100.0)
    effective_base = base_price - partner_discount_amount
    target_net = effective_base * (1.0 + profit_pct / 100.0)

    floor, ceiling = _shipping_bounds(settings, cur)
    k = 1.0 - fee_rate * (1.0 + vat_rate) - infra_rate * (1.0 + (vat_rate if vat_on_infra else 0.0))

    raw_price = solve_sell_price(target_net, k, ship_rate, floor, ceiling)
    sell_price = _round_up(raw_price, _from_toman(rounding_toman, cur))

    # Re-derive every deduction from the (rounded) selling price.
    platform_fee = sell_price * fee_rate
    infra_fee = sell_price * infra_rate
    vat_amount = vat_rate * (platform_fee + (infra_fee if vat_on_infra else 0.0))

    raw_shipping = sell_price * ship_rate
    shipping = min(max(raw_shipping, floor), ceiling)
    if raw_shipping < floor:
        shipping_badge = "حداقل"
    elif raw_shipping > ceiling:
        shipping_badge = "حداکثر"
    else:
        shipping_badge = _pct_label(shipping_pct)

    net_received = sell_price - platform_fee - infra_fee - vat_amount - shipping
    profit = net_received - effective_base
    actual_profit_pct = (profit / effective_base * 100.0) if effective_base > 0 else 0.0

    site_price = effective_base * (1.0 + site_markup_pct / 100.0)

    return PriceBreakdown(
        currency=cur,
        base_price=base_price,
        partner_discount_pct=partner_discount_pct,
        partner_discount_amount=partner_discount_amount,
        effective_base=effective_base,
        profit_pct=profit_pct,
        target_net=target_net,
        sell_price=sell_price,
        platform_fee_pct=platform_fee_pct,
        platform_fee=platform_fee,
        infra_fee_pct=infra_fee_pct,
        infra_fee=infra_fee,
        vat_pct=vat_pct,
        vat_amount=vat_amount,
        shipping_pct=shipping_pct,
        shipping=shipping,
        shipping_badge=shipping_badge,
        net_received=net_received,
        profit=profit,
        actual_profit_pct=actual_profit_pct,
        site_markup_pct=site_markup_pct,
        site_price=site_price,
    )


@dataclass(frozen=True)
class BundleBreakdown:
    """Selling ``quantity`` units together in ONE order (shipping is paid once)."""

    quantity: int
    single: PriceBreakdown  # one unit sold alone
    bundle: PriceBreakdown  # the whole bundle priced as a single order
    unit_price_in_bundle: float  # bundle price / quantity
    saving_pct: float  # price drop per unit vs selling alone (can be ~0 or negative)
    min_qty_for_max_shipping: int  # units needed in one order to hit the shipping ceiling

    @property
    def bundle_price(self) -> float:
        return self.bundle.sell_price

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def calculate_bundle(
    base_price: float,
    settings: Mapping[str, Any],
    quantity: int | None = None,
    currency: Currency | None = None,
) -> BundleBreakdown:
    """
    Price a bundle of ``quantity`` identical units sold as one order.

    Everything is linear in the list price, so the bundle is priced exactly like
    one product whose list price is ``base_price * quantity``: the target profit
    scales with quantity, but shipping (floor/percent/ceiling) applies once to
    the whole order. This mirrors columns M-Q of the Excel sheet.
    """
    qty = int(quantity if quantity is not None else settings.get("bundle_quantity", 3))
    if qty < 1:
        raise ValueError("تعداد در باندل باید حداقل ۱ باشد.")

    cur: Currency = currency or settings.get("default_currency", "toman")  # type: ignore[assignment]
    single = calculate_price(base_price, settings, currency=cur)
    bundle = calculate_price(base_price * qty, settings, currency=cur)

    unit_price = bundle.sell_price / qty
    saving_pct = (1.0 - unit_price / single.sell_price) * 100.0 if single.sell_price > 0 else 0.0

    ship_rate = float(settings.get("shipping_pct", 7)) / 100.0
    _, ceiling = _shipping_bounds(settings, cur)
    if ship_rate > 0 and single.sell_price > 0:
        min_qty = math.ceil(round(ceiling / ship_rate / single.sell_price, 9))
    else:
        min_qty = 0

    return BundleBreakdown(
        quantity=qty,
        single=single,
        bundle=bundle,
        unit_price_in_bundle=unit_price,
        saving_pct=saving_pct,
        min_qty_for_max_shipping=min_qty,
    )
