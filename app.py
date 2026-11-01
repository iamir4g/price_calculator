"""
Price Calculator — Streamlit UI (RTL / Persian).

Run locally:
    streamlit run app.py
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from services.batch import (
    parse_number,
    detect_base_price_column,
    process_dataframe,
    read_upload,
    to_xlsx_bytes,
)
from services.calculator import (
    Currency,
    calculate_bundle,
    calculate_competitive_price,
    calculate_price,
    convert_amount,
    format_amount,
    list_price_for_discount,
)
from services.report import (
    BuyPriceStore,
    PRODUCT_COLUMNS,
    ORDER_COLUMNS,
    add_profit,
    analysis_to_xlsx,
    compare_with_settings,
    parse_settlement_report,
)
from services.settings import SettingsStore

st.set_page_config(
    page_title="ماشین‌حساب قیمت محصول",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

RTL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700&display=swap');

html, body, [class*="css"], .stApp, .stMarkdown, .stText, .stButton,
.stSelectbox, .stNumberInput, .stTextInput, .stRadio, .stTabs, label,
div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {
    font-family: "Vazirmatn", Tahoma, sans-serif !important;
}

.stApp {
    direction: rtl;
    text-align: right;
}

[data-testid="stSidebar"] {
    direction: rtl;
    text-align: right;
}

h1, h2, h3, h4 {
    text-align: right !important;
}

div[data-testid="stMetric"] {
    background: linear-gradient(145deg, #f8fafc 0%, #eef2ff 100%);
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1rem 1.1rem;
}

.breakdown-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    padding: 1.1rem 1.25rem;
    margin-bottom: 0.75rem;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}

.breakdown-card .label {
    color: #64748b;
    font-size: 0.9rem;
    margin-bottom: 0.25rem;
}

.breakdown-card .value {
    color: #0f172a;
    font-size: 1.15rem;
    font-weight: 600;
}

.breakdown-card .value.money {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    justify-content: flex-start;
    gap: 0.15rem 1rem;
}

.money-main { unicode-bidi: isolate; }

.money-sub {
    unicode-bidi: isolate;
    font-size: 0.85rem;
    font-weight: 500;
    color: #64748b;
    padding-right: 1rem;
    border-right: 1px solid #cbd5e1;
}

.breakdown-card.highlight .money-sub {
    color: #99f6e4;
    border-right-color: rgba(255, 255, 255, 0.35);
}

.breakdown-card.highlight {
    background: linear-gradient(135deg, #0f766e 0%, #115e59 100%);
    border: none;
    color: #fff;
}

.breakdown-card.highlight .label { color: #ccfbf1; }
.breakdown-card.highlight .value { color: #ffffff; }

.badge {
    display: inline-block;
    background: #fef3c7;
    color: #92400e;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 0.15rem 0.55rem;
    border-radius: 999px;
    margin-inline-start: 0.4rem;
}

.badge.floor { background: #dbeafe; color: #1e40af; }
.badge.ceiling { background: #fee2e2; color: #991b1b; }
.badge.pct { background: #d1fae5; color: #065f46; }

.hero-sub {
    color: #64748b;
    margin-top: -0.5rem;
    margin-bottom: 1.5rem;
}
</style>
"""


def inject_rtl() -> None:
    st.markdown(RTL_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <script>
        const doc = window.parent.document;
        doc.documentElement.setAttribute('dir', 'rtl');
        doc.documentElement.setAttribute('lang', 'fa');
        doc.body.setAttribute('dir', 'rtl');
        </script>
        """,
        unsafe_allow_html=True,
    )


def get_store() -> SettingsStore:
    if "settings_store" not in st.session_state:
        st.session_state.settings_store = SettingsStore()
    return st.session_state.settings_store


def platform_options(store: SettingsStore) -> dict[str, str]:
    """Map display name -> id (unique names assumed; fall back to id in label)."""
    options: dict[str, str] = {}
    for p in store.list_platforms():
        label = p["name"]
        if label in options:
            label = f"{p['name']} ({p['id']})"
        options[label] = p["id"]
    return options


def _format_money_input(key: str) -> None:
    """on_change callback: rewrite the text field with thousand separators."""
    try:
        value = parse_number(st.session_state.get(key, ""))
    except ValueError:
        return  # leave invalid text as typed; the tab shows an error
    st.session_state[key] = f"{value:,.0f}" if float(value).is_integer() else f"{value:,}"


def render_amount_card(
    label: str,
    amount: float,
    currency: Currency,
    *,
    badge: str | None = None,
    highlight: bool = False,
) -> None:
    badge_html = ""
    if badge:
        css = "pct"
        if badge == "حداقل":
            css = "floor"
        elif badge == "حداکثر":
            css = "ceiling"
        badge_html = f'<span class="badge {css}">{badge}</span>'

    cls = "breakdown-card highlight" if highlight else "breakdown-card"
    other: Currency = "rial" if currency == "toman" else "toman"
    main_txt = format_amount(amount, currency)
    sub_txt = format_amount(convert_amount(amount, currency, other), other)
    st.markdown(
        f"""
        <div class="{cls}">
            <div class="label">{label}{badge_html}</div>
            <div class="value money">
                <span class="money-main" dir="rtl">{main_txt}</span>
                <span class="money-sub" dir="rtl">{sub_txt}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_text_card(label: str, text: str, *, highlight: bool = False) -> None:
    cls = "breakdown-card highlight" if highlight else "breakdown-card"
    st.markdown(
        f"""
        <div class="{cls}">
            <div class="label">{label}</div>
            <div class="value">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def tab_calculator(store: SettingsStore) -> None:
    platforms = platform_options(store)
    if not platforms:
        st.warning("هیچ پلتفرمی تعریف نشده است. از تب تنظیمات یک پلتفرم بسازید.")
        return

    active = store.get_active_platform()
    default_label = next(
        (name for name, pid in platforms.items() if active and pid == active["id"]),
        next(iter(platforms)),
    )

    col_a, col_b = st.columns([2, 1])
    with col_a:
        selected_label = st.selectbox(
            "پلتفرم فعال",
            options=list(platforms.keys()),
            index=list(platforms.keys()).index(default_label),
            key="calc_platform",
        )
    platform_id = platforms[selected_label]
    if active is None or platform_id != active["id"]:
        store.set_active_platform(platform_id)
        active = store.get_platform(platform_id)

    assert active is not None

    with col_b:
        currency_label = st.radio(
            "واحد پول ورودی",
            options=["تومان", "ریال"],
            horizontal=True,
            index=0 if active.get("default_currency", "toman") == "toman" else 1,
            key="calc_currency",
        )
    currency: Currency = "toman" if currency_label == "تومان" else "rial"

    input_mode = st.radio(
        "ورودی قیمت",
        options=["قیمت لیست تولیدکننده", "قیمت خرید واقعی"],
        horizontal=True,
        key="calc_input_mode",
        help="اگر قیمت خرید واقعی‌ات را می‌دانی، همان را وارد کن؛ در این حالت تخفیف خرید اعمال نمی‌شود.",
    )
    buy_mode = input_mode == "قیمت خرید واقعی"
    # In buy-price mode no partner discount is applied on top of the entered price.
    calc_settings = {**active, "partner_discount_pct": 0.0} if buy_mode else active

    raw_price = st.text_input(
        f"{input_mode} ({currency_label})",
        key="calc_base_price_text",
        placeholder="مثلاً 1,020,000",
        on_change=_format_money_input,
        args=("calc_base_price_text",),
    )

    if not raw_price.strip():
        st.info("قیمت را وارد کنید تا محاسبه به‌صورت لحظه‌ای نمایش داده شود.")
        return
    try:
        base_price = parse_number(raw_price)
    except ValueError:
        st.error("قیمت نامعتبر است. فقط عدد وارد کنید (جداکننده هزارگان مجاز است).")
        return
    if base_price <= 0:
        st.info("قیمت باید بیشتر از صفر باشد.")
        return

    try:
        result = calculate_price(base_price, calc_settings, currency=currency)
    except ValueError as exc:
        st.error(str(exc))
        return

    st.subheader("تفکیک هزینه")
    left, right = st.columns(2)

    with left:
        if not buy_mode:
            render_amount_card("قیمت لیست", result.base_price, currency)
        if not buy_mode and result.partner_discount_pct > 0:
            render_amount_card(
                f"تخفیف خرید ({result.partner_discount_pct:g}٪)",
                result.partner_discount_amount,
                currency,
            )
        render_amount_card("قیمت خرید شما", result.effective_base, currency)
        render_amount_card(
            f"سود هدف ({result.profit_pct:g}٪ از قیمت خرید)",
            result.target_net - result.effective_base,
            currency,
        )
        render_amount_card(
            f"سود خالص واقعی ({result.actual_profit_pct:.1f}٪ از قیمت خرید)",
            result.profit,
            currency,
        )
        render_amount_card(
            "قیمت فروش سایت (بدون کارمزد پلتفرم)",
            result.site_price,
            currency,
        )

    with right:
        render_amount_card(
            "قیمت فروش در پلتفرم",
            result.sell_price,
            currency,
            highlight=True,
        )
        render_amount_card(
            f"کمیسیون پلتفرم ({result.platform_fee_pct:g}٪ از قیمت فروش)",
            result.platform_fee,
            currency,
        )
        render_amount_card(
            f"هزینه توسعه زیرساخت ({result.infra_fee_pct:g}٪ از قیمت فروش)",
            result.infra_fee,
            currency,
        )
        render_amount_card(
            f"مالیات ارزش افزوده ({result.vat_pct:g}٪)",
            result.vat_amount,
            currency,
        )
        render_amount_card(
            "هزینه ارسال اکسپرس تهران/کرج",
            result.shipping,
            currency,
            badge=result.shipping_badge,
        )
        render_amount_card(
            "مبلغ خالص دریافتی از پلتفرم",
            result.net_received,
            currency,
        )

    st.subheader("قیمت‌گذاری رقابتی")
    cc1, cc2 = st.columns(2)
    with cc1:
        raw_comp = st.text_input(
            f"کمترین قیمت رقیب ({currency_label})",
            key="calc_competitor_text",
            placeholder="اختیاری، مثلاً 580,000",
            on_change=_format_money_input,
            args=("calc_competitor_text",),
        )
    with cc2:
        floor_pct = st.number_input(
            "کف سود (٪ از قیمت خرید)",
            min_value=0.0,
            value=float(active.get("floor_profit_pct", 3.0)),
            step=0.5,
            key=f"calc_floor_pct_{platform_id}",
        )

    competitor = None
    if raw_comp.strip():
        try:
            competitor = parse_number(raw_comp)
        except ValueError:
            st.error("قیمت رقیب نامعتبر است. فقط عدد وارد کنید.")

    try:
        quote = calculate_competitive_price(
            base_price, calc_settings, competitor, floor_profit_pct=floor_pct, currency=currency
        )
    except ValueError as exc:
        st.error(str(exc))
        return

    q_left, q_right = st.columns(2)
    with q_left:
        render_amount_card(
            f"قیمت کف (سود {quote.floor_profit_pct:g}٪)",
            quote.floor.sell_price,
            currency,
        )
        if competitor:
            render_amount_card("قیمت رقیب", competitor, currency)
    with q_right:
        render_amount_card(
            "قیمت پیشنهادی",
            quote.recommended_price,
            currency,
            highlight=True,
        )
        render_amount_card(
            f"سود خالص در قیمت پیشنهادی ({quote.recommended.actual_profit_pct:.1f}٪ از قیمت خرید)",
            quote.recommended.profit,
            currency,
        )

    if quote.status == "follow":
        gap = abs(quote.gap_to_competitor or 0)
        st.success(f"قیمت پیشنهادی {format_amount(gap, currency)} زیر رقیب است و از کف سود بالاتر می‌ماند.")
        if quote.recommended_price > quote.target.sell_price:
            st.info("رقیب گران‌تر از قیمت هدف شماست؛ می‌توانید بالاتر از سود هدف هم بفروشید.")
    elif quote.status == "floor":
        gap = quote.gap_to_competitor or 0
        st.warning(
            f"رقیب زیر کف سود شماست. قیمت کف پیشنهاد می‌شود که {format_amount(gap, currency)} "
            "گران‌تر از رقیب است؛ دنبال قیمت رقیب نروید (یا کف را آگاهانه پایین بیاورید)."
        )
    else:
        st.caption("قیمت رقیب را وارد کنید تا قیمت پیشنهادی نسبت به آن محاسبه شود؛ بدون آن، قیمت هدف نمایش داده می‌شود.")

    st.subheader("سود در تعداد مختلف سفارش")
    st.caption(
        "هزینه ارسال به ازای هر سفارش (کل مبلغ سفارش) گرفته می‌شود، نه هر عدد؛ پس سود یک قیمت ثابت "
        "با تعداد سفارش فرق می‌کند."
    )
    t1, t2 = st.columns(2)
    with t1:
        raw_check = st.text_input(
            f"قیمت فروش هر عدد برای بررسی ({currency_label})",
            key="calc_check_price_text",
            placeholder="خالی = قیمت پیشنهادی",
            on_change=_format_money_input,
            args=("calc_check_price_text",),
        )
    with t2:
        seller_discount_pct = st.number_input(
            "تخفیف فروشنده در پنل (٪)",
            min_value=0.0,
            max_value=99.0,
            value=0.0,
            step=0.5,
            key=f"calc_seller_discount_{platform_id}",
        )

    check_price = quote.recommended_price
    if raw_check.strip():
        try:
            check_price = parse_number(raw_check)
        except ValueError:
            st.error("قیمت بررسی نامعتبر است.")
    if check_price > 0:
        if seller_discount_pct > 0:
            try:
                list_p, final_p = list_price_for_discount(
                    check_price, seller_discount_pct, calc_settings, currency
                )
                d1, d2 = st.columns(2)
                with d1:
                    render_amount_card("قیمت قبل از تخفیف (در پنل بزنید)", list_p, currency)
                with d2:
                    render_amount_card(
                        f"قیمت نهایی مشتری بعد از {seller_discount_pct:g}٪ تخفیف",
                        final_p,
                        currency,
                    )
                st.caption("کمیسیون، مالیات و ارسال روی قیمت نهایی (بعد از تخفیف) حساب می‌شوند؛ جدول زیر هم با همان.")
                check_price = final_p
            except ValueError as exc:
                st.error(str(exc))

        typical = int(calc_settings.get("bundle_quantity", 3))
        quantities = sorted({1, 2, 3, 5, 10, typical})
        rows = []
        for q in quantities:
            r_q = calculate_price(
                base_price * q, calc_settings, currency=currency, fixed_sell_price=check_price * q
            )
            share = r_q.shipping / r_q.sell_price * 100 if r_q.sell_price else 0.0
            rows.append(
                {
                    "تعداد در سفارش": q,
                    "مبلغ سفارش": f"{r_q.sell_price:,.0f}",
                    "هزینه ارسال": f"{r_q.shipping:,.0f}",
                    "سهم ارسال از فروش ٪": f"{share:.1f}",
                    "سود خالص سفارش": f"{r_q.profit:,.0f}",
                    "سود هر عدد": f"{r_q.profit / q:,.0f}",
                    "سود از قیمت خرید ٪": f"{r_q.actual_profit_pct:.1f}",
                }
            )
        st.dataframe(rows, use_container_width=True, hide_index=True)
        loss_qty = [r["تعداد در سفارش"] for r in rows if float(r["سود خالص سفارش"].replace(",", "")) < 0]
        if loss_qty:
            st.warning(
                "در این قیمت، سفارش‌های " + "، ".join(str(q) for q in loss_qty) + " عددی زیان‌ده‌اند."
            )

    st.subheader("فروش باندل (چند عدد در یک سفارش)")
    qty = st.number_input(
        "تعداد در باندل",
        min_value=1,
        value=int(active.get("bundle_quantity", 3)),
        step=1,
        key=f"calc_bundle_qty_{platform_id}",
    )
    try:
        bun = calculate_bundle(base_price, calc_settings, int(qty), currency=currency)
    except ValueError as exc:
        st.error(str(exc))
        return

    b_left, b_right = st.columns(2)
    with b_left:
        render_amount_card(
            f"مبلغ کل باندل ({bun.quantity} عددی)",
            bun.bundle.sell_price,
            currency,
            highlight=True,
        )
        render_amount_card(
            "قیمت هر عدد در باندل",
            bun.unit_price_in_bundle,
            currency,
        )
        render_text_card(
            "کاهش قیمت هر عدد نسبت به فروش تکی",
            f"{bun.saving_pct:.2f}٪",
        )
    with b_right:
        render_amount_card(
            "هزینه ارسال کل باندل (یک‌بار)",
            bun.bundle.shipping,
            currency,
            badge=bun.bundle.shipping_badge,
        )
        render_amount_card(
            f"سود خالص باندل ({bun.bundle.actual_profit_pct:.1f}٪ از قیمت خرید)",
            bun.bundle.profit,
            currency,
        )
        if bun.min_qty_for_max_shipping > 0:
            render_text_card(
                "حداقل تعداد در یک سفارش برای رسیدن به سقف ارسال",
                f"{bun.min_qty_for_max_shipping} عدد",
            )

    with st.expander("جزئیات تنظیمات پلتفرم فعال"):
        st.write(
            {
                "پلتفرم": active["name"],
                "کمیسیون": f"{active['platform_fee_pct']}%",
                "هزینه زیرساخت": f"{active.get('infra_fee_pct', 0)}%",
                "مالیات ارزش افزوده": f"{active.get('vat_pct', 0)}%"
                + ("" if active.get("vat_on_infra", True) else " (فقط روی کمیسیون)"),
                "سود هدف (از قیمت خرید)": f"{active['profit_pct']}%",
                "تخفیف خرید": f"{active['partner_discount_pct']}%",
                "مارک‌آپ سایت": f"{active['site_markup_pct']}%",
                "درصد ارسال": f"{active['shipping_pct']}%",
                "کف ارسال (تومان)": f"{active['shipping_floor_toman']:,.0f}",
                "سقف ارسال (تومان)": f"{active['shipping_ceiling_toman']:,.0f}",
                "گرد کردن قیمت (تومان)": f"{active.get('rounding_toman', 1000):,.0f}",
                "تعداد پیش‌فرض باندل": f"{active.get('bundle_quantity', 3)}",
                "کف سود": f"{active.get('floor_profit_pct', 3)}%",
                "گام زیر رقیب (تومان)": f"{active.get('competitor_step_toman', 1000):,.0f}",
            }
        )


def tab_batch(store: SettingsStore) -> None:
    platforms = platform_options(store)
    if not platforms:
        st.warning("هیچ پلتفرمی تعریف نشده است.")
        return

    active = store.get_active_platform()
    default_label = next(
        (name for name, pid in platforms.items() if active and pid == active["id"]),
        next(iter(platforms)),
    )

    c1, c2 = st.columns([2, 1])
    with c1:
        selected_label = st.selectbox(
            "پلتفرم برای پردازش",
            options=list(platforms.keys()),
            index=list(platforms.keys()).index(default_label),
            key="batch_platform",
        )
    platform = store.get_platform(platforms[selected_label])
    assert platform is not None

    with c2:
        currency_label = st.radio(
            "واحد قیمت فایل",
            options=["تومان", "ریال"],
            horizontal=True,
            index=0 if platform.get("default_currency", "toman") == "toman" else 1,
            key="batch_currency",
        )
    currency: Currency = "toman" if currency_label == "تومان" else "rial"

    bundle_qty = st.number_input(
        "تعداد در باندل",
        min_value=1,
        value=int(platform.get("bundle_quantity", 3)),
        step=1,
        key=f"batch_bundle_qty_{platform['id']}",
    )

    uploaded = st.file_uploader(
        "آپلود فایل CSV یا Excel",
        type=["csv", "xlsx", "xls"],
        key="batch_upload",
    )

    if not uploaded:
        st.caption("می‌توانید از فایل نمونه `sample_products.csv` در ریشه پروژه استفاده کنید.")
        st.caption(
            "اختیاری: اگر فایل ستونی با نام شامل «کمیسیون» یا «زیرساخت» (به‌صورت درصد، مثلاً 9 یا 0.6) "
            "داشته باشد، برای هر ردیف به‌جای مقدار پیش‌فرض پلتفرم استفاده می‌شود."
        )
        st.caption(
            "اختیاری: ستونی با نام شامل «رقیب» (کمترین قیمت رقیب) ستون‌های قیمت پیشنهادی و وضعیت رقابت را اضافه می‌کند. "
            "ستونی با نام «کف حاشیه» (درصد) کف سود را برای همان ردیف عوض می‌کند."
        )
        return

    try:
        df = read_upload(uploaded.getvalue(), uploaded.name)
    except Exception as exc:  # noqa: BLE001
        st.error(f"خواندن فایل ناموفق بود: {exc}")
        return

    st.write("پیش‌نمایش داده ورودی:")
    st.dataframe(df.head(20), use_container_width=True)

    detected = detect_base_price_column(list(df.columns.astype(str)))
    col_options = list(df.columns.astype(str))
    default_idx = col_options.index(detected) if detected in col_options else 0

    base_col = st.selectbox(
        "ستون قیمت لیست (پایه)",
        options=col_options,
        index=default_idx,
        key="batch_base_col",
    )

    if st.button("محاسبه و آماده‌سازی خروجی", type="primary", key="batch_run"):
        try:
            out = process_dataframe(
                df,
                platform,
                base_price_column=base_col,
                currency=currency,
                bundle_quantity=int(bundle_qty),
            )
            st.session_state.batch_result = out
            st.success("محاسبه با موفقیت انجام شد.")
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))
            return

    if "batch_result" in st.session_state:
        out_df = st.session_state.batch_result
        st.write("نتیجه:")
        st.dataframe(out_df, use_container_width=True)
        xlsx = to_xlsx_bytes(out_df)
        st.download_button(
            label="دانلود فایل Excel",
            data=xlsx,
            file_name="priced_products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="batch_download",
        )


def tab_settings(store: SettingsStore) -> None:
    platforms = store.list_platforms()
    names = {p["id"]: p["name"] for p in platforms}

    st.subheader("مدیریت پلتفرم‌ها")
    left, right = st.columns([1, 2])

    with left:
        selected_id = st.radio(
            "انتخاب پلتفرم",
            options=list(names.keys()),
            format_func=lambda pid: names[pid],
            key="settings_platform_radio",
        )
        if st.button("کپی پلتفرم", key="dup_platform"):
            try:
                copied = store.duplicate_platform(selected_id)
                st.success(f"کپی ساخته شد: {copied['name']}")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

        if st.button("حذف پلتفرم", key="del_platform"):
            try:
                store.delete_platform(selected_id)
                st.success("پلتفرم حذف شد.")
                st.rerun()
            except ValueError as exc:
                st.error(str(exc))

        st.divider()
        st.markdown("**افزودن پلتفرم جدید**")
        new_name = st.text_input("نام پلتفرم جدید", key="new_platform_name")
        if st.button("افزودن", key="add_platform"):
            if not new_name.strip():
                st.error("نام را وارد کنید.")
            else:
                store.upsert_platform({"name": new_name.strip()})
                st.success("پلتفرم اضافه شد.")
                st.rerun()

    platform = store.get_platform(selected_id)
    if not platform:
        st.warning("پلتفرم یافت نشد.")
        return

    with right:
        st.markdown(f"### ویرایش: {platform['name']}")
        with st.form(f"edit_form_{selected_id}"):
            name = st.text_input("نام", value=platform["name"])
            c1, c2 = st.columns(2)
            with c1:
                platform_fee_pct = st.number_input(
                    "کمیسیون پلتفرم (٪ از قیمت فروش)",
                    min_value=0.0,
                    value=float(platform["platform_fee_pct"]),
                    step=0.5,
                )
                infra_fee_pct = st.number_input(
                    "هزینه توسعه زیرساخت (٪ از قیمت فروش)",
                    min_value=0.0,
                    value=float(platform.get("infra_fee_pct", 0.0)),
                    step=0.1,
                    format="%.2f",
                )
                vat_pct = st.number_input(
                    "مالیات بر ارزش افزوده (٪ روی کمیسیون)",
                    min_value=0.0,
                    value=float(platform.get("vat_pct", 9.0)),
                    step=0.5,
                )
                vat_on_infra = st.checkbox(
                    "مالیات روی هزینه زیرساخت هم اعمال شود",
                    value=bool(platform.get("vat_on_infra", True)),
                )
                profit_pct = st.number_input(
                    "سود هدف (٪ نسبت به قیمت خرید)",
                    min_value=0.0,
                    value=float(platform["profit_pct"]),
                    step=0.5,
                )
                partner_discount_pct = st.number_input(
                    "تخفیف خرید از قیمت لیست (٪)",
                    min_value=0.0,
                    max_value=99.9,
                    value=float(platform["partner_discount_pct"]),
                    step=0.5,
                )
            with c2:
                shipping_pct = st.number_input(
                    "درصد هزینه ارسال (٪ از قیمت فروش)",
                    min_value=0.0,
                    value=float(platform["shipping_pct"]),
                    step=0.5,
                )
                shipping_floor = st.number_input(
                    "کف ارسال (تومان)",
                    min_value=0.0,
                    value=float(platform["shipping_floor_toman"]),
                    step=1000.0,
                    format="%.0f",
                )
                shipping_ceiling = st.number_input(
                    "سقف ارسال (تومان)",
                    min_value=0.0,
                    value=float(platform["shipping_ceiling_toman"]),
                    step=1000.0,
                    format="%.0f",
                )
                rounding_toman = st.number_input(
                    "گرد کردن قیمت فروش رو به بالا (تومان)",
                    min_value=0.0,
                    value=float(platform.get("rounding_toman", 1000)),
                    step=100.0,
                    format="%.0f",
                )
                bundle_quantity = st.number_input(
                    "تعداد پیش‌فرض در باندل",
                    min_value=1,
                    value=int(platform.get("bundle_quantity", 3)),
                    step=1,
                )
                floor_profit_pct = st.number_input(
                    "کف سود در قیمت‌گذاری رقابتی (٪ از قیمت خرید)",
                    min_value=0.0,
                    value=float(platform.get("floor_profit_pct", 3.0)),
                    step=0.5,
                )
                competitor_step = st.number_input(
                    "گام زیر قیمت رقیب (تومان)",
                    min_value=0.0,
                    value=float(platform.get("competitor_step_toman", 1000)),
                    step=100.0,
                    format="%.0f",
                )
                site_markup_pct = st.number_input(
                    "مارک‌آپ قیمت فروش سایت (٪)",
                    min_value=0.0,
                    value=float(platform["site_markup_pct"]),
                    step=0.5,
                )
                default_currency = st.selectbox(
                    "واحد پول پیش‌فرض",
                    options=["toman", "rial"],
                    index=0 if platform.get("default_currency") == "toman" else 1,
                    format_func=lambda x: "تومان" if x == "toman" else "ریال",
                )

            submitted = st.form_submit_button("ذخیره تنظیمات", type="primary")
            if submitted:
                try:
                    store.upsert_platform(
                        {
                            "id": selected_id,
                            "name": name.strip() or platform["name"],
                            "platform_fee_pct": platform_fee_pct,
                            "infra_fee_pct": infra_fee_pct,
                            "vat_pct": vat_pct,
                            "vat_on_infra": vat_on_infra,
                            "profit_pct": profit_pct,
                            "partner_discount_pct": partner_discount_pct,
                            "site_markup_pct": site_markup_pct,
                            "shipping_pct": shipping_pct,
                            "shipping_floor_toman": shipping_floor,
                            "shipping_ceiling_toman": shipping_ceiling,
                            "rounding_toman": rounding_toman,
                            "bundle_quantity": int(bundle_quantity),
                            "floor_profit_pct": floor_profit_pct,
                            "competitor_step_toman": competitor_step,
                            "default_currency": default_currency,
                        }
                    )
                    st.success("تنظیمات ذخیره شد.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))


_PCT_COLS = {"commission_pct", "infra_pct", "deduction_pct", "profit_pct", "shipping_pct"}
_TEXT_COLS = {"title", "category", "order", "order_date", "basis_date"}


def _display_table(df: pd.DataFrame, names: dict[str, str]) -> pd.DataFrame:
    """Format a report table for display: thousand separators, Persian headers."""
    out = pd.DataFrame()
    for key, label in names.items():
        if key not in df.columns:
            continue
        col = df[key]
        if key in _TEXT_COLS:
            out[label] = col.astype(str)
        elif key in _PCT_COLS:
            out[label] = col.map(lambda v: "" if pd.isna(v) else f"{v:.1f}")
        else:
            out[label] = col.map(lambda v: "" if pd.isna(v) else f"{v:,.0f}")
    return out


def tab_report(store: SettingsStore) -> None:
    st.subheader("تحلیل گزارش تسویه اسنپ‌شاپ")
    st.caption(
        "گزارش تسویه (فایل reports-summary...xlsx از پنل فروشنده) را آپلود کنید تا کسورات واقعی هر کالا و "
        "هر سفارش، و با وارد کردن قیمت خرید، سود واقعی را ببینید."
    )
    uploaded = st.file_uploader("آپلود گزارش تسویه (xlsx)", type=["xlsx"], key="report_upload")
    if uploaded is None:
        st.info("هنوز فایلی آپلود نشده است.")
        return

    try:
        report = parse_settlement_report(uploaded.getvalue())
    except Exception as exc:  # noqa: BLE001 - show any parsing problem to the user
        st.error(f"خواندن گزارش ممکن نشد: {exc}")
        return

    tot = report.totals
    deductions = tot["commission"] + tot["infra"] + tot["shipping"]
    st.markdown(f"**دوره:** {report.period}")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("جمع فروش (تومان)", f"{tot['sales']:,.0f}")
    m2.metric(
        "جمع کسورات (تومان)",
        f"{deductions:,.0f}",
        f"{deductions / tot['sales'] * 100:.1f}٪ از فروش" if tot["sales"] else None,
        delta_color="off",
    )
    m3.metric("خالص دریافتی (تومان)", f"{tot['net']:,.0f}")
    m4.metric("سفارش / عدد", f"{tot['orders']:,.0f} / {tot['units']:,.0f}")

    if report.payout_reported is not None:
        diff = report.payout_reported - tot["net"]
        if abs(diff) < 5:
            st.success("محاسبه این تب با «مانده طی دوره» گزارش یکی است.")
        else:
            st.warning(
                f"مانده طی دوره در گزارش {report.payout_reported:,.0f} تومان است و با محاسبه این تب "
                f"{diff:,.0f} تومان اختلاف دارد (احتمالاً موارد دیگری مثل برگشتی یا جریمه)."
            )

    if report.adjustments:
        st.warning(
            "این موارد در گزارش ثبت شده‌اند و در جدول‌های پایین لحاظ نشده‌اند: "
            + "؛ ".join(f"{name} ({rows} ردیف، {amount:,.0f} تومان)" for name, rows, amount in report.adjustments)
        )

    active = store.get_active_platform()
    if active:
        warnings = compare_with_settings(report, active)
        with st.expander(f"مغایرت با تنظیمات «{active['name']}» ({len(warnings)} مورد)", expanded=bool(warnings)):
            if warnings:
                for w in warnings:
                    st.warning(w)
            else:
                st.success("نرخ‌های گزارش با تنظیمات پلتفرم فعال می‌خواند.")

    # ---- buy prices -> real profit ----------------------------------------
    st.subheader("سود واقعی")
    st.caption("قیمت خرید هر عدد را وارد کنید (به‌صورت خودکار برای دفعه بعد قابل ذخیره است).")
    buy_store = BuyPriceStore()
    saved = buy_store.all()
    buy_label = "قیمت خرید هر عدد (تومان)"
    editor_df = pd.DataFrame(
        {
            "کالا": report.products["title"],
            buy_label: report.products["title"].map(saved).astype("float64"),
        }
    )
    edited = st.data_editor(
        editor_df,
        hide_index=True,
        use_container_width=True,
        disabled=["کالا"],
        num_rows="fixed",
        key=f"report_buy_editor_{uploaded.name}_{uploaded.size}",
        column_config={buy_label: st.column_config.NumberColumn(min_value=0, step=1000, format="%d")},
    )
    prices = {
        title: float(value)
        for title, value in zip(edited["کالا"], edited[buy_label])
        if pd.notna(value) and float(value) > 0
    }
    if st.button("ذخیره قیمت‌های خرید", key="report_save_buy"):
        buy_store.update(dict(zip(edited["کالا"], edited[buy_label])))
        st.success("قیمت‌های خرید ذخیره شد.")

    products = add_profit(report.products, prices)
    known = products[products["buy_price"].notna()]
    if not known.empty:
        p1, p2, p3 = st.columns(3)
        profit_total = known["profit"].sum()
        cost_total = known["cost"].sum()
        p1.metric("سود خالص (کالاهای دارای قیمت خرید)", f"{profit_total:,.0f}")
        p2.metric("سود از هزینه خرید", f"{profit_total / cost_total * 100:.1f}٪" if cost_total else "-")
        p3.metric("پوشش فروش", f"{known['sales'].sum() / tot['sales'] * 100:.0f}٪" if tot["sales"] else "-")
        losers = known[known["profit"] < 0]
        if not losers.empty:
            st.error("کالاهای زیان‌ده: " + "؛ ".join(f"{r.title} ({r.profit:,.0f})" for r in losers.itertuples()))

    # ---- tables ---------------------------------------------------------------
    st.subheader("به تفکیک کالا")
    names = dict(PRODUCT_COLUMNS)
    if known.empty:
        for k in ("buy_price", "cost", "profit", "profit_pct"):
            names.pop(k, None)
    st.dataframe(_display_table(products, names), use_container_width=True, hide_index=True)

    with st.expander("به تفکیک سفارش"):
        st.dataframe(_display_table(report.orders, ORDER_COLUMNS), use_container_width=True, hide_index=True)

    st.download_button(
        "دانلود تحلیل (xlsx)",
        data=analysis_to_xlsx(products, report.orders),
        file_name="settlement_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main() -> None:
    inject_rtl()
    store = get_store()

    st.title("ماشین‌حساب قیمت محصول")
    st.markdown(
        '<p class="hero-sub">محاسبه قیمت فروش در پلتفرم با احتساب کمیسیون، مالیات، ارسال و سود هدف — اجرای لوکال</p>',
        unsafe_allow_html=True,
    )

    tab1, tab2, tab4, tab3 = st.tabs(
        ["ماشین‌حساب تکی", "پردازش دسته‌ای (Excel/CSV)", "تحلیل گزارش تسویه", "تنظیمات پلتفرم‌ها"]
    )
    with tab1:
        tab_calculator(store)
    with tab2:
        tab_batch(store)
    with tab4:
        tab_report(store)
    with tab3:
        tab_settings(store)


if __name__ == "__main__":
    main()
