# ماشین‌حساب قیمت محصول

اپلیکیشن لوکال Streamlit برای محاسبه قیمت تمام‌شده محصول روی پلتفرم‌هایی مثل **اسنپ شاپ** و **دیجی کالا**، با پشتیبانی از تنظیمات جداگانه برای هر پلتفرم، تخفیف همکار، هزینه ارسال پویا، و پردازش دسته‌ای Excel/CSV.

اجرای داخلی روی **ویندوز** و **مک** — بدون نیاز به دیپلوی وب.

## امکانات

- محاسبه لحظه‌ای قیمت تکی (ریال و تومان همزمان)
- تفکیک: قیمت پایه، تخفیف همکار، کارمزد پلتفرم، سود، ارسال (با نشان حداقل/حداکثر)، قیمت تمام‌شده، قیمت فروش سایت
- پروفایل جدا برای هر پلتفرم (درصدها + کف/سقف ارسال) ذخیره‌شده در `data/platforms.json`
- آپلود CSV/Excel و دانلود خروجی `.xlsx`
- رابط کاربری فارسی و RTL

## پیش‌نیاز

- Python 3.10 یا بالاتر

## نصب و اجرا

### مک / لینوکس

```bash
cd price_calculator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

### ویندوز (PowerShell)

```powershell
cd price_calculator
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

مرورگر معمولاً روی `http://localhost:8501` باز می‌شود.

## ساختار پروژه

```
app.py                  # رابط کاربری Streamlit
services/
  calculator.py         # موتور محاسبه
  settings.py           # مدیریت پروفایل پلتفرم‌ها
  batch.py              # پردازش Excel/CSV
data/platforms.json     # تنظیمات ذخیره‌شده
sample_products.csv     # نمونه برای تست batch
```

## منطق محاسبه (خلاصه)

1. `effective_base = base_price × (1 − partner_discount%)`
2. کارمزد و سود روی `effective_base`
3. ارسال = clamp(`effective_base × shipping%`, کف، سقف) — کف/سقف بر حسب تومان در تنظیمات
4. `total_cost = effective_base + fee + profit + shipping`
5. `site_price = effective_base × (1 + site_markup%)`

تبدیل واحد: **۱ تومان = ۱۰ ریال**

## تست سریع موتور محاسبه

با فعال بودن venv:

```bash
python -c "
from services.calculator import calculate_price
s = {
  'platform_fee_pct': 10, 'profit_pct': 10, 'partner_discount_pct': 0,
  'site_markup_pct': 10, 'shipping_pct': 7,
  'shipping_floor_toman': 30000, 'shipping_ceiling_toman': 190000,
  'default_currency': 'toman',
}
r = calculate_price(100_000, s, 'toman')
assert r.shipping == 30000  # کف
assert round(r.total_cost) == 100000 + 10000 + 10000 + 30000
r2 = calculate_price(5_000_000, s, 'toman')
assert r2.shipping == 190000  # سقف
print('OK', r.total_cost, r2.shipping)
"
```

## گیت و ارسال به GitHub

ریپازیتوری محلی از قبل آماده است. برای ساخت ریپو روی GitHub و push:

```bash
# روی GitHub یک مخزن خالی بسازید، سپس:
git remote add origin https://github.com/<USERNAME>/<REPO>.git
git branch -M main
git push -u origin main
```

## نکات

- تنظیمات پلتفرم‌ها در `data/platforms.json` ذخیره می‌شوند و بین اجراها باقی می‌مانند.
- ستون قیمت پایه در فایل batch با نام‌هایی مثل `قیمت پایه` یا `base_price` به‌صورت خودکار تشخیص داده می‌شود.
