"""
Platform profile persistence (JSON file).

Each platform stores its own percentages, shipping bounds, and default currency.
"""

from __future__ import annotations

import json
import re
import uuid
from copy import deepcopy
from pathlib import Path
from typing import Any

# Project root: services/ -> parent
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_SETTINGS_PATH = PROJECT_ROOT / "data" / "platforms.json"

DEFAULT_PLATFORM_TEMPLATE: dict[str, Any] = {
    "platform_fee_pct": 9.0,  # commission, % of selling price
    "infra_fee_pct": 0.6,  # infrastructure development fee, % of selling price
    "vat_pct": 9.0,  # VAT charged on the commission (and infra fee if enabled)
    "vat_on_infra": True,
    "profit_pct": 10.0,  # target profit, % of BUY price
    "partner_discount_pct": 0.0,  # discount off the manufacturer list price
    "site_markup_pct": 10.0,
    "shipping_pct": 7.0,  # % of selling price, clamped by floor/ceiling
    "shipping_floor_toman": 30000,
    "shipping_ceiling_toman": 190000,
    "rounding_toman": 1000,  # selling price is rounded UP to a multiple of this
    "bundle_quantity": 3,  # default number of units in a bundle
    "default_currency": "toman",
}

REQUIRED_FIELDS = (
    "id",
    "name",
    "platform_fee_pct",
    "infra_fee_pct",
    "vat_pct",
    "vat_on_infra",
    "profit_pct",
    "partner_discount_pct",
    "site_markup_pct",
    "shipping_pct",
    "shipping_floor_toman",
    "shipping_ceiling_toman",
    "rounding_toman",
    "bundle_quantity",
    "default_currency",
)


def _slugify(name: str) -> str:
    """Build a simple ASCII-ish id from a display name (fallback to uuid)."""
    ascii_part = re.sub(r"[^a-zA-Z0-9]+", "-", name.strip().lower()).strip("-")
    if ascii_part:
        return ascii_part[:40]
    return f"platform-{uuid.uuid4().hex[:8]}"


def _ensure_defaults(data: dict[str, Any]) -> dict[str, Any]:
    """Normalize loaded JSON and fill missing keys."""
    platforms = data.get("platforms") or []
    normalized: list[dict[str, Any]] = []
    for item in platforms:
        platform = {**DEFAULT_PLATFORM_TEMPLATE, **item}
        if not platform.get("id"):
            platform["id"] = _slugify(str(platform.get("name", "platform")))
        if platform.get("default_currency") not in ("rial", "toman"):
            platform["default_currency"] = "toman"
        normalized.append(platform)

    active = data.get("active_platform_id")
    ids = {p["id"] for p in normalized}
    if not active or active not in ids:
        active = normalized[0]["id"] if normalized else None

    return {"active_platform_id": active, "platforms": normalized}


class SettingsStore:
    """Load / save / CRUD platform profiles on disk."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = Path(path) if path else DEFAULT_SETTINGS_PATH
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            self._write(self._seed())
        self._data = _ensure_defaults(self._read())

    def _seed(self) -> dict[str, Any]:
        return {
            "active_platform_id": "snappshop",
            "platforms": [
                {
                    "id": "snappshop",
                    "name": "اسنپ شاپ",
                    **DEFAULT_PLATFORM_TEMPLATE,
                },
                {
                    "id": "digikala",
                    "name": "دیجی کالا",
                    **DEFAULT_PLATFORM_TEMPLATE,
                },
            ],
        }

    def _read(self) -> dict[str, Any]:
        with self.path.open("r", encoding="utf-8") as fh:
            return json.load(fh)

    def _write(self, data: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.open("w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
            fh.write("\n")

    def reload(self) -> None:
        self._data = _ensure_defaults(self._read())

    def save(self) -> None:
        self._write(self._data)

    @property
    def data(self) -> dict[str, Any]:
        return self._data

    def list_platforms(self) -> list[dict[str, Any]]:
        return deepcopy(self._data["platforms"])

    def get_platform(self, platform_id: str) -> dict[str, Any] | None:
        for platform in self._data["platforms"]:
            if platform["id"] == platform_id:
                return deepcopy(platform)
        return None

    def get_active_platform(self) -> dict[str, Any] | None:
        active_id = self._data.get("active_platform_id")
        if not active_id:
            return None
        return self.get_platform(active_id)

    def set_active_platform(self, platform_id: str) -> None:
        if not self.get_platform(platform_id):
            raise ValueError(f"پلتفرم یافت نشد: {platform_id}")
        self._data["active_platform_id"] = platform_id
        self.save()

    def upsert_platform(self, platform: dict[str, Any]) -> dict[str, Any]:
        """Create or update a platform profile. Returns the saved copy."""
        payload = {**DEFAULT_PLATFORM_TEMPLATE, **platform}
        if not payload.get("name"):
            raise ValueError("نام پلتفرم الزامی است.")
        if not payload.get("id"):
            payload["id"] = _slugify(str(payload["name"]))
            # Ensure uniqueness
            existing_ids = {p["id"] for p in self._data["platforms"]}
            base_id = payload["id"]
            suffix = 2
            while payload["id"] in existing_ids:
                payload["id"] = f"{base_id}-{suffix}"
                suffix += 1

        floor = float(payload["shipping_floor_toman"])
        ceiling = float(payload["shipping_ceiling_toman"])
        if floor < 0 or ceiling < 0:
            raise ValueError("کف و سقف ارسال نمی‌توانند منفی باشند.")
        if floor > ceiling:
            raise ValueError("کف ارسال نمی‌تواند از سقف بیشتر باشد.")

        if payload["default_currency"] not in ("rial", "toman"):
            raise ValueError("واحد پول باید rial یا toman باشد.")

        for key in (
            "platform_fee_pct",
            "infra_fee_pct",
            "vat_pct",
            "profit_pct",
            "partner_discount_pct",
            "site_markup_pct",
            "shipping_pct",
        ):
            value = float(payload[key])
            if value < 0:
                raise ValueError(f"مقدار {key} نمی‌تواند منفی باشد.")
            payload[key] = value

        rounding = float(payload["rounding_toman"])
        if rounding < 0:
            raise ValueError("مقدار گرد کردن نمی‌تواند منفی باشد.")

        if payload["partner_discount_pct"] >= 100:
            raise ValueError("تخفیف خرید باید کمتر از ۱۰۰٪ باشد.")

        # Commission + infra fee (with VAT) + shipping % must leave something over.
        vat = payload["vat_pct"] / 100.0
        k = (
            1.0
            - payload["platform_fee_pct"] / 100.0 * (1.0 + vat)
            - payload["infra_fee_pct"] / 100.0 * (1.0 + (vat if payload["vat_on_infra"] else 0.0))
        )
        if k <= 0:
            raise ValueError("مجموع کمیسیون، هزینه زیرساخت و مالیات باید کمتر از ۱۰۰٪ باشد.")

        bundle_qty = int(float(payload["bundle_quantity"]))
        if bundle_qty < 1:
            raise ValueError("تعداد در باندل باید حداقل ۱ باشد.")
        payload["bundle_quantity"] = bundle_qty

        payload["vat_on_infra"] = bool(payload["vat_on_infra"])
        payload["rounding_toman"] = rounding
        payload["shipping_floor_toman"] = floor
        payload["shipping_ceiling_toman"] = ceiling

        replaced = False
        for idx, existing in enumerate(self._data["platforms"]):
            if existing["id"] == payload["id"]:
                self._data["platforms"][idx] = payload
                replaced = True
                break
        if not replaced:
            self._data["platforms"].append(payload)
            if not self._data.get("active_platform_id"):
                self._data["active_platform_id"] = payload["id"]

        self.save()
        return deepcopy(payload)

    def delete_platform(self, platform_id: str) -> None:
        platforms = self._data["platforms"]
        if len(platforms) <= 1:
            raise ValueError("حداقل یک پلتفرم باید باقی بماند.")
        new_list = [p for p in platforms if p["id"] != platform_id]
        if len(new_list) == len(platforms):
            raise ValueError(f"پلتفرم یافت نشد: {platform_id}")
        self._data["platforms"] = new_list
        if self._data.get("active_platform_id") == platform_id:
            self._data["active_platform_id"] = new_list[0]["id"]
        self.save()

    def duplicate_platform(self, platform_id: str, new_name: str | None = None) -> dict[str, Any]:
        source = self.get_platform(platform_id)
        if not source:
            raise ValueError(f"پلتفرم یافت نشد: {platform_id}")
        copy = deepcopy(source)
        copy.pop("id", None)
        copy["name"] = new_name or f"{source['name']} (کپی)"
        return self.upsert_platform(copy)
