# -*- coding: utf-8 -*-
"""Хранение настроек приложения (колонки, производители, каталог, список прайсов)."""
from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def default_columns() -> Dict[str, List[str]]:
    return {
        "key_candidates": [
            "Артикул",
            "Код",
            "SKU",
            "ID",
            "PartNumber",
            "НоменклатураКод",
        ],
        "name_candidates": [
            "Наименование",
            "Наименование товаров",
            "Наименование товара",
            "Наименование препарата",
            "Наименование препаратов",
            "Товары",
            "Махсулот номи",
            "Номенклатура",
            "Товар",
            "Название",
            "Name",
            "Наимменование",
        ],
        "price_candidates": [
            "Цена (100%) с НДС",
            "Цена При 100 % оплате цена",
            "При 100 % оплате цена",
            "Тип цены Сотув",
            "Цена С НДС",
            "Цена 100%",
            "Цена с НДС 100%",
            "Цена для клиента с предоплатой 100%",
            "Цена с НДС",
            "Цена без НДС",
            "Цена",
            "Цена Спец мин 1млн",
            "Тип цены При 100% Оплате (перечисление)",
            "При 100% Оплате",
            "Цена договорная",
            "Цена (с НДС)",
            "Цена Договорная",
            "Цена переч",
            "Цена переч:",
            "Цена реал",
            "ПриходЦена",
            "Сумма заказ",
            "Сумма",
            "Отп. Цена",
            "Цена, руб",
            "Price",
            "Нарх",
            "Розница",
            "Опт",
        ],
        "stock_candidates": ["Остаток", "Наличие", "Stock", "Qty", "Количество"],
        "producer_candidates": [
            "Производитель",
            "Бренд",
            "Manufacturer",
            "Ишлаб чикарувчи",
        ],
        "expiry_candidates": ["Срок годности", "Срок", "Годен до", "Expiry"],
    }


@dataclass
class PriceFileEntry:
    filename: str
    display_name: str = ""
    read_preset: str = ""  # "", "khoji_akbar", "grand_farm"

    def to_dict(self) -> Dict[str, str]:
        return {"filename": self.filename, "display_name": self.display_name, "read_preset": self.read_preset or ""}

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "PriceFileEntry":
        return PriceFileEntry(
            filename=str(d.get("filename", "")),
            display_name=str(d.get("display_name", "")),
            read_preset=str(d.get("read_preset", "") or ""),
        )


@dataclass
class AppSettings:
    columns: Dict[str, List[str]] = field(default_factory=default_columns)
    producers: List[str] = field(default_factory=list)
    catalog: List[Dict[str, Any]] = field(default_factory=list)
    price_files: List[PriceFileEntry] = field(default_factory=list)
    use_catalog_matching: bool = False
    legacy_priority_rules: bool = False

    def catalog_as_tuples(self) -> List[Tuple[str, List[str]]]:
        out: List[Tuple[str, List[str]]] = []
        for item in self.catalog:
            canon = str(item.get("canonical", "")).strip()
            if not canon:
                continue
            aliases = item.get("aliases") or []
            if isinstance(aliases, str):
                aliases = [a.strip() for a in aliases.split(";") if a.strip()]
            else:
                aliases = [str(a).strip() for a in aliases if str(a).strip()]
            out.append((canon, aliases))
        return out

    def to_json_dict(self) -> Dict[str, Any]:
        return {
            "columns": self.columns,
            "producers": self.producers,
            "catalog": self.catalog,
            "price_files": [e.to_dict() for e in self.price_files],
            "use_catalog_matching": self.use_catalog_matching,
            "legacy_priority_rules": self.legacy_priority_rules,
        }

    @staticmethod
    def from_json_dict(d: Dict[str, Any]) -> "AppSettings":
        cols = d.get("columns")
        base = default_columns()
        if not isinstance(cols, dict):
            cols = dict(base)
        else:
            merged = dict(base)
            for k, v in base.items():
                if k in cols and isinstance(cols[k], list):
                    merged[k] = [str(x) for x in cols[k]]
            cols = merged
        pf = []
        for row in d.get("price_files") or []:
            if isinstance(row, dict) and row.get("filename"):
                pf.append(PriceFileEntry.from_dict(row))
        cat = d.get("catalog") or []
        if not isinstance(cat, list):
            cat = []
        return AppSettings(
            columns=cols,
            producers=[str(p).strip() for p in (d.get("producers") or []) if str(p).strip()],
            catalog=cat,
            price_files=pf,
            use_catalog_matching=bool(d.get("use_catalog_matching", False)),
            legacy_priority_rules=bool(d.get("legacy_priority_rules", False)),
        )


def app_data_root() -> Path:
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    else:
        base = str(Path.home() / ".local" / "share")
    root = Path(base) / "PriceAggregator"
    return root


def prices_dir() -> Path:
    p = app_data_root() / "prices"
    p.mkdir(parents=True, exist_ok=True)
    return p


def output_dir() -> Path:
    p = app_data_root() / "output"
    p.mkdir(parents=True, exist_ok=True)
    return p


def settings_path() -> Path:
    return app_data_root() / "settings.json"


def load_settings() -> AppSettings:
    path = settings_path()
    if not path.is_file():
        return AppSettings()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return AppSettings()
        return AppSettings.from_json_dict(data)
    except (OSError, json.JSONDecodeError):
        return AppSettings()


def save_settings(settings: AppSettings) -> None:
    root = app_data_root()
    root.mkdir(parents=True, exist_ok=True)
    with open(settings_path(), "w", encoding="utf-8") as f:
        json.dump(settings.to_json_dict(), f, ensure_ascii=False, indent=2)


def unique_target_name(dest_dir: Path, original_name: str) -> Path:
    name = Path(original_name).name
    target = dest_dir / name
    if not target.exists():
        return target
    stem = target.stem
    suf = target.suffix
    n = 1
    while True:
        cand = dest_dir / f"{stem}_{n}{suf}"
        if not cand.exists():
            return cand
        n += 1
