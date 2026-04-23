# -*- coding: utf-8 -*-
"""Запуск сводного прайса, вертикали, ассортимента и подсветки из настроек приложения."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import openpyxl
import pandas as pd

import assortment_prices as assortment
import best_suppliers_analysis as best_suppliers
import cheapest_svodny_prices as cheapest
from app_settings import AppSettings, output_dir, prices_dir
from svodny_price import (
    build_matrix_from_extracted,
    extract_materia_any,
    format_summary_sheet,
    friendly_source_label,
    fuzzy_merge_canonicals,
    read_price_table,
)
from svodny_price_vertical import build_matrix_from_extracted as build_vertical_matrix
from svodny_price_vertical import format_summary_sheet as format_vertical_sheet


def _column_lists(s: AppSettings) -> Tuple[list, list, list, list, list]:
    c = s.columns
    return (
        list(c.get("key_candidates", [])),
        list(c.get("name_candidates", [])),
        list(c.get("price_candidates", [])),
        list(c.get("stock_candidates", [])),
        list(c.get("producer_candidates", [])),
    )


def _display_label(entry_display: str, path: Path) -> str:
    d = (entry_display or "").strip()
    if d:
        return d
    return friendly_source_label(path.stem)


def _extract_matrix_rows(
    s: AppSettings,
    progress: Optional[Callable[[str], None]] = None,
) -> Tuple[List[dict], List[Tuple[str, List[str]]], List[str]]:
    catalog = s.catalog_as_tuples()
    if s.use_catalog_matching and not catalog:
        raise ValueError("Добавьте препараты в настройках (каталог: название и синонимы).")
    if not s.producers:
        raise ValueError("Добавьте хотя бы одну подстроку производителя в настройках.")

    key_c, name_c, price_c, stock_c, prod_c = _column_lists(s)
    pdir = prices_dir()
    entries = [e for e in s.price_files if (pdir / e.filename).is_file()]
    if not entries:
        raise ValueError("Нет файлов прайсов. Нажмите «Загрузить прайсы».")

    all_labels: List[str] = []
    for e in entries:
        path = pdir / e.filename
        all_labels.append(_display_label(e.display_name, path))
    all_labels = sorted(dict.fromkeys(all_labels))

    extracted: List[dict] = []
    total = len(entries)
    for i, e in enumerate(entries, 1):
        path = pdir / e.filename
        if progress:
            progress(f"[{i}/{total}] {path.name}")
        preset = (e.read_preset or "").strip() or None
        df = read_price_table(
            path,
            key_c,
            name_c,
            price_c,
            stock_c,
            use_key=True,
            use_name=True,
            read_preset=preset,
        )
        lbl = _display_label(e.display_name, path)
        extracted.extend(
            extract_materia_any(
                df,
                path,
                key_c,
                name_c,
                price_c,
                stock_c,
                prod_c,
                catalog if s.use_catalog_matching else [],
                source_label=lbl,
                producer_substrings=s.producers,
                legacy_priority_rules=s.legacy_priority_rules,
            )
        )
    return fuzzy_merge_canonicals(extracted, threshold=82), (catalog if s.use_catalog_matching else []), all_labels


def _save_matrix(df: pd.DataFrame, out: Path, vertical: bool) -> Path:
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out, index=False, engine="openpyxl")
    wb = openpyxl.load_workbook(out)
    for sheet in wb.worksheets:
        if vertical:
            format_vertical_sheet(sheet)
        else:
            format_summary_sheet(sheet)
    wb.save(out)
    return out


def run_svodny_horizontal(
    s: AppSettings,
    progress: Optional[Callable[[str], None]] = None,
) -> Path:
    extracted, catalog, labels = _extract_matrix_rows(s, progress)
    matrix = build_matrix_from_extracted(extracted, catalog, labels)
    today = datetime.now().strftime("%d.%m.%Y")
    out = output_dir() / f"Сводный прайс {today}.xlsx"
    return _save_matrix(matrix, out, vertical=False)


def run_svodny_vertical(
    s: AppSettings,
    progress: Optional[Callable[[str], None]] = None,
) -> Path:
    extracted, catalog, labels = _extract_matrix_rows(s, progress)
    matrix = build_vertical_matrix(extracted, catalog, labels)
    today = datetime.now().strftime("%d.%m.%Y")
    out = output_dir() / f"Сводный прайс вертикально {today}.xlsx"
    return _save_matrix(matrix, out, vertical=True)


def run_assortment(
    s: AppSettings,
    progress: Optional[Callable[[str], None]] = None,
) -> Path:
    key_c, name_c, price_c, stock_c, _ = _column_lists(s)
    pdir = prices_dir()
    entries = [e for e in s.price_files if (pdir / e.filename).is_file()]
    if not entries:
        raise ValueError("Нет файлов прайсов.")
    rows: List[dict] = []
    total = len(entries)
    for i, e in enumerate(entries, 1):
        path = pdir / e.filename
        if progress:
            progress(f"[{i}/{total}] {path.name}")
        preset = (e.read_preset or "").strip() or None
        disp = _display_label(e.display_name, path)
        try:
            rows.append(
                assortment.count_assortment_for_file(
                    path,
                    key_candidates=key_c,
                    name_candidates=name_c,
                    price_candidates=price_c,
                    stock_candidates=stock_c,
                    read_preset=preset,
                    source_display=disp,
                )
            )
        except Exception as ex:
            rows.append(
                {
                    "file": path.name,
                    "source": disp,
                    "assortment": None,
                    "rows_total": None,
                    "rows_non_empty_id": None,
                    "id_column": None,
                    "id_kind": None,
                    "error": str(ex),
                }
            )

    out_df = pd.DataFrame(rows)
    if "error" in out_df.columns:
        out_df["_ok"] = out_df["error"].isna()
        out_df = out_df.sort_values(
            by=["_ok", "assortment", "source"], ascending=[False, False, True]
        ).drop(columns=["_ok"])
    today = datetime.now().strftime("%d.%m.%Y")
    out = output_dir() / f"Ассортимент товаров ({today}).xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_excel(out, index=False, engine="openpyxl")
    wb = openpyxl.load_workbook(out)
    for sheet in wb.worksheets:
        assortment.format_summary_sheet(sheet)
    wb.save(out)
    return out


def run_cheapest(s: AppSettings) -> str:
    return cheapest.highlight_minimum_prices(output_dir())


def run_best_suppliers_analysis(
    s: AppSettings,
    progress: Optional[Callable[[str], None]] = None,
) -> Path:
    """
    Анализ последнего сводного прайса в папке результатов: лучшие оптовики по фиксированному
    набору препаратов и сводка «корзина» (см. best_suppliers_analysis.py).
    """
    od = output_dir()
    return best_suppliers.run_analysis_to_excel(
        inputs_dir=od,
        report_directory=od,
        progress=progress,
        emit_console_report=False,
    )
