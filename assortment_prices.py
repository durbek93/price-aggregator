# -*- coding: utf-8 -*-
"""
Подсчёт ассортимента (кол-ва уникальных товаров) по каждому прайсу в папке prices/.

Логика чтения/поиска заголовков и выбора колонок переиспользуется из svodny_price.py.
Ассортимент считается как кол-во уникальных значений в колонке идентификатора:
- если найдена колонка из key_candidates (Артикул/Код/SKU/ID/...), используем её;
- иначе используем колонку наименования (name_candidates).

Результат сохраняется в output/Ассортимент товаров (ДД.ММ.ГГГГ).xlsx
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
import pandas as pd
import yaml

from svodny_price import (
    DEFAULT_KEY_CANDIDATES,
    DEFAULT_NAME_CANDIDATES,
    DEFAULT_PRICE_CANDIDATES,
    DEFAULT_STOCK_CANDIDATES,
    friendly_source_label,
    format_summary_sheet,
    normalize_name,
    pick_column,
    read_price_table,
)


def _clean_id_value(v: Any, *, normalize: bool) -> Optional[str]:
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).strip()
    if not s:
        return None
    return normalize_name(s) if normalize else s


def count_assortment_for_file(
    path: Path,
    *,
    key_candidates: List[str],
    name_candidates: List[str],
    price_candidates: List[str],
    stock_candidates: List[str],
    read_preset: Optional[str] = None,
    source_display: Optional[str] = None,
) -> Dict[str, Any]:
    df = read_price_table(
        path,
        key_candidates=key_candidates,
        name_candidates=name_candidates,
        price_candidates=price_candidates,
        stock_candidates=stock_candidates,
        use_key=True,
        use_name=True,
        read_preset=read_preset,
    )

    key_col = pick_column(list(df.columns), key_candidates)
    name_col = pick_column(list(df.columns), name_candidates)

    used_col = key_col or name_col
    if not used_col:
        raise ValueError("Не удалось определить колонку идентификатора (key/name).")

    normalize = used_col == name_col  # для наименований нормализуем, чтобы схлопнуть регистр/пробелы
    values = [
        _clean_id_value(v, normalize=normalize)
        for v in df.get(used_col, pd.Series([], dtype=object)).tolist()
    ]
    non_empty = [v for v in values if v is not None]
    unique_count = len(set(non_empty))

    src_label = (source_display or "").strip() or friendly_source_label(path.stem)
    return {
        "file": path.name,
        "source": src_label,
        "assortment": unique_count,
        "rows_total": int(len(df)),
        "rows_non_empty_id": int(len(non_empty)),
        "id_column": used_col,
        "id_kind": "key" if key_col else "name",
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Подсчёт ассортимента товаров по каждому прайсу в папке prices"
    )
    parser.add_argument("--config", type=Path, default=None, help="Путь к config.yml")
    parser.add_argument(
        "--inputs", type=str, default="prices", help="Папка или файл(ы) с прайсами"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Выходной Excel (по умолчанию: output/Ассортимент товаров (ДД.ММ.ГГГГ).xlsx)",
    )
    args = parser.parse_args()

    if args.output is None:
        today = datetime.now().strftime("%d.%m.%Y")
        args.output = Path("output") / f"Ассортимент товаров ({today}).xlsx"

    key_candidates = DEFAULT_KEY_CANDIDATES
    name_candidates = DEFAULT_NAME_CANDIDATES
    price_candidates = DEFAULT_PRICE_CANDIDATES
    stock_candidates = DEFAULT_STOCK_CANDIDATES

    if args.config and args.config.exists():
        with open(args.config, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        cols = cfg.get("columns", {})
        key_candidates = cols.get("key_candidates", key_candidates)
        name_candidates = cols.get("name_candidates", name_candidates)
        price_candidates = cols.get("price_candidates", price_candidates)
        stock_candidates = cols.get("stock_candidates", stock_candidates)

    inputs_path = Path(args.inputs)
    if inputs_path.is_dir():
        files = (
            list(inputs_path.glob("*.xlsx"))
            + list(inputs_path.glob("*.xls"))
            + list(inputs_path.glob("*.csv"))
        )
    else:
        files = [inputs_path]

    if not files:
        print("Нет файлов для обработки.", file=sys.stderr)
        sys.exit(1)

    rows: List[Dict[str, Any]] = []
    total = len(files)
    for i, path in enumerate(files, 1):
        try:
            print(f"[{i}/{total}] {path.name}", flush=True)
            if path.suffix.lower() == ".csv":
                # На случай если появятся CSV — считаем ассортимент по первой подходящей колонке.
                df = pd.read_csv(path, sep=None, engine="python")
                df.columns = [str(c).strip() for c in df.columns]
                key_col = pick_column(list(df.columns), key_candidates)
                name_col = pick_column(list(df.columns), name_candidates)
                used_col = key_col or name_col
                if not used_col:
                    raise ValueError("Не удалось определить колонку идентификатора (key/name) в CSV.")
                normalize = used_col == name_col
                vals = [_clean_id_value(v, normalize=normalize) for v in df[used_col].tolist()]
                non_empty = [v for v in vals if v is not None]
                rows.append(
                    {
                        "file": path.name,
                        "source": friendly_source_label(path.stem),
                        "assortment": len(set(non_empty)),
                        "rows_total": int(len(df)),
                        "rows_non_empty_id": int(len(non_empty)),
                        "id_column": used_col,
                        "id_kind": "key" if key_col else "name",
                    }
                )
            else:
                rows.append(
                    count_assortment_for_file(
                        path,
                        key_candidates=key_candidates,
                        name_candidates=name_candidates,
                        price_candidates=price_candidates,
                        stock_candidates=stock_candidates,
                    )
                )
        except Exception as e:
            rows.append(
                {
                    "file": path.name,
                    "source": friendly_source_label(path.stem),
                    "assortment": None,
                    "rows_total": None,
                    "rows_non_empty_id": None,
                    "id_column": None,
                    "id_kind": None,
                    "error": str(e),
                }
            )
            print(f"Ошибка при обработке {path.name}: {e}", file=sys.stderr)

    out_df = pd.DataFrame(rows)
    # Сортировка: сначала успешные, затем с ошибками; по убыванию ассортимента.
    out_df["_ok"] = out_df["error"].isna() if "error" in out_df.columns else True
    out_df = out_df.sort_values(by=["_ok", "assortment", "source"], ascending=[False, False, True]).drop(
        columns=["_ok"]
    )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_excel(args.output, index=False, engine="openpyxl")

    wb = openpyxl.load_workbook(args.output)
    for sheet in wb.worksheets:
        format_summary_sheet(sheet)
    wb.save(args.output)

    print(f"Сохранено: {args.output}")


if __name__ == "__main__":
    main()

