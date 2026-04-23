"""Копия запуска сводного прайса с транспонированной матрицей."""
from __future__ import annotations

import sys
import time
from datetime import datetime

import pandas as pd
import svodny_price as base

_BASE_FORMATTER = base.format_summary_sheet


def build_matrix_from_extracted(
    extracted,
    catalog,
    all_source_labels=None,
) -> pd.DataFrame:
    """Строит транспонированную матрицу: строки — оптовики, столбцы — препараты."""
    canon_order = [c[0] for c in catalog]

    # canon -> source -> min(price)
    by_canon = {c: {} for c in canon_order}
    sources_seen = set()
    for row in extracted:
        canon = row.get("canonical") or row.get("name")
        if isinstance(canon, str) and canon not in by_canon:
            by_canon[canon] = {}
        if not isinstance(canon, str):
            continue

        src = row.get("source_label") or base.friendly_source_label(row["source"])
        sources_seen.add(src)
        price = row.get("price")
        prev_price = by_canon[canon].get(src)
        if prev_price is None:
            by_canon[canon][src] = price
        elif price is not None:
            by_canon[canon][src] = min(prev_price, price)

    dynamic_canons = [c for c in by_canon.keys() if c not in canon_order]
    dynamic_canons.sort(key=lambda x: str(x).casefold())
    final_canons = canon_order + dynamic_canons

    sources = list(all_source_labels) if all_source_labels else sorted(sources_seen)
    cols = ["Оптовик"] + final_canons
    data = []
    for src in sources:
        row = [src] + [by_canon.get(canon, {}).get(src) for canon in final_canons]
        data.append(row)

    return pd.DataFrame(data, columns=cols)


def format_summary_sheet(sheet):
    """Форматирование для транспонированного листа."""
    _BASE_FORMATTER(sheet)
    # Для новой структуры: оптовики в A-колонке, препараты в шапке B..N.
    sheet.row_dimensions[1].height = 38


def main() -> None:
    # Если пользователь не задал --output, используем отдельное имя файла
    # чтобы результат отличался от стандартного скрипта.
    if "--output" not in sys.argv:
        today = datetime.now().strftime("%d.%m.%Y")
        sys.argv.extend(["--output", f"output/Сводный прайс вертикально {today}.xlsx"])

    # Подменяем построение матрицы и форматирование только для запуска из этого файла.
    base.build_matrix_from_extracted = build_matrix_from_extracted
    base.format_summary_sheet = format_summary_sheet
    base.main()


if __name__ == "__main__":
    _timer_prefix = "Сводный прайс (вертикально)"
    _t0 = time.perf_counter()
    _stop_live, _live_thread = base.start_live_console_timer(_t0, title_prefix=_timer_prefix)
    try:
        main()
    finally:
        _stop_live.set()
        _live_thread.join(timeout=2.0)
        _total = time.perf_counter() - _t0
        base.finalize_live_console_timer_title(_timer_prefix, _total)
        print(base.format_elapsed_seconds(_total), flush=True)
