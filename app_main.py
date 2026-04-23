# -*- coding: utf-8 -*-
"""Десктопное приложение: сводный прайс, вертикаль, ассортимент, подсветка минимума."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QThread, QUrl, Qt, Signal
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from version import __version__

from app_settings import (
    AppSettings,
    PriceFileEntry,
    default_columns,
    load_settings,
    output_dir,
    prices_dir,
    save_settings,
    unique_target_name,
)
import pipeline


PRESET_CHOICES = [
    ("Авто (по имени файла)", ""),
    ("Хожи Акбар — заголовок с 12-й строки", "khoji_akbar"),
    ("Гранд Фарм — поиск таблицы в файле", "grand_farm"),
]


def _load_app_icon() -> Optional[QIcon]:
    """Иконка окна и панели задач: assets/app_icon.ico или assets/app_icon.png."""
    roots: list[Path] = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        roots.append(Path(sys._MEIPASS))
    if getattr(sys, "frozen", False):
        roots.append(Path(sys.executable).parent)
    roots.append(Path(__file__).resolve().parent)

    seen: set[Path] = set()
    for base in roots:
        try:
            key = base.resolve()
        except OSError:
            key = base
        if key in seen:
            continue
        seen.add(key)
        for name in ("app_icon.ico", "app_icon.png"):
            p = base / "assets" / name
            if p.is_file():
                ic = QIcon(str(p))
                if not ic.isNull():
                    return ic
    return None


class JobThread(QThread):
    ok = Signal(object)
    err = Signal(str)
    log = Signal(str)

    def __init__(self, fn):
        super().__init__()
        self._fn = fn

    def run(self):
        try:
            result = self._fn(self.log.emit)
            self.ok.emit(result)
        except Exception as e:
            self.err.emit(str(e))


class SettingsDialog(QDialog):
    def __init__(self, settings: AppSettings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(720, 520)
        self.setMinimumSize(520, 360)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self.setSizeGripEnabled(True)
        tabs = QTabWidget()

        # Производители
        w_prod = QWidget()
        vl = QVBoxLayout(w_prod)
        vl.addWidget(QLabel("Подстроки в колонке «Производитель» (любая совпадает — строка берётся):"))
        self.list_producers = QListWidget()
        for p in settings.producers:
            self.list_producers.addItem(p)
        hl = QHBoxLayout()
        self.ed_producer = QLineEdit()
        self.ed_producer.setPlaceholderText("Например: ООО Фарма")
        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self._add_producer)
        btn_del = QPushButton("Удалить выбранное")
        btn_del.clicked.connect(self._del_producer)
        hl.addWidget(self.ed_producer)
        hl.addWidget(btn_add)
        vl.addWidget(self.list_producers)
        vl.addLayout(hl)
        vl.addWidget(btn_del)
        tabs.addTab(w_prod, "Производители")

        # Каталог препаратов
        w_cat = QWidget()
        v2 = QVBoxLayout(w_cat)
        v2.addWidget(
            QLabel(
                "Каноническое название в отчёте и синонимы (через точку с запятой), "
                "как они могут встречаться в прайсах."
            )
        )
        self.tbl_catalog = QTableWidget(0, 2)
        self.tbl_catalog.setHorizontalHeaderLabels(["Препарат (в сводном)", "Синонимы; через; ;"])
        self.tbl_catalog.horizontalHeader().setStretchLastSection(True)
        for item in settings.catalog:
            self._append_catalog_row(
                str(item.get("canonical", "")),
                item.get("aliases") if isinstance(item.get("aliases"), list) else [],
            )
        if self.tbl_catalog.rowCount() == 0:
            self._append_catalog_row("", [])
        btn_row = QHBoxLayout()
        b_add = QPushButton("Строка +")
        b_add.clicked.connect(lambda: self._append_catalog_row("", []))
        b_rem = QPushButton("Удалить строку")
        b_rem.clicked.connect(self._remove_catalog_row)
        btn_row.addWidget(b_add)
        btn_row.addWidget(b_rem)
        btn_row.addStretch()
        v2.addWidget(self.tbl_catalog)
        v2.addLayout(btn_row)
        tabs.addTab(w_cat, "Препараты")

        # Колонки Excel
        w_col = QWidget()
        v3 = QVBoxLayout(w_col)
        v3.addWidget(QLabel("Кандидаты имён колонок (одна строка — один вариант заголовка)."))
        self.edits: dict[str, QPlainTextEdit] = {}
        for key, title in [
            ("key_candidates", "Ключ / артикул"),
            ("name_candidates", "Наименование"),
            ("price_candidates", "Цена"),
            ("stock_candidates", "Остаток"),
            ("producer_candidates", "Производитель"),
            ("expiry_candidates", "Срок годности (резерв)"),
        ]:
            g = QGroupBox(title)
            fl = QVBoxLayout(g)
            ed = QPlainTextEdit()
            lines = settings.columns.get(key, [])
            ed.setPlainText("\n".join(str(x) for x in lines))
            fl.addWidget(ed)
            self.edits[key] = ed
            v3.addWidget(g)
        v3.addStretch()
        tabs.addTab(w_col, "Колонки Excel")

        self.chk_use_catalog = QCheckBox(
            "Ограничивать сводный прайс по препаратам (дополнительная функция)"
        )
        self.chk_use_catalog.setChecked(settings.use_catalog_matching)
        self.chk_legacy = QCheckBox(
            "Доп. правила совпадения для Анаферон/Тенотен/Дуафалак (старый режим Материа Медика)"
        )
        self.chk_legacy.setChecked(settings.legacy_priority_rules)
        outer = QVBoxLayout(self)
        outer.addWidget(tabs)
        outer.addWidget(self.chk_use_catalog)
        outer.addWidget(self.chk_legacy)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Close)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        outer.addWidget(bb)

    def _add_producer(self):
        t = self.ed_producer.text().strip()
        if t:
            self.list_producers.addItem(t)
            self.ed_producer.clear()

    def _del_producer(self):
        for it in self.list_producers.selectedItems():
            self.list_producers.takeItem(self.list_producers.row(it))

    def _append_catalog_row(self, canon: str, aliases: list):
        r = self.tbl_catalog.rowCount()
        self.tbl_catalog.insertRow(r)
        self.tbl_catalog.setItem(r, 0, QTableWidgetItem(canon))
        if isinstance(aliases, list):
            al = "; ".join(str(a) for a in aliases)
        else:
            al = str(aliases or "")
        self.tbl_catalog.setItem(r, 1, QTableWidgetItem(al))

    def _remove_catalog_row(self):
        r = self.tbl_catalog.currentRow()
        if r >= 0:
            self.tbl_catalog.removeRow(r)

    def apply_to(self, target: AppSettings) -> None:
        target.producers = [
            self.list_producers.item(i).text().strip()
            for i in range(self.list_producers.count())
            if self.list_producers.item(i).text().strip()
        ]
        catalog = []
        for r in range(self.tbl_catalog.rowCount()):
            c0 = self.tbl_catalog.item(r, 0)
            c1 = self.tbl_catalog.item(r, 1)
            canon = (c0.text() if c0 else "").strip()
            if not canon:
                continue
            raw = (c1.text() if c1 else "").strip()
            aliases = [a.strip() for a in raw.split(";") if a.strip()]
            catalog.append({"canonical": canon, "aliases": aliases})
        target.catalog = catalog
        new_cols = dict(default_columns())
        for key, ed in self.edits.items():
            lines = [ln.strip() for ln in ed.toPlainText().splitlines() if ln.strip()]
            if lines:
                new_cols[key] = lines
        target.columns = new_cols
        target.use_catalog_matching = self.chk_use_catalog.isChecked()
        target.legacy_priority_rules = self.chk_legacy.isChecked()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Price Aggregator — {__version__}")
        self.resize(960, 540)
        self.settings = load_settings()

        cw = QWidget()
        self.setCentralWidget(cw)
        layout = QVBoxLayout(cw)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(500)
        self.log.setFixedHeight(120)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Файл", "Название в сводном", "Режим чтения"])
        hdr = self.table.horizontalHeader()
        hdr.setStretchLastSection(False)
        hdr.setMinimumSectionSize(60)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 300)
        self.table.setColumnWidth(2, 220)
        layout.addWidget(QLabel("Прайсы (файлы в рабочей папке программы):"))
        layout.addWidget(self.table)

        row = QHBoxLayout()
        self.btn_load = QPushButton("Загрузить прайсы…")
        self.btn_load.clicked.connect(self.on_load_prices)
        self.btn_remove = QPushButton("Удалить из списка")
        self.btn_remove.clicked.connect(self.on_remove_selected)
        self.btn_settings = QPushButton("Настройки…")
        self.btn_settings.clicked.connect(self.on_settings)
        self.btn_out = QPushButton("Открыть папку результатов")
        self.btn_out.clicked.connect(self.on_open_output)
        row.addWidget(self.btn_load)
        row.addWidget(self.btn_remove)
        row.addWidget(self.btn_settings)
        row.addWidget(self.btn_out)
        row.addStretch()
        layout.addLayout(row)

        row_mode = QHBoxLayout()
        row_mode.addWidget(QLabel("Режим сводного:"))
        self.cmb_match_mode = QComboBox()
        self.cmb_match_mode.addItem("Все по производителю", False)
        self.cmb_match_mode.addItem("По препаратам (доп.)", True)
        idx_mode = self.cmb_match_mode.findData(bool(self.settings.use_catalog_matching))
        self.cmb_match_mode.setCurrentIndex(idx_mode if idx_mode >= 0 else 0)
        self.cmb_match_mode.currentIndexChanged.connect(self.on_match_mode_changed)
        row_mode.addWidget(self.cmb_match_mode)
        row_mode.addStretch()
        layout.addLayout(row_mode)

        self.btn_svodny = QPushButton("Сводный прайс")
        self.btn_vertical = QPushButton("Сводный (вертикально)")
        self.btn_assort = QPushButton("Ассортимент")
        self.btn_cheap = QPushButton("Подсветить мин. цены")
        self.btn_best_sup = QPushButton("Анализ оптовиков")
        self.btn_best_sup.setToolTip(
            "Отчёт по последнему сводному прайсу в папке результатов (лучшие оптовики и корзина)."
        )
        for b in (
            self.btn_svodny,
            self.btn_vertical,
            self.btn_assort,
            self.btn_cheap,
            self.btn_best_sup,
        ):
            b.clicked.connect(self._make_runner(b))
        row2 = QHBoxLayout()
        row2.addWidget(self.btn_svodny)
        row2.addWidget(self.btn_vertical)
        row2.addWidget(self.btn_assort)
        row2.addStretch()
        layout.addLayout(row2)
        row2b = QHBoxLayout()
        row2b.addWidget(self.btn_cheap)
        row2b.addWidget(self.btn_best_sup)
        row2b.addStretch()
        layout.addLayout(row2b)
        layout.addWidget(QLabel("Журнал:"))
        layout.addWidget(self.log)

        self._worker: Optional[JobThread] = None
        self._settings_dialog: Optional[SettingsDialog] = None
        self.rebuild_table()

    def _make_runner(self, btn: QPushButton):
        def go():
            if btn is self.btn_svodny:
                self.run_job(lambda lg: pipeline.run_svodny_horizontal(self.settings, lg), "Сводный")
            elif btn is self.btn_vertical:
                self.run_job(lambda lg: pipeline.run_svodny_vertical(self.settings, lg), "Вертикально")
            elif btn is self.btn_assort:
                self.run_job(lambda lg: pipeline.run_assortment(self.settings, lg), "Ассортимент")
            elif btn is self.btn_cheap:
                self.run_job(lambda lg: pipeline.run_cheapest(self.settings), "Подсветка")
            elif btn is self.btn_best_sup:
                self.run_job(
                    lambda lg: pipeline.run_best_suppliers_analysis(self.settings, lg),
                    "Анализ оптовиков",
                )

        return go

    def rebuild_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        for i, e in enumerate(self.settings.price_files):
            p = prices_dir() / e.filename
            if not p.is_file():
                continue
            self.table.insertRow(self.table.rowCount())
            r = self.table.rowCount() - 1
            it = QTableWidgetItem(e.filename)
            it.setFlags(it.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(r, 0, it)
            name_item = QTableWidgetItem(e.display_name)
            name_item.setToolTip(e.display_name or p.name)
            self.table.setItem(r, 1, name_item)
            cb = QComboBox()
            for label, data in PRESET_CHOICES:
                cb.addItem(label, data)
            idx = cb.findData(e.read_preset or "")
            cb.setCurrentIndex(idx if idx >= 0 else 0)
            self.table.setCellWidget(r, 2, cb)
        self.table.blockSignals(False)

    def _persist_table_to_settings(self):
        new_list: list[PriceFileEntry] = []
        pdir = prices_dir()
        for r in range(self.table.rowCount()):
            fn_item = self.table.item(r, 0)
            if not fn_item:
                continue
            fn = fn_item.text().strip()
            if not (pdir / fn).is_file():
                continue
            dn = ""
            if self.table.item(r, 1):
                dn = self.table.item(r, 1).text()
            w = self.table.cellWidget(r, 2)
            preset = ""
            if isinstance(w, QComboBox):
                preset = w.currentData() or ""
            new_list.append(PriceFileEntry(filename=fn, display_name=dn, read_preset=preset or ""))
        self.settings.price_files = new_list

    def on_load_prices(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите прайсы",
            "",
            "Excel (*.xlsx *.xls);;Все файлы (*.*)",
        )
        if not paths:
            return
        dest = prices_dir()
        self._persist_table_to_settings()
        existing_names = {e.filename for e in self.settings.price_files}
        for src in paths:
            srcp = Path(src)
            target = unique_target_name(dest, srcp.name)
            shutil.copy2(srcp, target)
            name = target.name
            if name not in existing_names:
                self.settings.price_files.append(PriceFileEntry(filename=name, display_name="", read_preset=""))
                existing_names.add(name)
        save_settings(self.settings)
        self.rebuild_table()
        self.log.appendPlainText(f"Скопировано файлов: {len(paths)}")

    def on_remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "", "Выделите строки в таблице.")
            return
        self._persist_table_to_settings()
        to_remove = set()
        for r in rows:
            it = self.table.item(r, 0)
            if it:
                to_remove.add(it.text())
        self.settings.price_files = [e for e in self.settings.price_files if e.filename not in to_remove]
        save_settings(self.settings)
        self.rebuild_table()

    def on_settings(self):
        if self._settings_dialog and self._settings_dialog.isVisible():
            self._settings_dialog.raise_()
            self._settings_dialog.activateWindow()
            return

        dlg = SettingsDialog(self.settings, self)
        dlg.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)

        def _save_settings():
            dlg.apply_to(self.settings)
            save_settings(self.settings)
            idx_mode = self.cmb_match_mode.findData(bool(self.settings.use_catalog_matching))
            self.cmb_match_mode.blockSignals(True)
            self.cmb_match_mode.setCurrentIndex(idx_mode if idx_mode >= 0 else 0)
            self.cmb_match_mode.blockSignals(False)
            self.log.appendPlainText("Настройки сохранены.")

        def _on_close(_obj=None):
            self._settings_dialog = None

        dlg.accepted.connect(_save_settings)
        dlg.destroyed.connect(_on_close)
        self._settings_dialog = dlg
        dlg.show()

    def on_open_output(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(output_dir())))

    def on_match_mode_changed(self):
        value = bool(self.cmb_match_mode.currentData())
        self.settings.use_catalog_matching = value
        save_settings(self.settings)
        mode_txt = "По препаратам" if value else "Все по производителю"
        self.log.appendPlainText(f"Режим сводного: {mode_txt}")

    def _set_busy(self, busy: bool):
        for b in (
            self.btn_load,
            self.btn_remove,
            self.btn_settings,
            self.btn_svodny,
            self.btn_vertical,
            self.btn_assort,
            self.btn_cheap,
            self.btn_best_sup,
        ):
            b.setEnabled(not busy)
        self.cmb_match_mode.setEnabled(not busy)

    def run_job(self, fn, title: str):
        self._persist_table_to_settings()
        save_settings(self.settings)
        if self._worker and self._worker.isRunning():
            QMessageBox.warning(self, "", "Дождитесь завершения текущей операции.")
            return

        self._set_busy(True)
        self.log.appendPlainText(f"— {title}: старт —")
        self._worker = JobThread(fn)
        self._worker.log.connect(self.log.appendPlainText)

        def ok(res):
            self._set_busy(False)
            if isinstance(res, Path):
                self.log.appendPlainText(f"Готово: {res}")
            else:
                self.log.appendPlainText(f"Готово: {res}")

        def err(msg):
            self._set_busy(False)
            self.log.appendPlainText(f"Ошибка: {msg}")
            QMessageBox.critical(self, "Ошибка", msg)

        self._worker.ok.connect(ok)
        self._worker.err.connect(err)
        self._worker.finished.connect(self._worker.deleteLater)
        self._worker.start()

    def closeEvent(self, event):
        self._persist_table_to_settings()
        save_settings(self.settings)
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    icon = _load_app_icon()
    if icon is not None:
        app.setWindowIcon(icon)
    w = MainWindow()
    if icon is not None:
        w.setWindowIcon(icon)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
