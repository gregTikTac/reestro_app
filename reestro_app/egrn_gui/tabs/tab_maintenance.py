# -*- coding: utf-8 -*-
"""Вкладка «Обслуживание»: очистка дублей PDF/строк и пересборка из кэша."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QLabel,
    QFileDialog, QCheckBox, QGroupBox, QFormLayout, QPlainTextEdit, QMessageBox,
)

from ..engine_bridge import engine, _engine_dir
from ..paths import resolve_output_root, describe_output_folder, count_cache_json


class _MaintWorker(QThread):
    line = Signal(str)
    done = Signal(str)

    def __init__(self, op: str, out_dir: Path, *, apply: bool,
                 use_trash: bool, no_pdf: bool, parent=None):
        super().__init__(parent)
        self.op = op
        self.out_dir = Path(out_dir)
        self.apply = apply
        self.use_trash = use_trash
        self.no_pdf = no_pdf

    def run(self):
        try:
            import cleanup_output as cu  # из каталога движка
        except Exception as exc:  # noqa: BLE001
            self.done.emit(f"Не удалось загрузить модуль очистки: {exc}")
            return
        eng = engine()
        out = self.out_dir
        report = out / "report.xlsx"
        pdf_dir = out / "pdf"
        pdf_dir.mkdir(parents=True, exist_ok=True)

        try:
            if self.op == "rebuild":
                self._rebuild(cu, eng, out, report, pdf_dir)
            else:
                self._dedupe(cu, eng, out, report, pdf_dir)
        except Exception as exc:  # noqa: BLE001
            import traceback
            self.done.emit(f"Ошибка: {exc}\n{traceback.format_exc()}")

    def _dedupe(self, cu, eng, out, report, pdf_dir):
        if not report.exists():
            self.done.emit(f"Не найден {report}. Используйте «Пересобрать из кэша».")
            return
        _, existing = eng.load_existing_report(report, pdf_dir)
        rows, stats = cu.dedupe_report_rows(existing, pdf_dir)
        self.line.emit(
            f"report.xlsx: строк {stats['input_rows']}, уникальных КН "
            f"{stats['unique_kn']}, КН с дублями {stats['duplicate_kn']}, "
            f"удалить строк {stats['removed_rows']}, останется "
            f"{stats['output_rows']}")
        keep = cu.referenced_pdfs(rows)
        trash = (out / "pdf_old") if self.use_trash else None
        pstats = cu.remove_orphan_pdfs(pdf_dir, keep, apply=self.apply,
                                       trash_dir=trash)
        self.line.emit(f"PDF: оставить {pstats['pdf_kept']}, лишних "
                       f"{pstats['pdf_removed']}")
        if not self.apply:
            self.done.emit("Пробный прогон завершён. Ничего не изменено.")
            return
        uniq = len({eng.normalize_kn(r[1]) for r in rows if eng.normalize_kn(r[1])})
        eng.write_report_xlsx(rows, {
            "total": uniq, "ok": uniq, "failed": 0,
            "skipped_no_kn": sum(1 for r in rows if not eng.normalize_kn(r[1])),
            "skipped_already": 0,
        }, report)
        self.done.emit(
            f"Готово. В отчёте {len(rows)} строк, {uniq} КН. "
            f"PDF удалено/перемещено: {pstats['pdf_removed']}.")

    def _rebuild(self, cu, eng, out, report, pdf_dir):
        cache = out / "cache" / "json"
        if not cache.exists() or not any(cache.glob("*.json")):
            self.done.emit(f"Нет JSON-кэша: {cache}")
            return
        existing = []
        if report.exists():
            _, existing = eng.load_existing_report(report, pdf_dir)
        tz = _engine_dir() / "TZ" / "Запрос.xlsx"
        inputs = [tz] if tz.exists() else []
        n_json = len(list(cache.glob("*.json")))
        self.line.emit(f"Пересборка из кэша ({n_json} JSON)…")
        rows = cu.rebuild_report_from_cache(
            out, inputs, None, existing, regen_pdf=not self.no_pdf and self.apply)
        uniq = len({eng.normalize_kn(r[1]) for r in rows if eng.normalize_kn(r[1])})
        self.line.emit(f"Получено: {uniq} уникальных КН, {len(rows)} строк")
        if not self.apply:
            self.done.emit("Пробный прогон: отчёт НЕ записан (включите «Применить»).")
            return
        eng.write_report_xlsx(rows, {
            "total": uniq, "ok": uniq, "failed": 0,
            "skipped_no_kn": 0, "skipped_already": 0,
        }, report)
        keep = cu.referenced_pdfs(rows)
        trash = (out / "pdf_old") if self.use_trash else None
        pstats = cu.remove_orphan_pdfs(pdf_dir, keep, apply=True, trash_dir=trash)
        self.done.emit(f"Готово. {len(rows)} строк, {uniq} КН. Лишних PDF убрано: "
                       f"{pstats['pdf_removed']}.")


class MaintenanceTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.worker: _MaintWorker | None = None

        root = QVBoxLayout(self)

        box = QGroupBox("Папка результатов")
        form = QFormLayout(box)
        row = QHBoxLayout()
        self.output_edit = QLineEdit(win.settings.get("last_output", ""))
        self.output_edit.setPlaceholderText(
            r"напр. D:\parser\reestro\output  — НЕ pdf\ и НЕ cache\json\!")
        self.output_edit.textChanged.connect(self._refresh_folder_info)
        btn = QPushButton("Обзор…")
        btn.clicked.connect(self.pick_output)
        row.addWidget(self.output_edit)
        row.addWidget(btn)
        w = QWidget()
        w.setLayout(row)
        form.addRow("Папка результатов:", w)
        self.folder_info = QLabel("")
        self.folder_info.setWordWrap(True)
        self.folder_info.setStyleSheet("color: #555; font-size: 11px;")
        form.addRow("", self.folder_info)
        root.addWidget(box)

        hint = QLabel(
            "Выбирайте корень output (там же лежат папки pdf\\ и cache\\). "
            "Если выбрали pdf\\ или cache\\json\\ — программа подставит корень сама.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #827717;")
        root.addWidget(hint)

        opt = QGroupBox("Параметры")
        oform = QVBoxLayout(opt)
        self.use_trash = QCheckBox("Лишние PDF переносить в pdf_old (не удалять)")
        self.use_trash.setChecked(True)
        self.no_pdf = QCheckBox("При пересборке не пересоздавать PDF")
        self.no_pdf.setChecked(True)
        oform.addWidget(self.use_trash)
        oform.addWidget(self.no_pdf)
        root.addWidget(opt)

        ctl = QHBoxLayout()
        self.dry_btn = QPushButton("Анализ дублей (пробно)")
        self.dry_btn.clicked.connect(lambda: self.run_op("dedupe", apply=False))
        self.apply_btn = QPushButton("Очистить дубли (применить)")
        self.apply_btn.clicked.connect(lambda: self.run_op("dedupe", apply=True))
        self.rebuild_btn = QPushButton("Пересобрать из кэша")
        self.rebuild_btn.clicked.connect(lambda: self.run_op("rebuild", apply=True))
        ctl.addWidget(self.dry_btn)
        ctl.addWidget(self.apply_btn)
        ctl.addWidget(self.rebuild_btn)
        ctl.addStretch(1)
        root.addLayout(ctl)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        root.addWidget(self.log, 1)

        if self.output_edit.text().strip():
            self._refresh_folder_info()

    def _refresh_folder_info(self):
        p = self.output_edit.text().strip()
        if not p:
            self.folder_info.setText("")
            return
        if not Path(p).exists():
            self.folder_info.setText("Папка не существует.")
            self.folder_info.setStyleSheet("color: #c62828; font-size: 11px;")
            return
        self.folder_info.setText(describe_output_folder(p))
        n = count_cache_json(resolve_output_root(p))
        if n:
            self.folder_info.setStyleSheet("color: #2e7d32; font-size: 11px;")
        else:
            self.folder_info.setStyleSheet("color: #c62828; font-size: 11px;")

    def pick_output(self):
        path = QFileDialog.getExistingDirectory(
            self, "Папка результатов (корень output, не pdf)", self.output_edit.text())
        if path:
            root = resolve_output_root(path)
            self.output_edit.setText(str(root))
            self._refresh_folder_info()
            if Path(path).resolve() != root:
                self.log.appendPlainText(
                    f"Подсказка: выбрана подпапка {path}\n"
                    f"Используется корень output: {root}")

    def _output_root(self) -> Path | None:
        raw = self.output_edit.text().strip()
        if not raw:
            return None
        root = resolve_output_root(raw)
        if root != Path(raw).resolve():
            self.output_edit.setText(str(root))
            self._refresh_folder_info()
        return root

    def run_op(self, op: str, *, apply: bool):
        out_path = self._output_root()
        if not out_path or not out_path.exists():
            QMessageBox.warning(self, "Обслуживание", "Выберите папку результатов.")
            return
        n_json = count_cache_json(out_path)
        if op == "rebuild" and n_json == 0:
            cache = out_path / "cache" / "json"
            QMessageBox.warning(
                self, "Нет кэша",
                f"В папке нет JSON-кэша:\n{cache}\n\n"
                f"Выберите корень output (где лежат pdf\\ и cache\\), "
                f"не подпапку pdf\\.\n\n"
                f"Сейчас указано: {out_path}")
            return
        if apply and op == "dedupe":
            if QMessageBox.question(
                    self, "Подтверждение",
                    "Очистить дубли и убрать лишние PDF?") != QMessageBox.Yes:
                return
        self._set_busy(True)
        self.log.clear()
        self.worker = _MaintWorker(op, out_path, apply=apply,
                                   use_trash=self.use_trash.isChecked(),
                                   no_pdf=self.no_pdf.isChecked(), parent=self)
        self.worker.line.connect(lambda s: self.log.appendPlainText(s))
        self.worker.done.connect(self._on_done)
        self.worker.start()

    def _on_done(self, msg: str):
        self.log.appendPlainText(msg)
        self._set_busy(False)
        self.worker = None
        if self.output_edit.text().strip():
            self.win.logs_tab.refresh_files(resolve_output_root(self.output_edit.text()))

    def _set_busy(self, busy: bool):
        for b in (self.dry_btn, self.apply_btn, self.rebuild_btn):
            b.setEnabled(not busy)
