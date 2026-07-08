# -*- coding: utf-8 -*-
"""
Вкладка «Собственники (Росреестр)» — автоматический сбор формы собственности и
обременений с lk.rosreestr.ru и применение к report.xlsx и PDF.

Сбор полностью автоматический. Единственное действие оператора — РАЗОВЫЙ вход
через Госуслуги в открывшемся браузере (код подтверждения 2FA нельзя
автоматизировать). Сессия сохраняется, при следующих запусках вход не нужен.
"""
from __future__ import annotations

import re
import sys
import threading
from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit, QPushButton,
    QFileDialog, QGroupBox, QLabel, QPlainTextEdit, QMessageBox, QCheckBox,
    QProgressBar,
)

from ..logbus import RunLogger
from ..paths import resolve_output_root, count_cache_json


def _count_rosreestr_cache(out_root: Path) -> int:
    d = Path(out_root) / "cache" / "rosreestr"
    if not d.is_dir():
        return 0
    return sum(1 for _ in d.glob("*.json"))


class _CollectWorker(QThread):
    line = Signal(str)
    progress = Signal(int, int)
    need_login = Signal()
    form_ready = Signal()
    need_assist = Signal(str)
    result = Signal(object)

    def __init__(self, out_root: Path, only_with_rights: bool, redo: bool,
                 *, assisted: bool = False, apply_after: bool = False,
                 parent=None):
        super().__init__(parent)
        self.out_root = out_root
        self.only_with_rights = only_with_rights
        self.redo = redo
        self.assisted = assisted
        self.apply_after = apply_after
        self._login_event = threading.Event()
        self._assist_event = threading.Event()
        self._login_ok = True
        self._cancel = False
        self._skip = False

    def confirm_login(self):
        self._login_ok = True
        self._login_event.set()

    def confirm_assist(self):
        self._assist_event.set()

    def cancel(self):
        self._cancel = True
        self._login_event.set()
        self._assist_event.set()

    def skip_current(self):
        self._skip = True

    def is_login_done(self) -> bool:
        return self._login_event.is_set() and self._login_ok

    def should_skip(self) -> bool:
        return self._skip

    def consume_skip(self):
        self._skip = False

    def wait_assist(self, kn: str) -> bool:
        self._assist_event.clear()
        self.need_assist.emit(kn)
        while not self._assist_event.is_set():
            if self._cancel:
                return False
            try:
                from ..rosreestr_collector import _import_engine_modules
                _, fr = _import_engine_modules()
                fr._poll_assist_events(kn)
            except Exception:
                pass
            self._assist_event.wait(timeout=0.35)
        return not self._cancel

    def run(self):
        from ..rosreestr_collector import run_collection, CollectResult

        logger: RunLogger | None = None
        try:
            logger = RunLogger(self.out_root, prefix="rosreestr")
        except Exception:  # noqa: BLE001
            logger = None

        def _kn_from_msg(msg: str) -> str:
            m = re.match(r"\[(\d+)/(\d+)\]\s+([\d:]+)", msg.strip())
            return m.group(3) if m else ""

        def _level_from_msg(msg: str) -> str:
            low = msg.lower()
            if any(x in low for x in ("ошибка", "капча", "не распознана", "пропуск",
                                      "не найдена", "не появилась", "таймаут")):
                return "WARN"
            if "сохранено" in low:
                return "INFO"
            return "INFO"

        def on_log(msg: str) -> None:
            self.line.emit(msg)
            if not logger:
                return
            kn = _kn_from_msg(msg)
            logger.write_line(msg, level=_level_from_msg(msg), kn=kn)
            if kn and "сохранено:" in msg.lower():
                logger.write_event({
                    "type": "object",
                    "status": "ok",
                    "kn": kn,
                    "message": msg.strip(),
                })
            elif kn and any(x in msg.lower() for x in
                            ("не распознана", "ошибка:", "пропуск")):
                logger.write_event({
                    "type": "object",
                    "status": "failed",
                    "kn": kn,
                    "message": msg.strip(),
                })

        def on_need_login():
            self._login_event.clear()
            self._login_ok = True
            self.need_login.emit()

        def on_form_ready():
            self.form_ready.emit()

        if logger:
            logger.write_event({
                "type": "start",
                "message": "Сбор формы собственности и обременений (Росreestr)",
            })

        try:
            res = run_collection(
                self.out_root,
                only_with_rights=self.only_with_rights,
                redo=self.redo,
                assisted=self.assisted,
                apply_after_each=self.apply_after,
                on_log=on_log,
                on_progress=self.progress.emit,
                on_need_login=on_need_login,
                is_login_done=self.is_login_done,
                on_form_ready=on_form_ready,
                should_cancel=lambda: self._cancel,
                should_skip=self.should_skip,
                consume_skip=self.consume_skip,
                wait_assist=self.wait_assist if self.assisted else None,
            )
        except Exception as exc:  # noqa: BLE001
            res = CollectResult(error=str(exc))
            on_log(f"Ошибка: {exc}")
        finally:
            if logger:
                summary = (
                    f"Готово: сохранено {res.saved} из {res.total}, "
                    f"ошибок {res.failed}"
                    + (" (остановлено)" if res.cancelled else "")
                )
                if res.error:
                    summary = res.error
                logger.write_event({"type": "done", "message": summary})
                on_log(f"Лог записан: {logger.path}")
                logger.close()

        self.result.emit(res)


class OwnersTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self.worker: _CollectWorker | None = None
        root = QVBoxLayout(self)

        intro = QLabel(
            "Форма собственности и обременения берутся с lk.rosreestr.ru "
            "(вход кадастровым инженером через Госуслуги). API Контура их не отдаёт.\n"
            "Сбор автоматический. Один раз нужно войти через Госуслуги в браузере "
            "(код подтверждения 2FA вводится вручную) — дальше без участия оператора.")
        intro.setWordWrap(True)
        intro.setStyleSheet("color: #333;")
        root.addWidget(intro)

        box = QGroupBox("Папка результатов")
        form = QFormLayout(box)
        row = QHBoxLayout()
        self.output_edit = QLineEdit(win.settings.get("last_output", ""))
        self.output_edit.setPlaceholderText(r"корень output (где report.xlsx, cache\, pdf\)")
        self.output_edit.textChanged.connect(self._refresh_info)
        btn = QPushButton("Обзор…")
        btn.clicked.connect(self.pick_output)
        row.addWidget(self.output_edit)
        row.addWidget(btn)
        w = QWidget(); w.setLayout(row)
        form.addRow("Папка:", w)

        self.only_rights = QCheckBox("Только объекты с зарегистрированными правами")
        self.only_rights.setChecked(True)
        self.redo = QCheckBox("Пересобрать уже собранные")
        opts = QHBoxLayout()
        opts.addWidget(self.only_rights)
        opts.addWidget(self.redo)
        opts.addStretch(1)
        ow = QWidget(); ow.setLayout(opts)
        form.addRow("", ow)

        self.apply_after = QCheckBox(
            "После сбора сразу применить к report.xlsx и PDF")
        self.apply_after.setChecked(True)
        form.addRow("", self.apply_after)

        self.assisted_mode = QCheckBox(
            "Проверочный режим (с оператором в браузере)")
        self.assisted_mode.setToolTip(
            "Программа сама подставляет КН в поле поиска. Вы выбираете "
            "«Вид объекта», нажимаете «Найти» и открываете справку — затем "
            "«Данные готовы — записать и далее».")
        form.addRow("", self.assisted_mode)

        self.info = QLabel("")
        self.info.setWordWrap(True)
        self.info.setStyleSheet("color: #555; font-size: 11px;")
        form.addRow("", self.info)
        root.addWidget(box)

        # Управление сбором
        gbox = QGroupBox("Автоматический сбор")
        gl = QVBoxLayout(gbox)
        hint = QLabel(
            "Откроется браузер. Если потребуется — войдите через Госуслуги "
            "(один раз) и нажмите «Вход выполнен — продолжить». Дальше всё "
            "соберётся само и (по галочке выше) попадёт в Excel и PDF.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #827717; font-size: 11px;")
        gl.addWidget(hint)

        gbtns = QHBoxLayout()
        self.collect_btn = QPushButton("Запустить автоматический сбор")
        self.collect_btn.clicked.connect(self.start_collect)
        self.continue_btn = QPushButton("Вход выполнен — продолжить")
        self.continue_btn.clicked.connect(self.on_continue)
        self.continue_btn.setEnabled(False)
        self.cancel_btn = QPushButton("Остановить")
        self.cancel_btn.clicked.connect(self.on_cancel)
        self.cancel_btn.setEnabled(False)
        self.skip_btn = QPushButton("Пропустить объект")
        self.skip_btn.clicked.connect(self.on_skip)
        self.skip_btn.setEnabled(False)
        self.skip_btn.setToolTip(
            "Перейти к следующему КН, если текущий не находится или завис.")
        self.assist_btn = QPushButton("Данные готовы — записать и далее")
        self.assist_btn.clicked.connect(self.on_assist_confirm)
        self.assist_btn.setEnabled(False)
        self.assist_btn.setToolTip(
            "В проверочном режиме: после открытия справки нажмите, "
            "чтобы записать данные и перейти к следующему объекту.")
        self.reset_profile_btn = QPushButton("Сбросить профиль браузера")
        self.reset_profile_btn.clicked.connect(self.reset_browser_profile)
        self.reset_profile_btn.setToolTip(
            "Если браузер открывается с белым экраном или ERR_CONNECTION_RESET — "
            "сбросьте профиль и запустите сбор снова. Вход в Госуслуги потребуется заново.")
        gbtns.addWidget(self.collect_btn)
        gbtns.addWidget(self.continue_btn)
        gbtns.addWidget(self.cancel_btn)
        gbtns.addWidget(self.skip_btn)
        gbtns.addWidget(self.assist_btn)
        gbtns.addWidget(self.reset_profile_btn)
        gbtns.addStretch(1)
        gl.addLayout(gbtns)

        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        gl.addWidget(self.progress)
        root.addWidget(gbox)

        # Применение вручную
        abox = QGroupBox("Применить из кэша вручную (без браузера)")
        al = QVBoxLayout(abox)
        ahint = QLabel(
            "Переносит уже собранные форму и обременения из кэша в report.xlsx и "
            "PDF. Без запросов к API. Работает и в .exe.")
        ahint.setWordWrap(True)
        ahint.setStyleSheet("color: #555; font-size: 11px;")
        al.addWidget(ahint)
        self.apply_btn = QPushButton("Применить из кэша (Пересобрать)")
        self.apply_btn.clicked.connect(lambda: self.apply_from_cache())
        al.addWidget(self.apply_btn)
        root.addWidget(abox)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        root.addWidget(self.log, 1)

        if self.output_edit.text().strip():
            self._refresh_info()

    # -- helpers -------------------------------------------------------- #
    def _out_root(self) -> Path | None:
        raw = self.output_edit.text().strip()
        if not raw:
            return None
        root = resolve_output_root(raw)
        if str(root) != raw:
            self.output_edit.setText(str(root))
        return root

    def _refresh_info(self):
        raw = self.output_edit.text().strip()
        if not raw or not Path(raw).exists():
            self.info.setText("")
            return
        root = resolve_output_root(raw)
        n_json = count_cache_json(root)
        n_rr = _count_rosreestr_cache(root)
        self.info.setText(
            f"корень: {root}   |   объектов в кэше API: {n_json}   |   "
            f"собрано форм/обременений (Росреестр): {n_rr}")
        self.info.setStyleSheet(
            "color: #2e7d32; font-size: 11px;" if n_rr
            else "color: #555; font-size: 11px;")

    def pick_output(self):
        path = QFileDialog.getExistingDirectory(
            self, "Папка результатов (корень output)", self.output_edit.text())
        if path:
            self.output_edit.setText(str(resolve_output_root(path)))
            self._refresh_info()

    def _log(self, text: str):
        self.log.appendPlainText(text)

    def _set_busy(self, busy: bool):
        self.collect_btn.setEnabled(not busy)
        self.apply_btn.setEnabled(not busy)
        self.cancel_btn.setEnabled(busy)
        self.skip_btn.setEnabled(busy)
        if not busy:
            self.continue_btn.setEnabled(False)
            self.assist_btn.setEnabled(False)
            self.assisted_mode.setEnabled(True)

    # -- сбор ----------------------------------------------------------- #
    def start_collect(self):
        out = self._out_root()
        if not out or not out.exists():
            QMessageBox.warning(self, "Собственники", "Выберите папку результатов.")
            return
        if count_cache_json(out) == 0:
            QMessageBox.warning(
                self, "Нет кэша API",
                "Сначала обработайте объекты на вкладке «Пакетная обработка» — "
                "из их кэша берётся список кадастровых номеров.")
            return
        if not self.win.set_running(True):
            return
        self.log.clear()
        self.progress.setRange(0, 0)
        self._set_busy(True)
        self.assisted_mode.setEnabled(False)
        mode = "проверочный" if self.assisted_mode.isChecked() else "автоматический"
        self._log(f"Запуск {mode} сбора…")
        self.worker = _CollectWorker(
            out, self.only_rights.isChecked(), self.redo.isChecked(),
            assisted=self.assisted_mode.isChecked(),
            apply_after=self.apply_after.isChecked(),
            parent=self)
        self.worker.line.connect(self._log)
        self.worker.progress.connect(self._on_progress)
        self.worker.need_login.connect(self._on_need_login)
        self.worker.form_ready.connect(self._on_form_ready)
        self.worker.need_assist.connect(self._on_need_assist)
        self.worker.result.connect(self._on_result)
        self.worker.start()

    def _on_progress(self, i: int, n: int):
        if n:
            self.progress.setRange(0, n)
            self.progress.setValue(i)
            self.progress.setFormat(f"{i}/{n}")

    def _on_need_login(self):
        self.continue_btn.setEnabled(False)
        self._log("→ Войдите через Госуслуги, выберите пользователя. Программа сама "
                  "откроет «Справочную информацию online».")
        QMessageBox.information(
            self, "Вход в Росреестр",
            "Открылся браузер.\n\n"
            "1. Войдите через Госуслуги (кадастровый инженер, SMS-код).\n"
            "2. Выберите пользователя, если спросит.\n"
            "3. Программа сама перейдёт на «Справочную информацию online» "
            "(не «Мои заявки»).\n\n"
            "Когда кнопка «Вход выполнен — продолжить» станет активной — "
            "нажмите её.")

    def _on_form_ready(self):
        self.continue_btn.setEnabled(True)
        self._log("→ Форма ЛK готова (есть «Вид объекта»). "
                  "Нажмите «Вход выполнен — продолжить».")

    def on_continue(self):
        if self.worker:
            self.continue_btn.setEnabled(False)
            self._log("Продолжаю сбор…")
            self.worker.confirm_login()

    def on_cancel(self):
        if self.worker:
            self._log("Останавливаю…")
            self.worker.cancel()

    def on_skip(self):
        if self.worker:
            self._log("→ Пропуск текущего объекта…")
            self.worker.skip_current()

    def on_assist_confirm(self):
        if self.worker:
            self.assist_btn.setEnabled(False)
            self._log("→ Записываю данные с текущей страницы…")
            self.worker.confirm_assist()

    def _on_need_assist(self, kn: str):
        self.assist_btn.setEnabled(True)
        self._log(
            f"→ Проверочный режим [{kn}]: КН уже в поле поиска. "
            "Выберите «Вид объекта», нажмите «Найти», откройте справку — затем "
            "«Данные готовы — записать и далее».")

    def reset_browser_profile(self):
        if self.win.is_running():
            QMessageBox.warning(
                self, "Сброс профиля",
                "Сначала остановите текущий сбор.")
            return
        from ..rosreestr_collector import reset_browser_profile as do_reset
        reply = QMessageBox.question(
            self, "Сброс профиля браузера",
            "Профиль браузера будет сохранён в резервную копию и создан заново.\n\n"
            "После сброса при следующем сборе потребуется снова войти через "
            "Госуслуги.\n\n"
            "Продолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            backup = do_reset()
            self._log(f"Профиль браузера сброшен. Резервная копия: {backup}")
            QMessageBox.information(
                self, "Готово",
                "Профиль браузера сброшен.\n"
                "Запустите сбор снова — откроется Edge или Chrome, "
                "войдите через Госуслуги один раз.")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Ошибка", str(exc))

    def _on_result(self, res):
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self._set_busy(False)
        self.win.set_running(False)
        if getattr(res, "error", ""):
            self._log(f"Ошибка: {res.error}")
            QMessageBox.warning(self, "Собственники", res.error)
            return
        if getattr(res, "cancelled", False):
            self._log("Остановлено пользователем.")
        self._log(
            f"Готово. Сохранено: {res.saved} из {res.total}. "
            f"Пропущено (уже было): {res.skipped_existing}. Ошибок: {res.failed}.")
        if res.total and res.saved == 0 and not getattr(res, "cancelled", False):
            self._log(
                "→ Ни один объект не сохранён. Смотрите строки «не удалось выбрать» "
                "или «форма не распознана» выше. Убедитесь, что открыта вкладка "
                "«Справочная информация online» после входа в ЛК.")
            QMessageBox.warning(
                self, "Собственники",
                "Сбор завершён, но данные не сохранены.\n\n"
                "Частая причина — не выбран «Вид объекта» в форме ЛК.\n"
                "Проверьте журнал на вкладке и повторите сбор.\n"
                "Если браузер на другой странице — откройте «Справочную "
                "информацию online» и нажмите «Вход выполнен — продолжить».")
        self._refresh_info()
        out = self._out_root()
        if out:
            self.win.logs_tab.refresh_files(out)
        if (res.saved and self.apply_after.isChecked()
                and not (self.worker and self.worker.apply_after)):
            self._log("Применяю собранные данные к report.xlsx и PDF…")
            self.apply_from_cache(silent=True)

    # -- применение ----------------------------------------------------- #
    def apply_from_cache(self, silent: bool = False):
        out = self._out_root()
        if not out or not out.exists():
            if not silent:
                QMessageBox.warning(self, "Применить", "Выберите папку результатов.")
            return
        if count_cache_json(out) == 0:
            if not silent:
                QMessageBox.warning(
                    self, "Нет кэша API",
                    "В папке нет кэша ответов API (cache\\json).")
            return
        mt = self.win.maintenance_tab
        mt.output_edit.setText(str(out))
        self.win.tabs.setCurrentWidget(mt)
        mt.run_op("rebuild", apply=True)
        self._log("Пересборка запущена на вкладке «Обслуживание» — см. её журнал.")
