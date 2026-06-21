# -*- coding: utf-8 -*-
"""Вкладка «Подключение»: параметры API и проверка связи."""
from __future__ import annotations

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QHBoxLayout,
    QSpinBox, QDoubleSpinBox, QLabel, QGroupBox, QCheckBox, QMessageBox,
)

from ..engine_bridge import check_balance


class _BalanceWorker(QThread):
    done = Signal(dict)

    def __init__(self, cfg: dict, parent=None):
        super().__init__(parent)
        self._cfg = cfg

    def run(self):
        self.done.emit(check_balance(self._cfg))


class ConnectionTab(QWidget):
    def __init__(self, win):
        super().__init__()
        self.win = win
        self._worker: _BalanceWorker | None = None
        cfg = win.config

        root = QVBoxLayout(self)

        api_box = QGroupBox("Параметры подключения к Контур.Реестро")
        form = QFormLayout(api_box)
        self.base_url = QLineEdit(cfg.get("baseUrl", "https://api.kontur.ru"))
        self.base_url.setPlaceholderText("https://api.kontur.ru")
        self.api_key = QLineEdit(cfg.get("apiKey", ""))
        self.api_key.setEchoMode(QLineEdit.Password)
        self.api_key.setPlaceholderText("напр. 91e4ea5e-35bd-fd08-4c12-43f4bc178beb")
        self.api_key.setToolTip("Ключ API из личного кабинета Контур.Реестро "
                                "(формат UUID)")
        self.org_id = QLineEdit(cfg.get("orgId", ""))
        self.org_id.setPlaceholderText("напр. d92eaeed-cfc1-4a8e-8c0b-25886c731df0")
        self.org_id.setToolTip("Идентификатор организации (portal.orgid), формат UUID")
        self.proxy = QLineEdit(cfg.get("proxy", ""))
        self.proxy.setPlaceholderText("необязательно, напр. http://proxy.company.ru:8080")

        self.show_key = QCheckBox("Показать ключ")
        self.show_key.toggled.connect(
            lambda v: self.api_key.setEchoMode(
                QLineEdit.Normal if v else QLineEdit.Password))

        form.addRow("baseUrl:", self.base_url)
        form.addRow("apiKey:", self.api_key)
        form.addRow("", self.show_key)
        form.addRow("orgId:", self.org_id)
        form.addRow("proxy:", self.proxy)
        root.addWidget(api_box)

        net_box = QGroupBox("Параметры сети по умолчанию")
        nform = QFormLayout(net_box)
        self.timeout = QSpinBox()
        self.timeout.setRange(10, 600)
        self.timeout.setValue(int(cfg.get("timeout", 120)))
        self.timeout.setSuffix(" сек")
        self.retries = QSpinBox()
        self.retries.setRange(1, 20)
        self.retries.setValue(int(cfg.get("retries", 5)))
        self.pause = QDoubleSpinBox()
        self.pause.setRange(0.0, 30.0)
        self.pause.setSingleStep(0.5)
        self.pause.setValue(float(cfg.get("pause", 2.0)))
        self.pause.setSuffix(" сек")
        self.save_every = QSpinBox()
        self.save_every.setRange(0, 1000)
        self.save_every.setValue(int(cfg.get("save_every", 5)))
        nform.addRow("Таймаут ответа:", self.timeout)
        nform.addRow("Повторов при обрыве:", self.retries)
        nform.addRow("Пауза между запросами:", self.pause)
        nform.addRow("Автосохранение каждые (объектов):", self.save_every)
        root.addWidget(net_box)

        # Росреестр: логин кадастрового инженера (для сбора формы собственности).
        rr_box = QGroupBox("Госуслуги / Росреестр (для вкладки «Собственники»)")
        rform = QFormLayout(rr_box)
        self.gosuslugi_login = QLineEdit(
            self.win.settings.get("gosuslugi_login", ""))
        self.gosuslugi_login.setPlaceholderText(
            "логин/телефон/email кадастрового инженера (для удобства)")
        self.gosuslugi_login.setToolTip(
            "Только для подсказки. Пароль и код подтверждения вводятся в браузере "
            "вручную и НИГДЕ не сохраняются (Госуслуги требуют 2FA).")
        rform.addRow("Логин Госуслуг:", self.gosuslugi_login)
        warn = QLabel(
            "Пароль не хранится в программе: вход через Госуслуги выполняется в "
            "браузере с подтверждением по коду. Достаточно войти один раз — сессия "
            "сохранится. Сбор данных — на вкладке «Собственники (Росреестр)».")
        warn.setWordWrap(True)
        warn.setStyleSheet("color: #827717; font-size: 11px;")
        rform.addRow("", warn)
        root.addWidget(rr_box)

        btns = QHBoxLayout()
        self.save_btn = QPushButton("Сохранить")
        self.save_btn.clicked.connect(self.on_save)
        self.test_btn = QPushButton("Проверить подключение")
        self.test_btn.clicked.connect(self.on_test)
        btns.addWidget(self.save_btn)
        btns.addWidget(self.test_btn)
        btns.addStretch(1)
        root.addLayout(btns)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        root.addWidget(self.status)
        root.addStretch(1)

    def _collect(self) -> dict:
        cfg = self.win.config
        cfg["baseUrl"] = self.base_url.text().strip() or "https://api.kontur.ru"
        cfg["apiKey"] = self.api_key.text().strip()
        cfg["orgId"] = self.org_id.text().strip()
        cfg["proxy"] = self.proxy.text().strip()
        cfg["timeout"] = self.timeout.value()
        cfg["retries"] = self.retries.value()
        cfg["pause"] = self.pause.value()
        cfg["save_every"] = self.save_every.value()
        # логин Госуслуг храним в settings (не в config с API-ключами); без пароля
        self.win.settings["gosuslugi_login"] = self.gosuslugi_login.text().strip()
        return cfg

    def on_save(self):
        self._collect()
        self.win.save_all()
        self.status.setStyleSheet("color: #2e7d32;")
        self.status.setText(f"Сохранено в {self.win and ''}config.json")

    def on_test(self):
        cfg = self._collect()
        if not cfg["apiKey"] or not cfg["orgId"]:
            QMessageBox.warning(self, "Подключение",
                                "Укажите apiKey и orgId.")
            return
        self.test_btn.setEnabled(False)
        self.status.setStyleSheet("color: #555;")
        self.status.setText("Проверяю подключение...")
        self._worker = _BalanceWorker(cfg, self)
        self._worker.done.connect(self._on_result)
        self._worker.start()

    def _on_result(self, res: dict):
        self.test_btn.setEnabled(True)
        if res.get("ok"):
            self.status.setStyleSheet("color: #2e7d32;")
        else:
            self.status.setStyleSheet("color: #c62828;")
        self.status.setText(res.get("message", ""))
