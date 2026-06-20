# ЕГРН-Парсер (reestro + reestro_app)

Монорепозиторий проекта: движок парсера ЕГРН и десктопное приложение.

| Папка | Назначение |
|-------|------------|
| [`reestro/`](reestro/) | CLI-движок: API Контур.Реестро, PDF, Excel, кэш |
| [`reestro_app/`](reestro_app/) | GUI (PySide6), сборка `.exe`, инструкции |

## Быстрый старт (разработка)

```bat
cd reestro_app
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python -m egrn_gui.main
```

## Сборка для заказчика

```bat
cd reestro_app\build
build_exe.bat
python ..\release\prepare_delivery.py
```

## Документация

- Пользователь: [`reestro_app/ИНСТРУКЦИЯ.md`](reestro_app/ИНСТРУКЦИЯ.md)
- Пошагово со скринами: [`reestro_app/docs/ПОШАГОВАЯ_ИНСТРУКЦИЯ.md`](reestro_app/docs/ПОШАГОВАЯ_ИНСТРУКЦИЯ.md)
- CLI-движок: [`reestro/ИНСТРУКЦИЯ.md`](reestro/ИНСТРУКЦИЯ.md)
- Архитектура GUI: [`reestro_app/ARCHITECTURE.md`](reestro_app/ARCHITECTURE.md)

## Версионирование

Теги релизов: `v0.1.0`, `v0.1.1`, …

*by BeRealBear*
