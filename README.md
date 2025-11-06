# Inkcut Launcher

Небольшой лаунчер `run_inkcut.py`, который автоматически подготавливает окружение для запуска [Inkcut](https://www.codelv.com/projects/inkcut/) на Windows.

## Возможности

- Проверяет наличие Python 3.9.13 и при необходимости устанавливает его в тихом режиме.
- Настраивает переменные среды и, при запуске из exe, использует системный интерпретатор.
- Устанавливает зависимости `pyqt5` и `inkcut`, отображая прогресс и сохраняя его в лог.
- Запускает Inkcut командой `python -m inkcut` после успешной подготовки.
- Ведёт журнал действий в файле `inkcut_launcher.log` рядом со скриптом или exe.

## Быстрый старт

1. Скачайте `run_inkcut.py` и запустите его двойным кликом или через PowerShell:
   ```powershell
   python run_inkcut.py
   ```
2. Дождитесь завершения проверки Python и установки зависимостей.
3. После подготовки откроется Inkcut.

## Сборка исполняемого файла

Для упаковки в один exe используйте PyInstaller (установив его при необходимости):
```powershell
pip install pyinstaller
pyinstaller --onefile --windowed run_inkcut.py
```
Полученный файл `dist\run_inkcut.exe` можно запускать на любой поддерживаемой версии Windows.

## Логи и диагностика

- Все сообщения выводятся в консоль и дублируются в `inkcut_launcher.log`.
- При ошибках отображается всплывающее окно с описанием проблемы.

## Ссылки

- Основное приложение Inkcut: [https://www.codelv.com/projects/inkcut/](https://www.codelv.com/projects/inkcut/)
- Документация PyInstaller: [https://pyinstaller.org/en/stable/](https://pyinstaller.org/en/stable/)
