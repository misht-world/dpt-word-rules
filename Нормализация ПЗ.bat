@echo off
rem Запуск GUI нормализации Word (ПЗ). Кладётся рядом с gui.py.
cd /d "%~dp0"
where pythonw >nul 2>nul && (
    start "" pythonw "%~dp0gui.py"
) || (
    python "%~dp0gui.py"
)
