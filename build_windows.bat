@echo off
REM Автоматическая сборка Avito Price Analyzer в .exe (Windows)
REM Требуется установленный Python 3.10+

setlocal
set PROJDIR=%~dp0
cd /d "%PROJDIR%"

echo [1/4] Создаю виртуальное окружение...
python -m venv venv || goto :error

echo [2/4] Активирую окружение и ставлю зависимости...
call venv\Scripts\activate.bat || goto :error
python -m pip install --upgrade pip || goto :error
pip install -r requirements.txt pyinstaller || goto :error

echo [3/4] Сборка через PyInstaller...
pyinstaller AvitoPriceAnalyzer.spec || goto :error

echo [4/4] Готово!
echo Файл: dist\AvitoPriceAnalyzer\AvitoPriceAnalyzer.exe
pause
exit /b 0

:error
echo Ошибка сборки. Проверьте, что установлен Python и интернет соединение для pip.
pause
exit /b 1
