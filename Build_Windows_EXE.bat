@echo off
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt pyinstaller
py -3 -m PyInstaller --noconfirm --clean --windowed --name ScreenDataLogger screen_data_logger.py
echo.
echo EXE is in the dist folder.
pause
