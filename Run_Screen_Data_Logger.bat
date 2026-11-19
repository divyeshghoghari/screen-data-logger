@echo off
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt
py -3 screen_data_logger.py
pause
