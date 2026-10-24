@echo off
cd /d "%~dp0"
if not exist logs mkdir logs
set PYTHONIOENCODING=utf-8
echo ==== Run started %date% %time% in %cd% ==== >> logs\run_log.txt
".venv\Scripts\python.exe" src\report.py >> logs\run_log.txt 2>&1
set RC=%errorlevel%
echo Exit code: %RC% >> logs\run_log.txt
echo. >> logs\run_log.txt
exit /b %RC%