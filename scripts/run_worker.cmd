@echo off
REM Entry point invoked by Windows Task Scheduler.
REM Keeps logging encoding sane and ensures CWD is the project root.

setlocal
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
cd /d "%~dp0\.."
python worker.py
exit /b %ERRORLEVEL%
