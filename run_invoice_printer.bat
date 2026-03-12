@echo off
chcp 65001 >nul
set PYTHONUTF8=1
py -3.14 "%~dp0load_env_and_run.py" %*
