@echo off
call venv\Scripts\activate.bat
cd /d %~dp0
uvicorn main:app --reload --host 0.0.0.0 --port=8000
