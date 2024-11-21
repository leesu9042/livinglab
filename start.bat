@echo off
call C:\pythonProject1\venv\Scripts\activate.bat
cd C:\pythonProject1
uvicorn main:app --reload --host 0.0.0.0 --port=8000
