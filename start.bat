@echo off
call venv\Scripts\activate.bat
cd C:\livinglabpy
uvicorn main:app --reload --host 0.0.0.0 --port=8000
