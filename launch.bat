@echo off
title ExcelAI Launcher
color 0A

echo ============================================
echo    ExcelAI - Powered by Ollama gemma4:e4b
echo ============================================
echo.

rem Check Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 goto :NoPython
echo [OK] Python found.

rem Check if Ollama is running
echo [1/3] Checking Ollama...
curl -s http://localhost:11434/api/tags >nul 2>&1
if %errorlevel% neq 0 goto :StartOllama
echo [OK] Ollama is already running.
goto :CheckDeps

:StartOllama
echo [INFO] Starting Ollama in background...
start "" ollama serve
timeout /t 4 /nobreak >nul
echo [OK] Ollama started.

:CheckDeps
rem Install dependencies if needed
echo [2/3] Checking Python dependencies...
python -m pip install xlwings ollama --quiet
echo [OK] Dependencies ready.

rem Check model exists
echo [3/3] Checking gemma4:e4b model...
ollama list | find "gemma4:e4b" >nul 2>&1
if %errorlevel% neq 0 goto :PullModel
echo [OK] gemma4:e4b model found.
goto :LaunchApp

:PullModel
echo [INFO] Pulling gemma4:e4b model (first time only, may take a few minutes)...
ollama pull gemma4:e4b
echo [OK] Model pulled successfully.

:LaunchApp
echo.
echo [LAUNCH] Starting ExcelAI...
echo.
python main.py

echo.
echo [EXIT] App closed. Press any key to exit the launcher.
pause
exit /b

:NoPython
echo [ERROR] Python not found. Please install Python 3.10+
pause
exit /b
