@echo off
title ExcelAI Launcher
color 0A

echo ============================================
echo    ExcelAI - AI-Powered Excel Controller
echo    Powered by Ollama
echo ============================================
echo.

rem Check Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 goto :NoPython
echo [OK] Python found.

rem Create session directory
if not exist "%USERPROFILE%\.excelai" mkdir "%USERPROFILE%\.excelai"
echo [OK] Session directory ready.

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

rem Check if any model is available (not just gemma4)
echo [3/3] Checking Ollama models...
ollama list 2>nul | findstr /R "gemma llama qwen mistral deepseek phi" >nul 2>&1
if %errorlevel% neq 0 goto :PullModel
echo [OK] At least one Ollama model found.
goto :LaunchApp

:PullModel
echo [INFO] No common model found. Pulling gemma4:e4b...
echo        (You can select other models later from the app.)
ollama pull gemma4:e4b
echo [OK] Model pulled successfully.

:LaunchApp
echo.
echo [LAUNCH] Starting ExcelAI...
echo          - Pick your model from the Controls tab
echo          - Toggle Dry-Run to preview before executing
echo          - Toggle Analysis Mode to analyze data
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
