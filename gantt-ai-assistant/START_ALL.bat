@echo off
title Gantt AI Assistant

set "PROJ=c:\Users\hp\Desktop\aaaadaatassets\project\HEART-FAILURE-DEATH-PREDICTOR\gantt-ai-assistant"
set "PYTHON=%PROJ%\.venv\Scripts\python.exe"
set "CERT=%PROJ%\certs\localhost.crt"
set "KEY=%PROJ%\certs\localhost.key"

echo Starting Gantt AI Assistant...

:: Start FastAPI backend (minimized)
start /min "FastAPI Backend" cmd /c "cd /d %PROJ% && %PYTHON% -m uvicorn agent_backend:app --host 0.0.0.0 --port 8000 --ssl-keyfile %KEY% --ssl-certfile %CERT%"

:: Start Task Pane HTTPS server (minimized)
start /min "Task Pane Server" cmd /c "cd /d %PROJ%\gantt-outlook-addin && %PYTHON% https_server.py"

:: Start Streamlit (minimized)
start /min "Streamlit Dashboard" cmd /c "cd /d %PROJ% && %PYTHON% -m streamlit run app.py"

:: Wait for servers to be ready then open dashboard
timeout /t 5 /nobreak >nul
start "" "http://localhost:8501"

echo Gantt AI Assistant is running.
echo To stop: close the minimized server windows in the taskbar.

