@echo off
REM Launches the Bulk Edit Builder Streamlit app and opens it in the default browser.
REM Close this terminal window to stop the app.

cd /d "%~dp0"

REM Open the browser shortly after startup (Streamlit takes ~2s to bind).
start "" /B cmd /c "timeout /t 3 /nobreak >nul && start """" http://localhost:8501"

streamlit run app.py
