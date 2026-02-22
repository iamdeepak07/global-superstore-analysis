@echo off
REM Global Superstore Dashboard Launcher
echo.
echo ====================================
echo Global Superstore Dashboard
echo ====================================
echo.
echo Starting Streamlit server...
echo.
echo Dashboard will open at:
echo http://localhost:8501
echo.
echo Press CTRL+C to stop the server
echo.
timeout /t 3

cd /d "c:\Users\ASUS\OneDrive\Documents\Data Analyst"
python -m streamlit run app.py --server.port=8501 --logger.level=off

pause
