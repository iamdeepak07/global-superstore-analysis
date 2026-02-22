@echo off
cd /d "c:\Users\ASUS\OneDrive\Documents\Data Analyst"
echo Starting Global Superstore Dashboard...
echo Dashboard will open in your default browser at http://localhost:8501
timeout /t 2
streamlit run streamlit_dashboard.py
pause
