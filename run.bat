@echo off
cd /d "%~dp0"
python -m streamlit run dashboard_seasonal_monthly.py
pause