@echo off
chcp 65001 > nul
echo ========================================================
echo   Starting Mobile Specs Comparison Tool (Streamlit)...
echo ========================================================
python -m streamlit run app.py
pause
