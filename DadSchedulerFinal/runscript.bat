@echo off
python C:\Users\user1\Documents\PythonRPA\DadScheduler\webscraper.py
IF %ERRORLEVEL% EQU 0 (
  python C:\Users\user1\Documents\PythonRPA\DadScheduler\cleanup.py
) ELSE (
  echo Scraping failed, report generation skipped
)