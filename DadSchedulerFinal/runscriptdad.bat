@echo on
echo %date% %time% - Starting script execution >> C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\log.txt

:: Set working directory
cd /d C:\Users\Mouhamadou\Documents\PythonProject\Scheduler

:: Set any necessary environment variables
set PYTHONIOENCODING=utf-8

:: Activate virtual environment
call C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\.venv\Scripts\activate.bat

:: Output environment information for debugging
echo Current directory: >> C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\log.txt
cd >> C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\log.txt
echo Python path: >> C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\log.txt
where python >> C:\Users\Mouhamadou\Documents\PythonProject\Scheduler\log.txt

:: Run with full paths
echo %date% %time% - Running scraper script... >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
python C:\Users\user1\Documents\PythonRPA\DadScheduler\webscraper.py > C:\Users\user1\Documents\PythonRPA\DadScheduler\scraper_output.log 2>&1
IF %ERRORLEVEL% EQU 0 (
  echo %date% %time% - Scraper completed successfully >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
  echo %date% %time% - Running report script... >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
  python C:\Users\user1\Documents\PythonRPA\DadScheduler\cleanup.py > C:\Users\user1\Documents\PythonRPA\DadScheduler\report_output.log 2>&1
) ELSE (
  echo %date% %time% - Scraper failed with error code %ERRORLEVEL% >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
  type C:\PythonProject\scraper_output.log >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
  echo Scraping failed, report generation skipped
)

:: Deactivate virtual environment
deactivate

:: Signal task completion
echo %date% %time% - Task completed >> C:\Users\user1\Documents\PythonRPA\DadScheduler\log.txt
exit /b 0