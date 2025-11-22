@echo off
echo Installing/updating Python dependencies...
pip install -r requirements.txt

echo.
echo Running content extraction...
python extract_from_word.py

pause

