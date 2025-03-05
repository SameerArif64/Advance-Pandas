@echo off
set LIBRARY_DIR="D:\Work\Side Projects\My-Python-Libraries\file_utils"

echo Installing file_utils...

cd %LIBRARY_DIR%
rmdir /s /q build 2>nul
python setup.py build
pip install -e .

echo file_utils installed successfully!
pause
