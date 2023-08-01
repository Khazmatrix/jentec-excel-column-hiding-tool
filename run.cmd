@echo off
setlocal

rem Get the current directory
set "current_dir=%~dp0"

rem Run the Python script
python "%current_dir%script.py"

pause