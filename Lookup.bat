@echo off
SET /P fstr=Search string: 
findstr /S /I %fstr% *.txt
echo.
echo.
lookup