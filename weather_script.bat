@echo off
weather_script.exe
if %ERRORLEVEL% NEQ 0 (
	pause
) 
