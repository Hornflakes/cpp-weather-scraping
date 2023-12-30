@echo off
%~dp0app/weather_script.exe
if %ERRORLEVEL% NEQ 0 (
	pause
) 
