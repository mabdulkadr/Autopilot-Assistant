@ECHO OFF
TITLE Autopilot Assistant
ECHO Starting Autopilot Assistant...
POWERSHELL.EXE -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Autopilot Assistant.ps1"
IF ERRORLEVEL 1 ECHO The script exited with an error.
PAUSE
