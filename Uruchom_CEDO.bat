@echo off
cd /d "%~dp0"
%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Generator_CEDO.ps1"
if errorlevel 1 (
 echo.
 echo Program zakonczyl sie bledem. Zobacz komunikat powyzej.
 pause
)
