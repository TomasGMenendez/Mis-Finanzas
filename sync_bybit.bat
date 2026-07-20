@echo off
REM Auto-sync Bybit → Excel Finanzas Toto
REM Corre bybit_sync.py --sync y guarda log

set FOLDER=%~dp0
cd /d "%FOLDER%"

REM Timestamp para el log
for /f "tokens=1-4 delims=/ " %%i in ("%date%") do (set FECHA=%%l-%%k-%%j)
for /f "tokens=1-2 delims=: " %%i in ("%time%") do (set HORA=%%i-%%j)

echo. >> sync_log.txt
echo ============================================== >> sync_log.txt
echo Sync: %FECHA% %HORA% >> sync_log.txt
echo ============================================== >> sync_log.txt

python bybit_sync.py --sync >> sync_log.txt 2>&1

REM Si falla, el mensaje queda en sync_log.txt
