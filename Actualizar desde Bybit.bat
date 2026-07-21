@echo off
setlocal
cd /d "%~dp0Programas"

echo ============================================
echo   Actualizando tu portafolio desde Bybit...
echo ============================================
echo.

python bybit_sync.py --sync
if errorlevel 1 (
  echo.
  echo Algo fallo. Revisa el mensaje de arriba.
  pause
  exit /b 1
)

echo.
echo Actualizando el Dashboard con los datos nuevos...
python generar_datos_html.py

echo.
echo ============================================
echo   Listo. Abri "Mi Dashboard de Finanzas.html"
echo   para ver los cambios.
echo ============================================
pause
