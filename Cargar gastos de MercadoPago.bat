@echo off
setlocal
if "%~1"=="" (
  echo Arrastra tu archivo CSV de MercadoPago y soltalo arriba de este archivo.
  echo ^(No lo abras haciendo doble clic directamente^)
  echo.
  pause
  exit /b
)

cd /d "%~dp0Programas"

echo ============================================
echo   Cargando movimientos de MercadoPago...
echo ============================================
echo.

python mercadopago_processor.py "%~1"
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
