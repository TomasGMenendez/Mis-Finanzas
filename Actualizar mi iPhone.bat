@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   Actualizando el Dashboard con el Excel...
echo ============================================
echo.

cd /d "%~dp0Programas"
python generar_datos_html.py
if errorlevel 1 (
  echo.
  echo Algo fallo generando los datos. Revisa el mensaje de arriba.
  pause
  exit /b 1
)
cd /d "%~dp0"

echo.
echo ============================================
echo   Subiendo los cambios a Internet...
echo ============================================
echo.

git add "Mi Dashboard de Finanzas.html" "docs/index.html"
git commit -m "Actualizar datos del dashboard"
if errorlevel 1 (
  echo.
  echo No habia cambios nuevos para subir. Tu iPhone ya esta al dia.
  pause
  exit /b 0
)

git push
if errorlevel 1 (
  echo.
  echo Algo fallo al subir a Internet. Revisa que tengas conexion
  echo y volve a hacer doble clic en este archivo.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   Listo! En unos minutos tu iPhone va a
echo   mostrar los datos nuevos.
echo ============================================
pause
