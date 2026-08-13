@echo off
setlocal
rem Codepage UTF-8 para que los acentos se vean bien en esta ventana
chcp 65001 >nul
cd /d "%~dp0Programas"

echo ============================================
echo   Abriendo tu CRM de Finanzas...
echo ============================================
echo.
echo   En unos segundos se abre solo en el navegador.
echo   DEJA ESTA VENTANA ABIERTA mientras lo uses.
echo.

python crm_server.py

echo.
echo El CRM se cerro. Podes cerrar esta ventana.
pause
