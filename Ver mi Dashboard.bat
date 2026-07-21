@echo off
start "" msedge.exe "%~dp0Mi Dashboard de Finanzas.html"
if errorlevel 1 (
  start "" "%~dp0Mi Dashboard de Finanzas.html"
)
