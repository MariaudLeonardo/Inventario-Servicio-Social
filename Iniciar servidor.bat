@echo off
title Servidor de inventario
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0servidor.ps1"
pause
