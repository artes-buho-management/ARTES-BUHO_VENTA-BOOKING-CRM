@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\sync-realtime.ps1" -Branch main -DebounceSeconds 20
