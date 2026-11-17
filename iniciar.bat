@echo off
cd /d "C:\Users\salvi\Documents\anacom-ited"
start "ANACOM ITED" /min "C:\Users\salvi\anaconda3\python.exe" server.py
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"
