@echo off
cd /d "c:\cerveau 2 obsidian vault\ceveau 2 vault\claude safe"

C:\Users\arnau\AppData\Local\Programs\Python\Python312\python.exe coaching_followup.py --all 2>nul

powershell -ExecutionPolicy Bypass -File "c:\cerveau 2 obsidian vault\ceveau 2 vault\claude safe\notify_coaching.ps1"
