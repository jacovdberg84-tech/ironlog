@echo off
setlocal
cd /d C:\IRONLOG\api
set NODE_ENV=production
if not exist "C:\IRONLOG\api\logs" mkdir "C:\IRONLOG\api\logs"
"C:\Program Files\nodejs\node.exe" index.js >> "C:\IRONLOG\api\logs\startup.log" 2>&1
