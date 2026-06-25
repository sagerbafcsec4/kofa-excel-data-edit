@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py  >nul 2>&1 && ( start "" py -3 "soccer_format_app.pyw" & exit /b )
where python >nul 2>&1 && ( start "" python "soccer_format_app.pyw" & exit /b )
echo Python が見つかりません。先に EXEを作成.bat で .exe を作るか、Python を入れてください。
pause
