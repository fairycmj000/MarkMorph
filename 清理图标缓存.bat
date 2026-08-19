@echo off
echo 正在清理 Windows 图标缓存...
echo.

taskkill /f /im explorer.exe
timeout /t 2 /nobreak >nul

del /a /q "%localappdata%\IconCache.db"
del /a /f /q "%localappdata%\Microsoft\Windows\Explorer\iconcache*"

echo.
echo 图标缓存已清理完成！
echo 正在重启资源管理器...
start explorer.exe

echo.
echo 完成！请检查应用图标是否已更新。
pause
