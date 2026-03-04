@echo off
echo ========================================
echo   Kascade - Deploy to GitHub
echo ========================================
echo.

cd /d "C:\Users\klaud\Documents\Claude Projects\cascade-tools"

:: Stage all changes
git add -A

:: Ask for a commit message
set /p msg="Describe your changes (or press Enter for 'update'): "
if "%msg%"=="" set msg=update

:: Commit and push
git commit -m "%msg%"
git push origin main

echo.
echo ========================================
echo   Deployed! Changes will be live in
echo   about 1-2 minutes on GitHub Pages.
echo ========================================
pause
