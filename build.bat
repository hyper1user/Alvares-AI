@echo off
chcp 65001 >nul
echo ========================================
echo   АЛЬВАРЕС AI — Збірка інсталятора
echo ========================================
echo.

echo [1/2] Збірка PyInstaller...
python -m PyInstaller alvares.spec --noconfirm
if errorlevel 1 (
    echo ПОМИЛКА: PyInstaller завершився з помилкою!
    pause
    exit /b 1
)
echo PyInstaller — OK
echo.

echo [2/2] Збірка Inno Setup...
where iscc >nul 2>nul
if errorlevel 1 (
    if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
    ) else (
        echo УВАГА: ISCC.exe не знайдено. Встановіть Inno Setup 6 або додайте ISCC.exe в PATH.
        echo Ви можете відкрити installer.iss вручну через Inno Setup GUI.
        pause
        exit /b 1
    )
) else (
    iscc installer.iss
)
if errorlevel 1 (
    echo ПОМИЛКА: Inno Setup завершився з помилкою!
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Готово! Інсталятор: Output\AlvaresAI_Setup.exe
echo ========================================
pause
