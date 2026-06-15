@echo off
chcp 65001 >nul
title IP Location Checker
color 0A

echo ============================================================
echo   IP LOCATION CHECKER - Khong can Python
echo   Su dung PowerShell + ip-api.com (mien phi, khong can key)
echo ============================================================
echo.

powershell -Command "exit 0" >nul 2>&1
if errorlevel 1 (
    echo [LOI] PowerShell khong kha dung.
    pause & exit /b 1
)

set "DIR=%~dp0"
set "INPUT=%DIR%ip_list.txt"
set "PS1=%DIR%ip_checker.ps1"

if not exist "%PS1%" (
    echo [LOI] Khong tim thay file: %PS1%
    echo Vui long dat file ip_checker.ps1 cung thu muc voi file nay.
    pause & exit /b 1
)

if not exist "%INPUT%" (
    echo [LOI] Khong tim thay file: %INPUT%
    echo.
    echo Vui long tao file ip_list.txt, moi dong 1 dia chi IP. Vi du:
    echo    8.8.8.8
    echo    1.1.1.1
    echo.
    echo Tao file mau 10 IP de thu? [Y/N]
    set /p ANS="> "
    if /i "%ANS%"=="Y" (
        (
            echo 8.8.8.8
            echo 1.1.1.1
            echo 103.21.244.0
            echo 45.33.32.156
            echo 208.67.222.222
            echo 9.9.9.9
            echo 94.140.14.14
            echo 64.6.64.6
            echo 156.154.70.1
            echo 216.146.35.35
        ) > "%INPUT%"
        echo [OK] Da tao ip_list.txt mau.
        echo.
    ) else (
        pause & exit /b 1
    )
)

for /f "usebackq" %%T in (`powershell -NoProfile -Command "Get-Date -Format 'yyyyMMdd_HHmmss'"`) do set "TS=%%T"
set "OUTFILE=%DIR%IP_Location_%TS%.xlsx"

echo [*] Input : %INPUT%
echo [*] Output: %OUTFILE%
echo.
echo Dang kiem tra IP, vui long cho...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%INPUT%" "%OUTFILE%"
set PS_EXIT=%errorlevel%

if %PS_EXIT% neq 0 (
    echo.
    echo [LOI] Co loi khi chay. Kiem tra ket noi mang va thu lai.
    pause & exit /b 1
)

echo.
echo ============================================================
echo  HOAN THANH!  File: %OUTFILE%
echo  Gom 3 sheet: Summary / IP Location Detail / Country Summary
echo ============================================================
echo.
set /p OPENF="Mo file Excel ngay bay gio? [Y/N]: "
if /i "%OPENF%"=="Y" start "" "%OUTFILE%"
echo.
pause
