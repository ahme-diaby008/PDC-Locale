@echo off
REM ============================================
REM  Installateur PDC - Batch minimal et robuste
REM ============================================

REM 1) URL du script PowerShell heberge sur GitHub RAW
set "PS_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/Install-OfficeAddin.ps1"

REM 2) Chemin Powershell (Windows PowerShell)
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

REM 3) Fichier temporaire pour le script
set "PS_FILE=%TEMP%\Install-OfficeAddin.ps1"

echo Downloading installer script...

"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PS_URL%' -OutFile '%PS_FILE%' -UseBasicParsing } catch { Write-Host ('Download failed: ' + $_.Exception.Message); exit 1 }"

IF NOT EXIST "%PS_FILE%" (
  echo Failed to download the installer. Check Internet/proxy/URL.
  pause
  exit /b 1
)

echo Running installer...
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

echo.
echo =========== DONE ===========
echo You can now open Excel.
echo ============================
echo.
pause
