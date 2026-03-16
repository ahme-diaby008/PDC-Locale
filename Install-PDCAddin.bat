@echo off
setlocal

REM ------------------------------------------------------------
REM Installe automatiquement l’addin PDC depuis GitHub (manifest WEF)
REM ------------------------------------------------------------

REM 1) URL du script PowerShell hébergé sur GitHub RAW
set "PS_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/Install-OfficeAddin.ps1"

REM 2) Emplacement temporaire où enregistrer le script
set "PS_FILE=%TEMP%\Install-OfficeAddin.ps1"

echo Téléchargement du script d'installation...
REM Utilise PowerShell pour télécharger le fichier (compatible partout)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; ^
   Invoke-WebRequest -Uri '%PS_URL%' -OutFile '%PS_FILE%' -UseBasicParsing"

IF NOT EXIST "%PS_FILE%" (
  echo Echec du téléchargement du script. Verifiez l'accès Internet / proxy / URL.
  pause
  exit /b 1
)

echo Lancement de l'installation...
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

echo.
echo ============================================
echo  Installation terminee. Vous pouvez ouvrir Excel.
echo ============================================
echo.
pause
endlocal
