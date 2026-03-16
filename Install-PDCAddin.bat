@echo off
echo ============================================
echo  Installation du complément PDC...
echo ============================================

REM URL du script PowerShell sur GitHub RAW
set PS_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/Install-OfficeAddin.ps1

REM Chemin local temporaire
set PS_FILE=%TEMP%\Install-OfficeAddin.ps1

echo.
echo Téléchargement du script depuis GitHub...
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest '%PS_URL%' -OutFile '%PS_FILE%'"

IF NOT EXIST "%PS_FILE%" (
  echo.
  echo ❌ ÉCHEC : Impossible de télécharger le script.
  echo Vérifiez l'accès Internet ou le proxy.
  pause
  exit /b 1
)

echo.
echo Lancement du script d'installation PowerShell...
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_FILE%"

echo.
echo ============================================
echo  ✔ Installation terminée
echo  Vous pouvez maintenant ouvrir Excel.
echo ============================================
pause
