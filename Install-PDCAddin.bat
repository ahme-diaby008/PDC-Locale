@echo off
setlocal

echo ============================================
echo  Installation du complement PDC (sans admin)
echo ============================================

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/manifest.xml"
set "PS_TMP=%TEMP%\manifest.xml"
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

echo Telechargement du manifest...
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PS_URL%' -OutFile '%PS_TMP%' -UseBasicParsing" || (
  echo Echec du telechargement. Verifiez l'acces Internet/proxy.
  pause
  exit /b 1
)

if not exist "%WEF%" mkdir "%WEF%"

echo Nettoyage du cache WEF...
if exist "%WEF%\Cache" rmdir /s /q "%WEF%\Cache"

echo Fermeture d'Excel/Word/PowerPoint s'ils sont ouverts...
taskkill /f /im excel.exe >nul 2>&1
taskkill /f /im winword.exe >nul 2>&1
taskkill /f /im powerpnt.exe >nul 2>&1

echo Installation du manifest...
copy /y "%PS_TMP%" "%WEF%\manifest.xml" >nul

echo.
echo ============================================
echo  ✔ Installation terminee. Ouvrez Excel.
echo ============================================
pause
endlocal
