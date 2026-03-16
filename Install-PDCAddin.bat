@echo off
setlocal

echo ============================================
echo  Installation PDC (nettoyage anti-debug inclus)
echo ============================================

REM --- Chemins utiles ---
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
set "WEF_CACHE=%WEF%\Cache"
set "DEV_ADDINS=%LOCALAPPDATA%\Microsoft\Office\Addins"
set "REG_WEF=HKCU\Software\Microsoft\Office\16.0\WEF"

REM --- 0) Fermer Office ---
taskkill /f /im excel.exe >nul 2>&1
taskkill /f /im winword.exe >nul 2>&1
taskkill /f /im powerpnt.exe >nul 2>&1

REM --- 1) Nettoyage registre (mode developpeur) dans HKCU : pas besoin d'admin ---
reg delete "%REG_WEF%" /v Developer /f >nul 2>&1
reg delete "%REG_WEF%" /v DeveloperTools /f >nul 2>&1
reg delete "%REG_WEF%" /v RuntimeDebugging /f >nul 2>&1
reg delete "%REG_WEF%" /v RuntimeLogging /f >nul 2>&1
reg delete "%REG_WEF%" /v UseSharedRuntimeDebugging /f >nul 2>&1

REM --- 2) Supprimer le dossier de sideload DEV s'il existe ---
if exist "%DEV_ADDINS%" (
  echo Suppression du dossier DEV: %DEV_ADDINS%
  rmdir /s /q "%DEV_ADDINS%"
)

REM --- 3) Purge WEF ---
if exist "%WEF_CACHE%" rmdir /s /q "%WEF_CACHE%"
if not exist "%WEF%" mkdir "%WEF%"

REM --- 4) Telecharger le manifest (remplace l'URL par la tienne si besoin) ---
set "MAN_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/manifest.xml"
set "MAN_TMP=%TEMP%\manifest.xml"

echo Telechargement du manifest...
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MAN_URL%' -OutFile '%MAN_TMP%' -UseBasicParsing" ^
  || (echo Echec du telechargement. Verifiez l'acces Internet/proxy.& pause & exit /b 1)

REM --- 5) Remplacer la SourceLocation par /index.html si besoin (sécurite) ---
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p='%MAN_TMP%'; $t=Get-Content $p -Raw; $t=$t -replace 'DefaultValue=""https://product-data-collect\.netlify\.app/""','DefaultValue=""https://product-data-collect.netlify.app/index.html""'; Set-Content -Path $p -Value $t -Encoding UTF8"

REM --- 6) Installer le manifest propre dans WEF ---
copy /y "%MAN_TMP%" "%WEF%\manifest.xml" >nul

echo.
echo ============================================
echo  ✔ Installation terminee. Ouvrez Excel.
echo ============================================
echo  (Si une fenetre 'debug' apparait encore,
echo   c'est qu'un outil re-ecrit la valeur 'Developer'.)
echo ============================================
pause
endlocal
