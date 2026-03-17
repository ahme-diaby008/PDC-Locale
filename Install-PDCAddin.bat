@echo off
setlocal

echo ============================================
echo  Installation PDC (per-user, sans admin)
echo  Mode: catalogue local + purge cache WEF
echo ============================================

REM --- Chemins ---
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
set "WEF_CACHE=%WEF%\Cache"
set "CATALOG_DIR=%LOCALAPPDATA%\OfficeAddins\Manifests"
set "REG_WEF=HKCU\Software\Microsoft\Office\16.0\WEF"
set "REG_CATALOG=%REG_WEF%\TrustedCatalogs"
set "GUID={01234567-89ab-cedf-0123-456789abcedf}"

REM --- Fermer Office ---
for %%A in (excel.exe winword.exe powerpnt.exe) do taskkill /f /im %%A >nul 2>&1

REM --- Nettoyage options dev (optionnel) ---
reg delete "%REG_WEF%" /v Developer /f           >nul 2>&1
reg delete "%REG_WEF%" /v DeveloperTools /f      >nul 2>&1
reg delete "%REG_WEF%" /v RuntimeDebugging /f    >nul 2>&1
reg delete "%REG_WEF%" /v RuntimeLogging /f      >nul 2>&1
reg delete "%REG_WEF%" /v UseSharedRuntimeDebugging /f >nul 2>&1

REM --- Purge cache WEF (chemin officiel) ---
if exist "%WEF_CACHE%" rmdir /s /q "%WEF_CACHE%"
if not exist "%WEF%" mkdir "%WEF%"

REM --- Dossier catalogue local ---
if not exist "%CATALOG_DIR%" mkdir "%CATALOG_DIR%"

REM --- 1) Telechargement du manifeste ---
set "MAN_URL=https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/manifest.xml"
set "MAN_TMP=%TEMP%\manifest.xml"
echo Telechargement du manifeste...
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MAN_URL%' -OutFile '%MAN_TMP%' -UseBasicParsing" ^
  || (echo Echec du telechargement. Verifiez l'acces Internet/proxy.& pause & exit /b 1)

REM --- 2) (Optionnel) Normaliser SourceLocation via XML, pas regex ---
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p='%MAN_TMP%'; $xml=[xml](Get-Content $p -Raw);" ^
  "$ns=New-Object System.Xml.XmlNamespaceManager($xml.NameTable);" ^
  "$ns.AddNamespace('m','http://schemas.microsoft.com/office/appforoffice/1.1');" ^
  "$node=$xml.SelectSingleNode('//m:SourceLocation', $ns);" ^
  "if($node -and $node.DefaultValue -match '^https://product-data-collect\.netlify\.app/?$'){ $node.DefaultValue='https://product-data-collect.netlify.app/index.html'; $xml.Save($p) }"

REM --- 3) Copier le manifeste dans le catalogue local ---
copy /y "%MAN_TMP%" "%CATALOG_DIR%\PDC.manifest.xml" >nul

REM --- 4) Declarer le catalogue de confiance (HKCU) ---
reg add "%REG_CATALOG%\%GUID%" /v Id    /t REG_SZ    /d "%GUID%" /f
reg add "%REG_CATALOG%\%GUID%" /v Url   /t REG_SZ    /d "%CATALOG_DIR%" /f
reg add "%REG_CATALOG%\%GUID%" /v Flags /t REG_DWORD /d 1 /f

echo.
echo ============================================
echo  ✔ Installation terminee (per-user)
echo ============================================
echo  Ouvrez Excel/Word/PowerPoint :
echo   - Accueil > Modules complementaires > Avance > SHARED FOLDER
echo   - Selectionnez "PDC" puis "Ajouter"
echo  (Si le bouton tarde a apparaitre, relancez l'app.
echo   Au besoin, videz le cache : %LOCALAPPDATA%\Microsoft\Office\16.0\Wef)
echo ============================================
pause
endlocal
