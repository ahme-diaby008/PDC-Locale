# ------------------------------------------------------------
# INSTALLATION AUTOMATIQUE DU COMPLÉMENT OFFICE PDC
# ------------------------------------------------------------

Write-Host "Installation du complément Office PDC..." -ForegroundColor Cyan

# 1. URL RAW GitHub du manifest
$ManifestUrl = "https://raw.githubusercontent.com/ahme-diaby008/PDC-Locale/refs/heads/main/manifest.xml"

# 2. Fichier temporaire
$TempManifest = "$env:TEMP\manifest.xml"

# 3. Télécharger le manifest
Write-Host "Téléchargement du manifest depuis GitHub..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $ManifestUrl -OutFile $TempManifest -UseBasicParsing

# 4. Dossier WEF pour sideload stable
$WefFolder = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"

# Créer le dossier si nécessaire
if (-Not (Test-Path $WefFolder)) {
    Write-Host "Création du dossier WEF..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $WefFolder | Out-Null
}

# 5. Installer le manifest dans WEF
Write-Host "Installation du manifest dans : $WefFolder" -ForegroundColor Cyan
Copy-Item -Path $TempManifest -Destination "$WefFolder\manifest.xml" -Force

# 6. Purger le cache Office
$CacheFolder = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache"
if (Test-Path $CacheFolder) {
    Write-Host "Nettoyage du cache Office..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force $CacheFolder
}

# 7. Redémarrer les applications Office
Write-Host "Fermeture des applications Office..." -ForegroundColor Yellow
Get-Process excel, winword, powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host "✅ Installation terminée !"
Write-Host "➡️ Vous pouvez maintenant relancer Excel." -ForegroundColor Green
