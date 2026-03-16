# ------------------------------------------------------------
# INSTALLATION D’UN COMPLÉMENT OFFICE VIA MANIFEST WEF
# Auteur : Diaby Ahme
# ------------------------------------------------------------

Write-Host "Installation du complément Office..." -ForegroundColor Cyan

# 1. Chemin du manifeste à installer
$ManifestPath = "C:\Chemin\vers\ton\manifest.xml"   # <-- mets ton chemin ici

if (-Not (Test-Path $ManifestPath)) {
    Write-Host "❌ Le fichier manifest.xml est introuvable." -ForegroundColor Red
    exit
}

# 2. Dossier WEF (sideload stable, sans debug)
$WefFolder = "$env:LOCALAPPDATA\Microsoft\Office\Wef"

# 3. Création du dossier si manquant
if (-Not (Test-Path $WefFolder)) {
    Write-Host "Création du dossier WEF..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $WefFolder | Out-Null
}

# 4. Copier le manifeste dans WEF
Write-Host "Copie du manifeste dans : $WefFolder" -ForegroundColor Cyan
Copy-Item -Path $ManifestPath -Destination $WefFolder -Force

# 5. Nettoyer potentiels fichiers en cache
Write-Host "Nettoyage du cache de sideload..." -ForegroundColor Yellow
$CacheFolder = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\Cache"
if (Test-Path $CacheFolder) {
    Remove-Item -Recurse -Force $CacheFolder
}

# 6. Redémarrage des applications Office
Write-Host "Fermeture des applications Office..." -ForegroundColor Yellow
Get-Process excel,winword,powerpnt -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host "✅ Installation terminée !" -ForegroundColor Green
Write-Host "➡ Vous pouvez maintenant relancer Excel / Word." -ForegroundColor Green
