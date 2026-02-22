# SyncToDrive.ps1
# Script de synchronisation automatique vers Google Drive
# Surveille les modifications et synchronise vers le dossier Drive spécifié

# Configuration
$SourceFolder = "C:\Users\Philippe\Documents\zebra-printer-monitor"
$RcloneRemote = "gdrive"
$DriveFolderId = "1lGpKCfSENKSGF6lz4NO6v7DcN9lmgRPs"
$LogFile = "$SourceFolder\sync_log.txt"
$RclonePath = "C:\Users\Philippe\AppData\Local\Microsoft\WinGet\Packages\Rclone.Rclone_Microsoft.Winget.Source_8wekyb3d8bbwe\rclone-v1.73.0-windows-amd64\rclone.exe"

# Fichiers Apps Script à synchroniser avec clasp
$AppsScriptFiles = @("Code.gs", "Index.html", "appsscript.json")

# Fichiers/dossiers à exclure de la synchronisation
$ExcludePatterns = @(
    ".git/**",
    "sync_log.txt",
    "printer_status_history.json",
    "*.output"
)

# Fonction pour écrire dans le log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Write-Host $logMessage
    Add-Content -Path $LogFile -Value $logMessage
}

# Fonction pour vérifier si rclone est configuré
function Test-RcloneConfig {
    $remotes = & $RclonePath listremotes 2>$null
    return $remotes -match "^${RcloneRemote}:"
}

# Fonction pour configurer rclone (guide l'utilisateur)
function Initialize-RcloneConfig {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  Configuration de rclone pour Google Drive" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Veuillez suivre les instructions ci-dessous :" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "1. Tapez 'n' pour nouveau remote"
    Write-Host "2. Nom: gdrive"
    Write-Host "3. Storage: tapez 'drive' ou le numero correspondant a Google Drive"
    Write-Host "4. client_id: laissez vide (appuyez sur Entree)"
    Write-Host "5. client_secret: laissez vide (appuyez sur Entree)"
    Write-Host "6. scope: tapez '1' pour full access"
    Write-Host "7. root_folder_id: laissez vide (appuyez sur Entree)"
    Write-Host "8. service_account_file: laissez vide (appuyez sur Entree)"
    Write-Host "9. Edit advanced config: 'n'"
    Write-Host "10. Use auto config: 'y' (un navigateur s'ouvrira)"
    Write-Host "11. Configure as team drive: 'n'"
    Write-Host "12. Confirmez avec 'y', puis 'q' pour quitter"
    Write-Host ""
    Write-Host "Lancement de rclone config..." -ForegroundColor Green
    Write-Host ""

    & $RclonePath config
}

# Fonction de synchronisation vers Google Apps Script (clasp push)
function Sync-ToAppsScript {
    Write-Log "Push vers Google Apps Script..."

    try {
        $originalLocation = Get-Location
        Set-Location $SourceFolder

        $result = & clasp push --force 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Apps Script mis a jour: Code.gs, Index.html"
            Write-Host "  -> Apps Script mis a jour!" -ForegroundColor Green
        } else {
            Write-Log "Erreur clasp push: $result"
            Write-Host "  -> Erreur Apps Script: $result" -ForegroundColor Red
        }

        Set-Location $originalLocation
    }
    catch {
        Write-Log "Exception clasp: $_"
    }
}

# Fonction de synchronisation vers Google Drive
function Sync-ToGoogleDrive {
    Write-Log "Demarrage de la synchronisation..."

    # Construire les arguments d'exclusion
    $excludeArgs = $ExcludePatterns | ForEach-Object { "--exclude"; $_ }

    # Destination avec l'ID du dossier
    $destination = "${RcloneRemote},root_folder_id=${DriveFolderId}:"

    try {
        # Exécuter rclone copy (au lieu de sync pour ne pas supprimer les fichiers existants sur Drive)
        $result = & $RclonePath copy $SourceFolder $destination $excludeArgs --verbose 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Synchronisation reussie!"
        } else {
            Write-Log "Erreur lors de la synchronisation: $result"
        }
    }
    catch {
        Write-Log "Exception: $_"
    }
}

# Mode de fonctionnement
param(
    [switch]$Watch,      # Mode surveillance continue
    [switch]$Once,       # Synchronisation unique
    [switch]$Configure   # Configurer rclone
)

# Vérifier la configuration rclone
if (-not (Test-RcloneConfig)) {
    Write-Host ""
    Write-Host "rclone n'est pas configure pour Google Drive." -ForegroundColor Red
    Write-Host "Lancement de la configuration..." -ForegroundColor Yellow
    Initialize-RcloneConfig

    if (-not (Test-RcloneConfig)) {
        Write-Host "Configuration echouee. Veuillez reessayer." -ForegroundColor Red
        exit 1
    }
}

if ($Configure) {
    Initialize-RcloneConfig
    exit 0
}

if ($Once -or (-not $Watch)) {
    # Synchronisation unique
    Sync-ToGoogleDrive
    if (-not $Watch) {
        exit 0
    }
}

if ($Watch) {
    Write-Log "Mode surveillance active. Ctrl+C pour arreter."
    Write-Host ""
    Write-Host "Surveillance du dossier: $SourceFolder" -ForegroundColor Green
    Write-Host "Destination Google Drive: Dossier ID $DriveFolderId" -ForegroundColor Green
    Write-Host ""

    # Créer le FileSystemWatcher
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $SourceFolder
    $watcher.Filter = "*.*"
    $watcher.IncludeSubdirectories = $true
    $watcher.EnableRaisingEvents = $true

    # Variables pour le debounce (éviter les synchronisations multiples)
    $script:lastSync = [DateTime]::MinValue
    $script:syncDelay = 5  # Secondes d'attente après une modification

    # Action à exécuter lors d'une modification
    $action = {
        $path = $Event.SourceEventArgs.FullPath
        $changeType = $Event.SourceEventArgs.ChangeType
        $name = $Event.SourceEventArgs.Name

        # Ignorer certains fichiers
        if ($name -match "\.git|sync_log\.txt|printer_status_history\.json") {
            return
        }

        $now = [DateTime]::Now
        $timeSinceLastSync = ($now - $script:lastSync).TotalSeconds

        if ($timeSinceLastSync -ge $script:syncDelay) {
            Write-Host ""
            Write-Host "[$($now.ToString('HH:mm:ss'))] Modification detectee: $name ($changeType)" -ForegroundColor Yellow
            $script:lastSync = $now

            # Attendre un peu pour regrouper les modifications
            Start-Sleep -Seconds 2

            # Synchroniser vers Google Drive
            Sync-ToGoogleDrive

            # Si c'est un fichier Apps Script, pousser aussi vers Google Apps Script
            if ($name -match "^(Code\.gs|Index\.html|appsscript\.json)$") {
                Sync-ToAppsScript
            }
        }
    }

    # Enregistrer les événements
    Register-ObjectEvent $watcher "Created" -Action $action | Out-Null
    Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null
    Register-ObjectEvent $watcher "Deleted" -Action $action | Out-Null
    Register-ObjectEvent $watcher "Renamed" -Action $action | Out-Null

    # Faire une première synchronisation
    Sync-ToGoogleDrive

    Write-Host ""
    Write-Host "En attente de modifications... (Ctrl+C pour arreter)" -ForegroundColor Cyan

    # Boucle infinie
    try {
        while ($true) {
            Start-Sleep -Seconds 1
        }
    }
    finally {
        # Nettoyage
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
        Get-EventSubscriber | Unregister-Event
        Write-Log "Surveillance arretee."
    }
}
