# HP Bureautique Monitor - Script de Test de Connexion
# Ce script récupère automatiquement la liste des imprimantes depuis Google Apps Script
# puis teste leur disponibilité via TCP (port 9100) et envoie les statuts
# Note: Utilise TCP au lieu de ICMP (ping) car les pings sont souvent bloqués en entreprise

# Configuration - URL de votre Web App déployée
$WebAppUrl = "https://script.google.com/macros/s/AKfycbznnKwXXvEdLohEAf2_a3iMYVrLwItRaiJKS0Vo5P40wW4MWbOEdxAMpwEMGBcYewgI/exec"

# Charger la clé API depuis le fichier secrets (non versionné)
$SecretsFile = "$PSScriptRoot\secrets.ps1"
if (Test-Path $SecretsFile) {
    . $SecretsFile
} else {
    # Log pas encore initialisé, écrire directement
    $errorMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Fichier secrets.ps1 introuvable!"
    Add-Content -Path "$PSScriptRoot\ping_log.txt" -Value $errorMsg -Encoding UTF8
    Write-Host "ERREUR: Fichier secrets.ps1 introuvable!" -ForegroundColor Red
    Write-Host "Créez le fichier secrets.ps1 avec la variable `$ApiSecretKey" -ForegroundColor Yellow
    exit 1
}

# Fichier pour stocker les statuts précédents (pour détecter les changements)
$StatusHistoryFile = "$PSScriptRoot\printer_status_history.json"

# Fichier résumé des derniers résultats
$SummaryFile = "$PSScriptRoot\ping_summary.txt"

# Fichier de log pour le debugging
$LogFile = "$PSScriptRoot\ping_log.txt"

# Fonction pour écrire dans le log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"  # INFO, WARNING, ERROR
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    # Écrire dans le fichier
    Add-Content -Path $LogFile -Value $logLine -Encoding UTF8

    # Afficher à l'écran avec couleur selon le niveau
    switch ($Level) {
        "ERROR"   { Write-Host $logLine -ForegroundColor Red }
        "WARNING" { Write-Host $logLine -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
        default   { Write-Host $logLine -ForegroundColor Gray }
    }
}

# Initialiser le fichier de log (nouveau fichier à chaque exécution, garde les 5 dernières exécutions)
function Initialize-Log {
    $separator = "=" * 80
    $header = @"

$separator
  NOUVELLE EXECUTION - $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  PC: $env:COMPUTERNAME | User: $env:USERNAME
$separator

"@

    # Si le fichier dépasse 100KB, le tronquer (garder les dernières lignes)
    if ((Test-Path $LogFile) -and ((Get-Item $LogFile).Length -gt 100KB)) {
        $content = Get-Content $LogFile -Tail 500
        $content | Set-Content $LogFile -Encoding UTF8
        Add-Content -Path $LogFile -Value "`n[LOG TRONQUÉ - Fichier trop volumineux]`n" -Encoding UTF8
    }

    Add-Content -Path $LogFile -Value $header -Encoding UTF8
}

# Initialiser le log au démarrage
Initialize-Log

# Fonction pour récupérer la liste des imprimantes depuis Google Apps Script
function Get-PrintersFromServer {
    Write-Log "Connexion à l'API Google Apps Script..." "INFO"
    Write-Log "URL: $WebAppUrl" "INFO"

    try {
        $body = @{
            action = "getPrinters"
            apiKey = $ApiSecretKey
        } | ConvertTo-Json
        $response = Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $body -ContentType "application/json"

        if ($response.success) {
            Write-Log "Liste des imprimantes récupérée: $($response.printers.Count) imprimante(s)" "SUCCESS"
            return $response.printers
        } else {
            Write-Log "Erreur serveur: $($response.error)" "ERROR"
            return $null
        }
    }
    catch {
        Write-Log "Erreur lors de la récupération des imprimantes: $_" "ERROR"
        Write-Log "Exception détaillée: $($_.Exception.Message)" "ERROR"
        if ($_.Exception.Response) {
            Write-Log "Code HTTP: $($_.Exception.Response.StatusCode.value__)" "ERROR"
        }
        return $null
    }
}

# Fonction pour tester la connexion d'une imprimante via TCP (port 9100)
# Utilise TCP au lieu de ICMP car les pings sont souvent bloqués en entreprise
function Test-PrinterConnection {
    param (
        [string]$IPAddress,
        [int]$Port = 9100,
        [int]$TimeoutMs = 2000
    )

    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $connect = $tcpClient.BeginConnect($IPAddress, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        $stopwatch.Stop()

        if ($wait -and $tcpClient.Connected) {
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            return @{
                Status = "online"
                ResponseTime = $stopwatch.ElapsedMilliseconds
            }
        } else {
            $tcpClient.Close()
            return @{
                Status = "offline"
                ResponseTime = 0
            }
        }
    }
    catch {
        return @{
            Status = "offline"
            ResponseTime = 0
        }
    }
}

# Fonction pour charger l'historique des statuts
function Get-StatusHistory {
    if (Test-Path $StatusHistoryFile) {
        try {
            $content = Get-Content $StatusHistoryFile -Raw
            return $content | ConvertFrom-Json -AsHashtable
        }
        catch {
            return @{}
        }
    }
    return @{}
}

# Fonction pour sauvegarder l'historique des statuts
function Save-StatusHistory {
    param ([hashtable]$History)
    $History | ConvertTo-Json | Set-Content $StatusHistoryFile
}

# Fonction pour détecter les changements de statut
function Get-StatusChanges {
    param (
        [array]$CurrentStatuses,
        [hashtable]$PreviousStatuses
    )

    $changes = @()

    foreach ($status in $CurrentStatuses) {
        $previousStatus = $PreviousStatuses[$status.ip]

        if ($previousStatus -and $previousStatus -ne $status.status) {
            $changes += @{
                ip = $status.ip
                previousStatus = $previousStatus
                newStatus = $status.status
            }
        }
    }

    return $changes
}

# Fonction pour créer le fichier résumé
function Write-SummaryFile {
    param (
        [array]$Statuses,
        [array]$PrinterList,
        [int]$OnlineCount,
        [int]$OfflineCount
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $separator = "=" * 60

    $summary = @"
$separator
  ZEBRA PRINTER MONITOR - RAPPORT DE PING
  $timestamp
$separator

RESUME:
  Total imprimantes: $($Statuses.Count)
  En ligne:          $OnlineCount
  Hors ligne:        $OfflineCount

$separator
DETAILS PAR IMPRIMANTE:
$separator

"@

    # Grouper par statut
    $onlinePrinters = @()
    $offlinePrinters = @()

    foreach ($status in $Statuses) {
        $printerName = ($PrinterList | Where-Object { $_.ip -eq $status.ip }).name
        $entry = "  - $printerName ($($status.ip))"
        if ($status.responseTime -gt 0) {
            $entry += " - $($status.responseTime)ms"
        }

        if ($status.status -eq "online") {
            $onlinePrinters += $entry
        } else {
            $offlinePrinters += $entry
        }
    }

    $summary += "EN LIGNE ($OnlineCount):`n"
    if ($onlinePrinters.Count -gt 0) {
        $summary += ($onlinePrinters -join "`n") + "`n"
    } else {
        $summary += "  (aucune)`n"
    }

    $summary += "`nHORS LIGNE ($OfflineCount):`n"
    if ($offlinePrinters.Count -gt 0) {
        $summary += ($offlinePrinters -join "`n") + "`n"
    } else {
        $summary += "  (aucune)`n"
    }

    $summary += "`n$separator`n"

    # Écrire le fichier
    $summary | Set-Content -Path $SummaryFile -Encoding UTF8
    Write-Host "Fichier résumé créé: $SummaryFile" -ForegroundColor Cyan
}

# Fonction pour envoyer les statuts au Google Apps Script
function Send-StatusToGoogleApps {
    param (
        [array]$Statuses,
        [array]$StatusChanges
    )

    Write-Log "Envoi de $($Statuses.Count) statuts à l'API..." "INFO"

    $body = @{
        action = "updateStatus"
        apiKey = $ApiSecretKey
        statuses = $Statuses
        statusChanges = $StatusChanges
    } | ConvertTo-Json -Depth 3

    try {
        $response = Invoke-RestMethod -Uri $WebAppUrl -Method Post -Body $body -ContentType "application/json"
        Write-Log "Statuts envoyés avec succès" "SUCCESS"

        # Afficher les détails de la réponse pour diagnostic
        if ($response.error) {
            Write-Log "Erreur serveur: $($response.error)" "ERROR"
        } elseif ($response.success) {
            Write-Log "Mises à jour réussies: $($response.updated)" "SUCCESS"
            if ($response.errors -gt 0) {
                Write-Log "Erreurs lors de la mise à jour: $($response.errors)" "ERROR"
                $response.results | Where-Object { $_.success -eq $false } | ForEach-Object {
                    Write-Log "  - $($_.ip): $($_.error)" "ERROR"
                }
            }
        }

        return $response
    }
    catch {
        Write-Log "Erreur lors de l'envoi des statuts: $_" "ERROR"
        Write-Log "Exception détaillée: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Script principal
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  HP Bureautique Monitor - Ping Script" -ForegroundColor Cyan
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Récupérer la liste des imprimantes depuis le serveur
Write-Host "Récupération de la liste des imprimantes..." -ForegroundColor Yellow
$Printers = Get-PrintersFromServer

if (-not $Printers -or $Printers.Count -eq 0) {
    Write-Host "Aucune imprimante trouvée ou erreur de connexion." -ForegroundColor Red
    Write-Host "Terminé." -ForegroundColor Cyan
    exit 1
}

Write-Host ""
$statuses = @()

# Filtrer les imprimantes sans IP valide
$ValidPrinters = $Printers | Where-Object { $_.ip -and $_.ip -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' }
$SkippedCount = $Printers.Count - $ValidPrinters.Count

if ($SkippedCount -gt 0) {
    Write-Host "Note: $SkippedCount imprimante(s) ignorée(s) (IP manquante ou invalide)" -ForegroundColor Yellow
    Write-Host ""
}

foreach ($printer in $ValidPrinters) {
    Write-Host "Test $($printer.name) ($($printer.ip):9100)... " -NoNewline

    $result = Test-PrinterConnection -IPAddress $printer.ip

    if ($result.Status -eq "online") {
        Write-Host "EN LIGNE ($($result.ResponseTime)ms)" -ForegroundColor Green
    } else {
        Write-Host "HORS LIGNE" -ForegroundColor Red
    }

    $statuses += @{
        ip = $printer.ip
        status = $result.Status
        responseTime = $result.ResponseTime
    }
}

Write-Host ""
Write-Host "Résumé:" -ForegroundColor Yellow
$online = ($statuses | Where-Object { $_.status -eq "online" }).Count
$offline = ($statuses | Where-Object { $_.status -eq "offline" }).Count
Write-Host "  En ligne: $online" -ForegroundColor Green
Write-Host "  Hors ligne: $offline" -ForegroundColor Red

# Charger l'historique et détecter les changements
$previousStatuses = Get-StatusHistory
$statusChanges = Get-StatusChanges -CurrentStatuses $statuses -PreviousStatuses $previousStatuses

# Afficher les changements détectés
if ($statusChanges.Count -gt 0) {
    Write-Host ""
    Write-Host "Changements détectés:" -ForegroundColor Magenta
    foreach ($change in $statusChanges) {
        $printerName = ($Printers | Where-Object { $_.ip -eq $change.ip }).name
        if ($change.newStatus -eq "offline") {
            Write-Host "  [ALERTE] $printerName : $($change.previousStatus) -> $($change.newStatus)" -ForegroundColor Red
        } else {
            Write-Host "  [OK] $printerName : $($change.previousStatus) -> $($change.newStatus)" -ForegroundColor Green
        }
    }
}

# Sauvegarder l'historique actuel
$newHistory = @{}
foreach ($status in $statuses) {
    $newHistory[$status.ip] = $status.status
}
Save-StatusHistory -History $newHistory

# Envoyer les statuts
Write-Host ""
Write-Host "Envoi des statuts à Google Apps Script..." -ForegroundColor Yellow
$response = Send-StatusToGoogleApps -Statuses $statuses -StatusChanges $statusChanges

if ($response -and $response.alertsSent -gt 0) {
    Write-Host "Alertes envoyées: $($response.alertsSent)" -ForegroundColor Magenta
}

# Créer le fichier résumé
Write-SummaryFile -Statuses $statuses -PrinterList $Printers -OnlineCount $online -OfflineCount $offline

Write-Host ""
Write-Host "Terminé." -ForegroundColor Cyan

# Pour créer une tâche planifiée (exécuter en tant qu'administrateur):
# schtasks /create /tn "HPBureautiqueMonitor" /tr "powershell.exe -ExecutionPolicy Bypass -File C:\chemin\vers\PingPrinters.ps1" /sc minute /mo 5
