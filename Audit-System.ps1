<#
.SYNOPSIS
    Script d'audit système complet avec génération d'un rapport HTML professionnel et gestion de thèmes.

.DESCRIPTION
    Ce script collecte des informations détaillées sur le système (matériel, réseau, mises à jour)
    et génère un rapport HTML stylisé avec navigation par onglets et sélection de thèmes de couleurs.

.PARAMETER OutputPath
    Chemin de destination pour le rapport HTML (par défaut: Bureau de l'utilisateur)

.PARAMETER Theme
    Thème de couleur à utiliser (Light, Dark, Hacker, Custom)

.EXAMPLE
    .\Audit-System.ps1
    Lance l'audit avec sélection interactive du thème

.EXAMPLE
    .\Audit-System.ps1 -Theme Dark
    Lance l'audit directement avec le thème Dark

.NOTES
    Auteur: Tshienda Cherubin
    Version: 4.0 (Tabs & UI Update)
    Date: $(Get-Date -Format 'yyyy-MM-dd')
    Requiert: PowerShell 5.1 ou supérieur, droits administrateur recommandés
    Licence: MIT
#>

#requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Light", "Dark", "Hacker", "Custom", "Interactive")]
    [string]$Theme = "Interactive"
)

# Configuration des chemins
$ReportName = "System_Audit_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').html"
$FullReportPath = Join-Path $OutputPath $ReportName

# Définition des palettes de couleurs prédéfinies
$Global:Themes = @{
    'Light' = @{
        '--bg-color' = '#f4f4f0'
        '--text-main' = '#111111'
        '--text-muted' = '#888888'
        '--border-color' = '#111111'
        '--accent-color' = '#009966'
        '--card-bg' = '#ffffff'
        '--nav-bg' = '#111111'
        '--nav-text' = '#ffffff'
        '--badge-success-bg' = '#e6f4ea'
        '--badge-success-text' = '#1e8e3e'
        '--badge-success-border' = '#cce8d6'
        '--badge-warning-bg' = '#fef7e0'
        '--badge-warning-text' = '#b06000'
        '--badge-warning-border' = '#fce8b2'
        '--badge-danger-bg' = '#fce8e6'
        '--badge-danger-text' = '#d93025'
        '--badge-danger-border' = '#fad2cf'
    }
    'Dark' = @{
        '--bg-color' = '#121212'
        '--text-main' = '#e8eaed'
        '--text-muted' = '#9aa0a6'
        '--border-color' = '#3c4043'
        '--accent-color' = '#34a853'
        '--card-bg' = '#1e1e1e'
        '--nav-bg' = '#e8eaed'
        '--nav-text' = '#121212'
        '--badge-success-bg' = 'rgba(30, 142, 62, 0.2)'
        '--badge-success-text' = '#81c995'
        '--badge-success-border' = '#1e8e3e'
        '--badge-warning-bg' = 'rgba(249, 171, 0, 0.2)'
        '--badge-warning-text' = '#fde293'
        '--badge-warning-border' = '#f9ab00'
        '--badge-danger-bg' = 'rgba(217, 48, 37, 0.2)'
        '--badge-danger-text' = '#f28b82'
        '--badge-danger-border' = '#d93025'
    }
    'Hacker' = @{
        '--bg-color' = '#0a0a0a'
        '--text-main' = '#00ff00'
        '--text-muted' = '#008800'
        '--border-color' = '#00ff00'
        '--accent-color' = '#00cc00'
        '--card-bg' = '#000000'
        '--nav-bg' = '#00ff00'
        '--nav-text' = '#000000'
        '--badge-success-bg' = '#002200'
        '--badge-success-text' = '#00ff00'
        '--badge-success-border' = '#00ff00'
        '--badge-warning-bg' = '#222200'
        '--badge-warning-text' = '#ffff00'
        '--badge-warning-border' = '#ffff00'
        '--badge-danger-bg' = '#220000'
        '--badge-danger-text' = '#ff0000'
        '--badge-danger-border' = '#ff0000'
    }
}

# Fonction pour écrire des messages formatés
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}

# Fonction pour collecter les informations système
function Get-SystemInfo {
    Write-Log "Collecte des informations système..." -Color "Cyan"
    
    try {
        $computerInfo = Get-ComputerInfo -ErrorAction Stop
        $os = Get-WmiObject Win32_OperatingSystem -ErrorAction Stop
        $bios = Get-WmiObject Win32_BIOS -ErrorAction Stop
        $processor = Get-WmiObject Win32_Processor -ErrorAction Stop
        $memory = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        $disk = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
        $network = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true } -ErrorAction SilentlyContinue
        $hotfixes = Get-WmiObject Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 10 -ErrorAction SilentlyContinue
        
        Write-Log "Informations collectees avec succes" -Color "Green"
        
        return [PSCustomObject]@{
            ComputerName = $computerInfo.CsName
            OS = $computerInfo.WindowsProductName
            OSBuild = $computerInfo.OsBuildNumber
            Manufacturer = $computerInfo.CsManufacturer
            Model = $computerInfo.CsModel
            BIOS = "$($bios.Manufacturer) $($bios.Name)"
            ProcessorName = $processor.Name
            ProcessorCores = $processor.NumberOfCores
            ProcessorLogical = $processor.NumberOfLogicalProcessors
            RAM = "{0:N2} GB" -f ($memory.TotalPhysicalMemory / 1GB)
            Disks = $disk
            Network = $network
            Hotfixes = $hotfixes
            Date = Get-Date -Format "yyyy-MM-dd HH:mm"
            Uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
        }
    }
    catch {
        Write-Log "Erreur lors de la collecte: $($_.Exception.Message)" -Color "Red"
        throw
    }
}

# Fonction pour générer le rapport HTML
function New-HtmlReport {
    param(
        $SystemInfo,
        $ThemeColors
    )
    
    Write-Log "Generation du rapport HTML..." -Color "Cyan"
    
    $html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Audit Systeme - $($SystemInfo.ComputerName)</title>
    <style>
        :root {
            --bg-color: $($ThemeColors['--bg-color']);
            --text-main: $($ThemeColors['--text-main']);
            --text-muted: $($ThemeColors['--text-muted']);
            --border-color: $($ThemeColors['--border-color']);
            --accent-color: $($ThemeColors['--accent-color']);
            --card-bg: $($ThemeColors['--card-bg']);
            --nav-bg: $($ThemeColors['--nav-bg']);
            --nav-text: $($ThemeColors['--nav-text']);
            
            --badge-success-bg: $($ThemeColors['--badge-success-bg']);
            --badge-success-text: $($ThemeColors['--badge-success-text']);
            --badge-success-border: $($ThemeColors['--badge-success-border']);
            
            --badge-warning-bg: $($ThemeColors['--badge-warning-bg']);
            --badge-warning-text: $($ThemeColors['--badge-warning-text']);
            --badge-warning-border: $($ThemeColors['--badge-warning-border']);
            
            --badge-danger-bg: $($ThemeColors['--badge-danger-bg']);
            --badge-danger-text: $($ThemeColors['--badge-danger-text']);
            --badge-danger-border: $($ThemeColors['--badge-danger-border']);

            --font-serif: 'Georgia', serif;
            --font-mono: 'Consolas', 'Courier New', monospace;
            --font-sans: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            background-color: var(--bg-color);
            color: var(--text-main);
            font-family: var(--font-sans);
            line-height: 1.6;
            transition: background-color 0.3s, color 0.3s;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 40px 20px;
        }

        .header-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            padding-bottom: 20px;
            border-bottom: 1px solid var(--border-color);
            margin-bottom: 30px;
        }

        .title-section h1 {
            font-family: var(--font-serif);
            font-style: italic;
            font-size: 3.5rem;
            font-weight: normal;
            margin: 0;
            letter-spacing: -1px;
        }

        .subtitle {
            font-family: var(--font-mono);
            font-size: 0.85rem;
            color: var(--text-muted);
            margin-top: 10px;
        }

        .nav-bar {
            display: flex;
            border-bottom: 1px solid var(--border-color);
            margin-bottom: 40px;
            flex-wrap: wrap;
        }

        .nav-item {
            padding: 15px 30px;
            font-family: var(--font-mono);
            font-size: 0.85rem;
            text-transform: uppercase;
            cursor: pointer;
            letter-spacing: 1px;
            color: var(--text-main);
            transition: all 0.2s ease;
            border: none;
            background: none;
        }

        .nav-item:hover {
            background-color: rgba(128, 128, 128, 0.1);
        }

        .nav-item.active {
            background-color: var(--nav-bg);
            color: var(--nav-text);
        }

        .tab-content {
            display: none;
            animation: fadeIn 0.4s ease;
        }

        .tab-content.active {
            display: block;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .stat-card {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            padding: 25px;
            transition: transform 0.2s;
        }

        .stat-card:hover {
            transform: translateY(-2px);
        }

        .stat-label {
            font-family: var(--font-mono);
            font-size: 0.75rem;
            text-transform: uppercase;
            color: var(--text-muted);
            margin-bottom: 15px;
            letter-spacing: 1px;
        }

        .stat-value {
            font-family: var(--font-serif);
            font-style: italic;
            font-size: 1.8rem;
            word-break: break-word;
        }

        .stat-value.sans {
            font-family: var(--font-mono);
            font-style: normal;
            font-size: 1rem;
        }

        .section-title {
            font-family: var(--font-serif);
            font-style: italic;
            font-size: 2.2rem;
            margin: 30px 0 20px;
            font-weight: normal;
        }

        .table-container {
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            margin-bottom: 40px;
            overflow-x: auto;
            border-radius: 4px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 600px;
        }

        th {
            background-color: var(--nav-bg);
            color: var(--nav-text);
            font-family: var(--font-mono);
            font-size: 0.75rem;
            text-transform: uppercase;
            padding: 15px 20px;
            text-align: left;
            letter-spacing: 1px;
        }

        td {
            padding: 15px 20px;
            border-bottom: 1px solid var(--border-color);
            font-family: var(--font-sans);
            font-size: 0.9rem;
            color: var(--text-main);
        }
        
        td.mono {
            font-family: var(--font-mono);
            font-size: 0.85rem;
        }

        tr:last-child td {
            border-bottom: none;
        }

        .badge {
            padding: 4px 12px;
            border-radius: 12px;
            font-family: var(--font-mono);
            font-size: 0.7rem;
            text-transform: uppercase;
            font-weight: bold;
            display: inline-block;
        }

        .badge-success { 
            background-color: var(--badge-success-bg); 
            color: var(--badge-success-text); 
            border: 1px solid var(--badge-success-border); 
        }
        
        .badge-warning { 
            background-color: var(--badge-warning-bg); 
            color: var(--badge-warning-text); 
            border: 1px solid var(--badge-warning-border); 
        }
        
        .badge-danger { 
            background-color: var(--badge-danger-bg); 
            color: var(--badge-danger-text); 
            border: 1px solid var(--badge-danger-border); 
        }

        .footer-box {
            border: 1px dashed var(--border-color);
            padding: 30px;
            text-align: center;
            font-family: var(--font-mono);
            font-size: 0.85rem;
            color: var(--text-muted);
            background-color: var(--card-bg);
            margin-top: 60px;
        }
        
        .btn-outline {
            background-color: transparent;
            color: var(--text-main);
            border: 1px solid var(--border-color);
            padding: 10px 20px;
            cursor: pointer;
            transition: all 0.2s;
            font-family: var(--font-mono);
            font-size: 0.85rem;
        }

        .btn-outline:hover {
            background-color: rgba(128, 128, 128, 0.1);
        }
        
        .actions {
            display: flex;
            gap: 10px;
        }

        @media print {
            .nav-bar, .actions, .footer-box {
                display: none;
            }
        }

        @media (max-width: 768px) {
            .title-section h1 {
                font-size: 2.5rem;
            }
            
            .nav-item {
                padding: 10px 15px;
                font-size: 0.75rem;
            }
            
            .stat-value {
                font-size: 1.4rem;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-top">
            <div class="title-section">
                <h1>Audit Systeme Complet</h1>
                <div class="subtitle">Rapport genere le $($SystemInfo.Date) par Tshienda Cherubin</div>
            </div>
            <div class="actions">
                <button class="btn-outline" onclick="window.print();" title="Imprimer">Imprimer</button>
            </div>
        </div>

        <!-- Navigation -->
        <div class="nav-bar">
            <button class="nav-item active" onclick="switchTab('tab-overview', this)">Vue d'ensemble</button>
            <button class="nav-item" onclick="switchTab('tab-system', this)">Systeme</button>
            <button class="nav-item" onclick="switchTab('tab-hardware', this)">Materiel & Reseau</button>
            <button class="nav-item" onclick="switchTab('tab-software', this)">Logiciels & Mises a jour</button>
        </div>

        <!-- Onglet 1 : Vue d'ensemble -->
        <div id="tab-overview" class="tab-content active">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-label">Machine</div>
                    <div class="stat-value">$($SystemInfo.ComputerName)</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Systeme d'exploitation</div>
                    <div class="stat-value sans">$($SystemInfo.OS)</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Fabricant</div>
                    <div class="stat-value sans">$($SystemInfo.Manufacturer)</div>
                </div>
            </div>
            
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-label">Processeur</div>
                    <div class="stat-value sans">$($SystemInfo.ProcessorCores) coeurs / $($SystemInfo.ProcessorLogical) threads</div>
                    <div style="margin-top:10px; font-size:0.85rem;">$($SystemInfo.ProcessorName)</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Memoire RAM</div>
                    <div class="stat-value sans">$($SystemInfo.RAM)</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Temps de fonctionnement</div>
                    <div class="stat-value sans">$($SystemInfo.Uptime.Days)j $($SystemInfo.Uptime.Hours)h $($SystemInfo.Uptime.Minutes)m</div>
                </div>
            </div>
        </div>

        <!-- Onglet 2 : Systeme -->
        <div id="tab-system" class="tab-content">
            <h2 class="section-title">Informations systeme</h2>
            <div class="table-container">
                <table>
                    <tr><th>Propriete</th><th>Valeur</th></tr>
                    <tr><td>Nom de l'ordinateur</td><td class="mono">$($SystemInfo.ComputerName)</td></tr>
                    <tr><td>Systeme d'exploitation</td><td>$($SystemInfo.OS)</td></tr>
                    <tr><td>Version du systeme</td><td class="mono">$($SystemInfo.OSBuild)</td></tr>
                    <tr><td>Fabricant</td><td>$($SystemInfo.Manufacturer)</td></tr>
                    <tr><td>Modele</td><td>$($SystemInfo.Model)</td></tr>
                    <tr><td>BIOS</td><td>$($SystemInfo.BIOS)</td></tr>
                    <tr><td>Dernier demarrage</td><td class="mono">$(Get-Date).AddSeconds(-$SystemInfo.Uptime.TotalSeconds) </td></tr>
                </table>
            </div>
        </div>

        <!-- Onglet 3 : Materiel & Reseau -->
        <div id="tab-hardware" class="tab-content">
            <h2 class="section-title">Disques durs</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Lettre</th>
                            <th>Nom</th>
                            <th>Systeme de fichiers</th>
                            <th>Taille totale</th>
                            <th>Espace libre</th>
                            <th>Utilisation</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($disk in $SystemInfo.Disks) {
        $totalSizeGB = [math]::Round($disk.Size / 1GB, 2)
        $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        $usedSpaceGB = $totalSizeGB - $freeSpaceGB
        $usagePercent = if ($totalSizeGB -gt 0) { [math]::Round(($usedSpaceGB / $totalSizeGB) * 100, 2) } else { 0 }
        
        $badgeClass = if ($usagePercent -gt 90) { "badge-danger" } elseif ($usagePercent -gt 75) { "badge-warning" } else { "badge-success" }
        $statusText = if ($usagePercent -gt 90) { "CRITIQUE" } elseif ($usagePercent -gt 75) { "AVERTISSEMENT" } else { "NORMALE" }
        
        $html += @"
                        <tr>
                            <td class="mono"><strong>$($disk.DeviceID)</strong></td>
                            <td>$($disk.VolumeName)</td>
                            <td class="mono">$($disk.FileSystem)</td>
                            <td class="mono">$totalSizeGB GB</td>
                            <td class="mono">$freeSpaceGB GB</td>
                            <td>
                                <span class="badge $badgeClass">$statusText ($usagePercent%)</span>
                            </td>
                        </tr>
"@
    }

    $html += @"
                    </tbody>
                </table>
            </div>

            <h2 class="section-title">Configuration reseau</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Description</th>
                            <th>Adresse IP</th>
                            <th>Masque</th>
                            <th>Passerelle</th>
                            <th>DNS</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($net in $SystemInfo.Network) {
        $ipAddress = $net.IPAddress -join ", "
        $subnetMask = $net.IPSubnet -join ", "
        $gateway = $net.DefaultIPGateway -join ", "
        $dnsServers = $net.DNSServerSearchOrder -join ", "
        
        $html += @"
                        <tr>
                            <td>$($net.Description)</td>
                            <td class="mono">$ipAddress</td>
                            <td class="mono">$subnetMask</td>
                            <td class="mono">$gateway</td>
                            <td class="mono">$dnsServers</td>
                        </tr>
"@
    }

    $html += @"
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Onglet 4 : Logiciels & Mises a jour -->
        <div id="tab-software" class="tab-content">
            <h2 class="section-title">Mises a jour installees</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>ID KB</th>
                            <th>Description</th>
                            <th>Installe le</th>
                            <th>Statut</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    if ($SystemInfo.Hotfixes.Count -gt 0) {
        foreach ($hotfix in $SystemInfo.Hotfixes) {
            $html += @"
                        <tr>
                            <td class="mono"><strong>$($hotfix.HotFixID)</strong></td>
                            <td>$($hotfix.Description)</td>
                            <td class="mono">$($hotfix.InstalledOn)</td>
                            <td><span class="badge badge-success">INSTALLE</span></td>
                        </tr>
"@
        }
    } else {
        $html += @"
                        <tr>
                            <td colspan="4" style="text-align: center;">Aucune mise a jour trouvee</td>
                        </tr>
"@
    }

    $html += @"
                    </tbody>
                </table>
            </div>
        </div>

        <div class="footer-box">
            <p>Rapport genere automatiquement par PowerShell Audit System</p>
            <p>Auteur: Tshienda Cherubin | Version: 4.0 | $(Get-Date -Format 'yyyy')</p>
        </div>
    </div>

    <script>
        function switchTab(tabId, element) {
            var contents = document.getElementsByClassName('tab-content');
            for (var i = 0; i < contents.length; i++) {
                contents[i].classList.remove('active');
            }
            
            var tabs = document.getElementsByClassName('nav-item');
            for (var i = 0; i < tabs.length; i++) {
                tabs[i].classList.remove('active');
            }
            
            document.getElementById(tabId).classList.add('active');
            element.classList.add('active');
        }
    </script>
</body>
</html>
"@

    return $html
}

# Fonction pour la selection interactive du theme
function Select-Theme {
    Write-Host "`nSelection du theme de couleurs pour le rapport :" -ForegroundColor Cyan
    Write-Host "1. Light  (Design clair original)"
    Write-Host "2. Dark   (Mode sombre elegant)"
    Write-Host "3. Hacker (Theme terminal vert/noir)"
    Write-Host "4. Custom (Configuration personnalisee)"
    
    $themeChoice = Read-Host "Votre choix (1-4) [Defaut: 1]"
    
    $SelectedThemeName = "Light"
    $CustomColors = @{}
    
    switch ($themeChoice) {
        '2' { $SelectedThemeName = "Dark" }
        '3' { $SelectedThemeName = "Hacker" }
        '4' { 
            $SelectedThemeName = "Custom" 
            Write-Host "`nConfiguration du theme personnalise (Entree = valeur par defaut) :" -ForegroundColor Yellow
            
            $bg = Read-Host "Couleur de fond (ex: #ffffff)"
            if ($bg) { $CustomColors['--bg-color'] = $bg }
            
            $text = Read-Host "Couleur du texte principal (ex: #000000)"
            if ($text) { $CustomColors['--text-main'] = $text }
            
            $accent = Read-Host "Couleur d'accentuation (ex: #009966)"
            if ($accent) { $CustomColors['--accent-color'] = $accent }
            
            $card = Read-Host "Couleur de fond des cartes (ex: #f9f9f9)"
            if ($card) { $CustomColors['--card-bg'] = $card }
        }
        Default { $SelectedThemeName = "Light" }
    }
    
    return @{
        Name = $SelectedThemeName
        Colors = $CustomColors
    }
}

# Fonction principale
function Start-SystemAudit {
    Write-Host @"
==================================================
      AUDIT SYSTEME COMPLET - GENERATION RAPPORT
==================================================
Auteur: Tshienda Cherubin
Version: 4.0
==================================================
"@ -ForegroundColor Cyan

    # Verification des droits administrateur
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "ATTENTION: Le script n'est pas execute avec les droits administrateur" -ForegroundColor Yellow
        Write-Host "Certaines informations pourraient ne pas etre accessibles" -ForegroundColor Yellow
        $choice = Read-Host "Voulez-vous continuer quand meme? (O/N)"
        if ($choice -ne 'O' -and $choice -ne 'o') {
            Write-Host "Audit annule par l'utilisateur" -ForegroundColor Red
            return
        }
    }

    # Selection du theme
    if ($Theme -eq "Interactive") {
        $themeSelection = Select-Theme
        $SelectedThemeName = $themeSelection.Name
        $CustomColors = $themeSelection.Colors
    } else {
        $SelectedThemeName = $Theme
        $CustomColors = @{}
    }

    # Fusion des couleurs
    $ThemeColors = $Global:Themes['Light'].Clone()
    if ($SelectedThemeName -ne 'Custom') {
        $ThemeColors = $Global:Themes[$SelectedThemeName]
    } else {
        foreach ($key in $CustomColors.Keys) {
            $ThemeColors[$key] = $CustomColors[$key]
        }
    }

    Write-Host "`nDemarrage de l'audit systeme avec le theme [$SelectedThemeName]..." -ForegroundColor Cyan

    try {
        # Creation du repertoire de destination si necessaire
        if (-not (Test-Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        }

        # Collecte des informations
        $systemInfo = Get-SystemInfo
        
        Write-Host "Generation du rapport HTML..." -ForegroundColor Yellow
        
        # Generation du rapport HTML
        $htmlReport = New-HtmlReport -SystemInfo $systemInfo -ThemeColors $ThemeColors
        
        # Sauvegarde du rapport
        $htmlReport | Out-File -FilePath $FullReportPath -Encoding UTF8
        
        Write-Host "Rapport genere avec succes!" -ForegroundColor Green
        Write-Host "Emplacement: $FullReportPath" -ForegroundColor Cyan
        
        # Verification de la taille du fichier
        $fileInfo = Get-Item $FullReportPath
        Write-Host "Taille du rapport: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Gray
        
        # Ouverture automatique du rapport
        $choice = Read-Host "`nVoulez-vous ouvrir le rapport maintenant? (O/N)"
        if ($choice -eq 'O' -or $choice -eq 'o') {
            Start-Process $FullReportPath
        }
    }
    catch {
        Write-Host "ERREUR: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Details: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

# Execution du script
Start-SystemAudit

# Attente de la touche pour quitter (methode robuste)
Write-Host "`nAppuyez sur une touche pour quitter..." -ForegroundColor Gray

try {
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
catch {
    Write-Host "(Appuyez sur Entree pour quitter)" -ForegroundColor DarkGray
    $null = Read-Host
}