
# IT Inventory Dashboard
# Inventaire local Windows - PowerShell 5.1+
# Les rapports sont crees dans le meme dossier que ce script.

$ErrorActionPreference = 'Stop'
$Dossier = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Dossier)) { $Dossier = (Get-Location).Path }

function Get-SafeValue {
    param($Value, [string]$Fallback = 'Indisponible')
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $Fallback }
    return [string]$Value
}

function ConvertTo-HtmlSafe {
    param($Value)
    return [System.Net.WebUtility]::HtmlEncode((Get-SafeValue $Value))
}

function Get-StatusClass {
    param([string]$Status)
    switch ($Status) {
        'Actif'   { return 'ok' }
        'Active'  { return 'ok' }
        'Chiffre' { return 'ok' }
        'Desactive' { return 'bad' }
        'Inactif' { return 'bad' }
        'Non chiffre' { return 'warn' }
        default   { return 'neutral' }
    }
}

Write-Host 'Collecte des informations du PC...' -ForegroundColor Cyan

try {
    $OS = Get-CimInstance -ClassName Win32_OperatingSystem
    $Systeme = Get-CimInstance -ClassName Win32_ComputerSystem
    $CPUInfo = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
    $Disque = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' |
        Where-Object { $_.DeviceID -eq $env:SystemDrive } |
        Select-Object -First 1

    $Adaptateur = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' |
        Where-Object { $_.IPAddress -and @($_.IPAddress | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }).Count -gt 0 } |
        Select-Object -First 1

    $NomPC = $env:COMPUTERNAME
    $Windows = Get-SafeValue $OS.Caption
    $VersionWindows = Get-SafeValue $OS.Version
    $Processeur = Get-SafeValue $CPUInfo.Name
    $RAM = if ($Systeme.TotalPhysicalMemory) { '{0:N1} Go' -f ($Systeme.TotalPhysicalMemory / 1GB) } else { 'Indisponible' }

    $IPv4 = 'Indisponible'
    $MAC = 'Indisponible'
    if ($Adaptateur) {
        $IPv4Selection = @($Adaptateur.IPAddress | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }) | Select-Object -First 1
        $IPv4 = Get-SafeValue $IPv4Selection
        $MAC = Get-SafeValue $Adaptateur.MACAddress
    }

    $CapaciteGo = 0
    $LibreGo = 0
    $UtilisePourcent = 0
    if ($Disque -and $Disque.Size -gt 0) {
        $CapaciteGo = [math]::Round($Disque.Size / 1GB, 1)
        $LibreGo = [math]::Round($Disque.FreeSpace / 1GB, 1)
        $UtilisePourcent = [math]::Round((($Disque.Size - $Disque.FreeSpace) / $Disque.Size) * 100, 0)
    }

    # Firewall : les trois profils sont verifies individuellement.
    $PareFeuDomain = 'Indisponible'
    $PareFeuPrivate = 'Indisponible'
    $PareFeuPublic = 'Indisponible'
    try {
        $Profils = @(Get-NetFirewallProfile -ErrorAction Stop)
        foreach ($Profil in $Profils) {
            $EtatProfil = if ($Profil.Enabled) { 'Actif' } else { 'Inactif' }
            switch ($Profil.Name) {
                'Domain'  { $PareFeuDomain = $EtatProfil }
                'Private' { $PareFeuPrivate = $EtatProfil }
                'Public'  { $PareFeuPublic = $EtatProfil }
            }
        }
    } catch {
        Write-Warning 'Informations du pare-feu indisponibles.'
    }

    # Defender et BitLocker peuvent etre absents selon l'edition de Windows.
    $Defender = 'Indisponible'
    try {
        $EtatDefender = Get-MpComputerStatus -ErrorAction Stop
        $Defender = if ($EtatDefender.RealTimeProtectionEnabled) { 'Actif' } else { 'Inactif' }
    } catch {
        Write-Warning 'Informations Microsoft Defender indisponibles.'
    }

    $BitLocker = 'Indisponible'
    try {
        $Volume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ($Volume.VolumeStatus -eq 'FullyEncrypted') {
            $BitLocker = 'Chiffre'
        } elseif ($Volume.VolumeStatus -eq 'FullyDecrypted') {
            $BitLocker = 'Non chiffre'
        } else {
            $BitLocker = Get-SafeValue $Volume.VolumeStatus
        }
    } catch {
        Write-Warning 'Informations BitLocker indisponibles.'
    }

    $DerniereMaj = 'Indisponible'
    try {
        $Correctif = Get-HotFix -ErrorAction Stop |
            Where-Object { $_.InstalledOn } |
            Sort-Object InstalledOn -Descending |
            Select-Object -First 1
        if ($Correctif) {
            $DerniereMaj = '{0} ({1})' -f $Correctif.HotFixID, $Correctif.InstalledOn.ToString('dd/MM/yyyy')
        }
    } catch {
        Write-Warning 'Historique des correctifs indisponible.'
    }

    $DateInventaire = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'

    $Inventaire = [PSCustomObject]@{
        DateInventaire       = $DateInventaire
        NomPC                = $NomPC
        Windows              = $Windows
        VersionWindows       = $VersionWindows
        Processeur           = $Processeur
        RAM                  = $RAM
        AdresseIPv4          = $IPv4
        AdresseMAC           = $MAC
        DisqueCapaciteGo     = $CapaciteGo
        DisqueLibreGo        = $LibreGo
        DisqueUtilisePct     = $UtilisePourcent
        PareFeuDomaine       = $PareFeuDomain
        PareFeuPrive         = $PareFeuPrivate
        PareFeuPublic        = $PareFeuPublic
        MicrosoftDefender    = $Defender
        BitLocker            = $BitLocker
        DerniereMiseAJour    = $DerniereMaj
    }

    $CheminCSV = Join-Path $Dossier 'inventaire.csv'
    $CheminTXT = Join-Path $Dossier 'inventaire.txt'
    $CheminHTML = Join-Path $Dossier 'inventaire.html'

    $Inventaire | Export-Csv -Path $CheminCSV -NoTypeInformation -Delimiter ';' -Encoding UTF8

    @(
        'IT INVENTORY DASHBOARD'
        "Date : $DateInventaire"
        '========================================'
        "Nom du PC            : $NomPC"
        "Windows              : $Windows"
        "Version Windows      : $VersionWindows"
        "Processeur           : $Processeur"
        "RAM                  : $RAM"
        "IPv4                 : $IPv4"
        "Adresse MAC          : $MAC"
        "Disque total         : $CapaciteGo Go"
        "Disque libre         : $LibreGo Go"
        "Disque utilise       : $UtilisePourcent %"
        "Pare-feu domaine     : $PareFeuDomain"
        "Pare-feu prive       : $PareFeuPrivate"
        "Pare-feu public      : $PareFeuPublic"
        "Microsoft Defender   : $Defender"
        "BitLocker            : $BitLocker"
        "Dernier correctif    : $DerniereMaj"
    ) | Set-Content -Path $CheminTXT -Encoding UTF8

    $NomPCHtml = ConvertTo-HtmlSafe $NomPC
    $WindowsHtml = ConvertTo-HtmlSafe $Windows
    $VersionHtml = ConvertTo-HtmlSafe $VersionWindows
    $CPUHtml = ConvertTo-HtmlSafe $Processeur
    $RAMHtml = ConvertTo-HtmlSafe $RAM
    $IPv4Html = ConvertTo-HtmlSafe $IPv4
    $MACHtml = ConvertTo-HtmlSafe $MAC
    $DateHtml = ConvertTo-HtmlSafe $DateInventaire
    $DefenderHtml = ConvertTo-HtmlSafe $Defender
    $BitLockerHtml = ConvertTo-HtmlSafe $BitLocker
    $MajHtml = ConvertTo-HtmlSafe $DerniereMaj
    $PareFeuDomainHtml = ConvertTo-HtmlSafe $PareFeuDomain
    $PareFeuPrivateHtml = ConvertTo-HtmlSafe $PareFeuPrivate
    $PareFeuPublicHtml = ConvertTo-HtmlSafe $PareFeuPublic
    $DefenderClasse = Get-StatusClass $Defender
    $BitLockerClasse = Get-StatusClass $BitLocker
    $DomaineClasse = Get-StatusClass $PareFeuDomain
    $PriveClasse = Get-StatusClass $PareFeuPrivate
    $PublicClasse = Get-StatusClass $PareFeuPublic

    $Html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>IT Inventory Dashboard - $NomPCHtml</title>
  <style>
    :root { color-scheme: dark; --bg:#0b1220; --card:#162235; --text:#edf3fc; --muted:#a9b7cb; --border:#2b3b53; --accent:#54b4ff; }
    * { box-sizing: border-box; }
    body { margin:0; padding:34px 20px; font-family:Segoe UI,Arial,sans-serif; background:var(--bg); color:var(--text); }
    .wrap { max-width:1140px; margin:auto; }
    h1 { margin:0 0 7px; font-size:clamp(27px,4vw,38px); }
    .subtitle { color:var(--muted); margin:0 0 29px; }
    .grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(310px,1fr)); gap:18px; }
    .card { background:var(--card); border:1px solid var(--border); border-radius:16px; padding:23px; box-shadow:0 9px 23px rgba(0,0,0,.1); }
    h2 { color:var(--accent); margin:0 0 20px; font-size:19px; }
    .line { display:flex; justify-content:space-between; align-items:flex-start; gap:16px; padding:10px 0; border-bottom:1px solid var(--border); }
    .line:last-child { border-bottom:0; }
    .key { color:var(--muted); flex-shrink:0; }
    .value { text-align:right; overflow-wrap:anywhere; }
    .badge { display:inline-block; padding:4px 10px; border-radius:999px; font-weight:650; font-size:13px; }
    .ok { background:#123e34; color:#75e6b2; }
    .bad { background:#4c2028; color:#ff9ca8; }
    .warn { background:#4a381a; color:#ffcf78; }
    .neutral { background:#344159; color:#ced6e4; }
    .meter { width:100%; height:13px; border-radius:20px; overflow:hidden; background:#2b3b53; margin:15px 0 8px; }
    .fill { width:$UtilisePourcent%; height:100%; border-radius:20px; background:linear-gradient(90deg,#54b4ff,#6bdfcc); }
    .help { color:var(--muted); font-size:13px; margin-top:9px; }
    footer { text-align:center; margin-top:28px; color:var(--muted); font-size:13px; }
  </style>
</head>
<body>
  <main class="wrap">
    <h1>IT Inventory Dashboard</h1>
    <p class="subtitle">Inventaire du poste <strong>$NomPCHtml</strong> - $DateHtml</p>
    <div class="grid">
      <section class="card">
        <h2>Systeme</h2>
        <div class="line"><span class="key">Nom du PC</span><span class="value">$NomPCHtml</span></div>
        <div class="line"><span class="key">Windows</span><span class="value">$WindowsHtml</span></div>
        <div class="line"><span class="key">Version</span><span class="value">$VersionHtml</span></div>
      </section>
      <section class="card">
        <h2>Materiel</h2>
        <div class="line"><span class="key">Processeur</span><span class="value">$CPUHtml</span></div>
        <div class="line"><span class="key">RAM</span><span class="value">$RAMHtml</span></div>
      </section>
      <section class="card">
        <h2>Reseau</h2>
        <div class="line"><span class="key">Adresse IPv4</span><span class="value">$IPv4Html</span></div>
        <div class="line"><span class="key">Adresse MAC</span><span class="value">$MACHtml</span></div>
      </section>
      <section class="card">
        <h2>Stockage ($env:SystemDrive)</h2>
        <div class="line"><span class="key">Capacite totale</span><span class="value">$CapaciteGo Go</span></div>
        <div class="line"><span class="key">Espace libre</span><span class="value">$LibreGo Go</span></div>
        <div class="meter"><div class="fill"></div></div>
        <div class="help">$UtilisePourcent % de l'espace utilise</div>
      </section>
      <section class="card">
        <h2>Securite Windows</h2>
        <div class="line"><span class="key">Defender (temps reel)</span><span class="badge $DefenderClasse">$DefenderHtml</span></div>
        <div class="line"><span class="key">BitLocker</span><span class="badge $BitLockerClasse">$BitLockerHtml</span></div>
        <div class="line"><span class="key">Pare-feu domaine</span><span class="badge $DomaineClasse">$PareFeuDomainHtml</span></div>
        <div class="line"><span class="key">Pare-feu prive</span><span class="badge $PriveClasse">$PareFeuPrivateHtml</span></div>
        <div class="line"><span class="key">Pare-feu public</span><span class="badge $PublicClasse">$PareFeuPublicHtml</span></div>
      </section>
      <section class="card">
        <h2>Maintenance</h2>
        <div class="line"><span class="key">Dernier correctif connu</span><span class="value">$MajHtml</span></div>
        <div class="help">Source : Get-HotFix. Certaines mises a jour peuvent ne pas apparaitre.</div>
      </section>
    </div>
    <footer>Genere localement par IT Inventory Dashboard - Aucun envoi de donnees</footer>
  </main>
</body>
</html>
"@

    $Html | Set-Content -Path $CheminHTML -Encoding UTF8

    Write-Host ''
    Write-Host 'Inventaire termine !' -ForegroundColor Green
    Write-Host "CSV  : $CheminCSV"
    Write-Host "TXT  : $CheminTXT"
    Write-Host "HTML : $CheminHTML"
    Write-Host ''
    Write-Host 'Ouverture du tableau de bord...' -ForegroundColor Cyan
    Start-Process -FilePath $CheminHTML
} catch {
    Write-Error "Impossible de terminer l'inventaire : $($_.Exception.Message)"
}
