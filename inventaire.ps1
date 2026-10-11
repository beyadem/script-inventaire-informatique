# Projet : SCRIPT D'INVENTAIRE INFORMATIQUE
# Auteur : beyadem
# ==========================================

# 1. Creation du dossier Projet
$Dossier = "C:\Projet"
if (-not (Test-Path $Dossier)) {
    New-Item -Path $Dossier -ItemType Directory | Out-Null
}

# 2. Informations sur le PC
$NomPC = $env:COMPUTERNAME
$Windows = Get-CimInstance -ClassName Win32_OperatingSystem
$Processeur = Get-CimInstance -ClassName Win32_Processor

# 3. Memoire RAM en Go
$Ordinateur = Get-CimInstance -ClassName Win32_ComputerSystem
$RAM = [math]::Round($Ordinateur.TotalPhysicalMemory / 1GB, 2)

# 4. Informations reseau
$ConfigurationReseau = Get-NetIPConfiguration |
    Where-Object {
        $_.IPv4Address -and
        $_.IPv4Address.IPAddress -notlike "169.254.*"
    } |
    Select-Object -First 1
$AdresseIP = $ConfigurationReseau.IPv4Address.IPAddress
$Interface = Get-NetAdapter |
    Where-Object {
        $_.InterfaceIndex -eq $ConfigurationReseau.InterfaceIndex
    }
$AdresseMAC = $Interface.MacAddress

# 5. Informations sur les disques
$Disques = Get-CimInstance -ClassName Win32_LogicalDisk `
    -Filter "DriveType=3"
$InformationsDisques = ($Disques | ForEach-Object {
    $Taille = [math]::Round($_.Size / 1GB, 2)
    $Libre = [math]::Round($_.FreeSpace / 1GB, 2)
    "$($_.DeviceID) : $Libre Go libres sur $Taille Go"
}) -join " | "

# Calcul de l'utilisation du disque C:
$DisqueC = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"

if ($DisqueC -and $DisqueC.Size -gt 0) {
    $TailleDisque = [math]::Round($DisqueC.Size / 1GB, 2)
    $EspaceLibre = [math]::Round($DisqueC.FreeSpace / 1GB, 2)
    $EspaceUtilise = [math]::Round(($DisqueC.Size - $DisqueC.FreeSpace) / 1GB, 2)
    $PourcentageDisque = [math]::Round((($DisqueC.Size - $DisqueC.FreeSpace) / $DisqueC.Size) * 100, 1)
}
else {
    $TailleDisque = 0
    $EspaceLibre = 0
    $EspaceUtilise = 0
    $PourcentageDisque = 0
}

# Verification du pare-feu Windows

$PareFeu = Get-NetFirewallProfile

$PareFeuDomain = ($PareFeu |
    Where-Object { $_.Name -eq "Domain" }).Enabled

$PareFeuPrivate = ($PareFeu |
    Where-Object { $_.Name -eq "Private" }).Enabled

$PareFeuPublic = ($PareFeu |
    Where-Object { $_.Name -eq "Public" }).Enabled

# 6. Creation de l'inventaire
$Inventaire = [PSCustomObject]@{

    PareFeuDomain         = $PareFeuDomain
    PareFeuPrivate        = $PareFeuPrivate
    PareFeuPublic         = $PareFeuPublic
    NomPC                 = $NomPC
    Systeme               = $Windows.Caption
    VersionWindows        = $Windows.Version
    Processeur            = $Processeur.Name.Trim()
    RAM_Go                = $RAM
    AdresseIP             = $AdresseIP
    AdresseMAC            = $AdresseMAC
    Disques               = $InformationsDisques
    TailleDisque_Go       = $TailleDisque
    EspaceLibre_Go        = $EspaceLibre
    EspaceUtilise_Go      = $EspaceUtilise
    UtilisationDisque_Pct = $PourcentageDisque
    DateInventaire        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

# 7. Affichage dans PowerShell
$Inventaire | Format-List

# 8. Export Excel (CSV)
$Inventaire | Export-Csv `
    -Path "$Dossier\inventaire.csv" `
    -NoTypeInformation `
    -Encoding UTF8

# 9. Export Bloc-notes (TXT)
$Inventaire | Format-List |
    Out-String |
    Set-Content -Path "$Dossier\inventaire.txt" -Encoding UTF8

# 10. Generation du tableau de bord HTML

# Fonction pour securiser les valeurs HTML
function Convertir-HTML($Valeur) {
    return [System.Net.WebUtility]::HtmlEncode([string]$Valeur)
}

# Recuperation des informations
$PC_HTML = Convertir-HTML $Inventaire.NomPC
$Windows_HTML = Convertir-HTML $Inventaire.Systeme
$CPU_HTML = Convertir-HTML $Inventaire.Processeur
$RAM_HTML = Convertir-HTML $Inventaire.RAM_Go
$IP_HTML = Convertir-HTML $Inventaire.AdresseIP
$MAC_HTML = Convertir-HTML $Inventaire.AdresseMAC
$Disques_HTML = Convertir-HTML $Inventaire.Disques
$Date_HTML = Convertir-HTML $Inventaire.DateInventaire

# Informations de stockage
$Taille_HTML = $TailleDisque
$Libre_HTML = $EspaceLibre
$Utilise_HTML = $EspaceUtilise

# Format du pourcentage compatible avec le CSS
$Pourcentage_HTML = $PourcentageDisque.ToString(
    [System.Globalization.CultureInfo]::InvariantCulture
)

# Couleur selon le pourcentage d'utilisation
if ($PourcentageDisque -ge 90) {
    $CouleurDisque = "#ef4444"
    $EtatDisque = "Espace disque faible"
}
elseif ($PourcentageDisque -ge 75) {
    $CouleurDisque = "#f59e0b"
    $EtatDisque = "Stockage a surveiller"
}
else {
    $CouleurDisque = "#22c55e"
    $EtatDisque = "Stockage disponible"
}

# Preparation des indicateurs du pare-feu
function Get-EtatPareFeu($Actif) {
    if ($null -eq $Actif) {
        return "Inconnu"
    }
    if ($Actif) {
        return "Active"
    }
    return "Desactive"
}

$EtatDomain = Get-EtatPareFeu $PareFeuDomain
$EtatPrivate = Get-EtatPareFeu $PareFeuPrivate
$EtatPublic = Get-EtatPareFeu $PareFeuPublic

# Creation du contenu HTML
$HTML = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>IT Inventory Dashboard</title>

    <style>
        * {
            box-sizing: border-box;
        }

        body {
            background: #0f172a;
            color: #f8fafc;
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 40px;
        }

        .container {
            max-width: 1100px;
            margin: auto;
        }

        h1 {
            color: #38bdf8;
            margin-bottom: 8px;
        }

        .subtitle {
            color: #94a3b8;
            margin-bottom: 35px;
        }

        .dashboard {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        }

        .card {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 24px;
        }

        .card:hover {
            border-color: #38bdf8;
        }

        .card h2 {
            color: #94a3b8;
            font-size: 14px;
            font-weight: normal;
            margin-top: 0;
        }

        .card p {
            color: #ffffff;
            font-size: 20px;
            font-weight: bold;
            overflow-wrap: anywhere;
        }

        /* Jauge de stockage */

        .storage-card {
            margin-top: 30px;
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 24px;
        }

        .storage-card h2 {
            color: #38bdf8;
        }

        .storage-status {
            color: #cbd5e1;
            font-size: 14px;
        }

        .progress-track {
            width: 100%;
            height: 18px;
            background: #334155;
            border-radius: 20px;
            overflow: hidden;
            margin-top: 20px;
        }

        .progress-fill {
            height: 100%;
            border-radius: 20px;
        }

        .storage-values {
            display: flex;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 10px;
            color: #94a3b8;
            font-size: 14px;
        }

        .storage-total {
            color: #94a3b8;
            font-size: 13px;
            margin-top: 20px;
        }

        .details {
            margin-top: 30px;
            background: #1e293b;
            border-radius: 14px;
            padding: 24px;
        }

        .details p {
            overflow-wrap: anywhere;
        }

        .details strong {
            color: #38bdf8;
        }

        footer {
            text-align: center;
            margin-top: 40px;
            color: #94a3b8;
            font-size: 13px;
        }

        @media (max-width: 800px) {
            .dashboard {
                grid-template-columns: repeat(2, 1fr);
            }
        }

        @media (max-width: 500px) {
            body {
                padding: 20px;
            }

            .dashboard {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>

<body>
    <div class="container">

        <h1>IT Inventory Dashboard</h1>
        <p class="subtitle">
            Tableau de bord d'administration systeme et reseau
        </p>

        <div class="dashboard">

            <div class="card">
                <h2>Nom de l'ordinateur</h2>
                <p>$PC_HTML</p>
            </div>

            <div class="card">
                <h2>Systeme d'exploitation</h2>
                <p>$Windows_HTML</p>
            </div>

            <div class="card">
                <h2>Processeur</h2>
                <p>$CPU_HTML</p>
            </div>

            <div class="card">
                <h2>Memoire RAM</h2>
                <p>$RAM_HTML Go</p>
            </div>

            <div class="card">
                <h2>Adresse IPv4</h2>
                <p>$IP_HTML</p>
            </div>

            <div class="card">
                <h2>Adresse MAC</h2>
                <p>$MAC_HTML</p>
            </div>

        </div>

<!-- Audit du pare-feu Windows -->
<div class="storage-card">
    <h2>Audit du pare-feu Windows</h2>

    <p>Profil Domaine : $EtatDomain</p>
    <p>Profil Prive : $EtatPrivate</p>
    <p>Profil Public : $EtatPublic</p>
</div>

        <!-- Jauge de stockage -->
        <div class="storage-card">
            <h2>Stockage du disque C:</h2>

            <p class="storage-status">$EtatDisque</p>

            <div class="progress-track">
                <div class="progress-fill"
                     style="width: $Pourcentage_HTML%; background-color: $CouleurDisque;">
                </div>
            </div>

            <p>$Pourcentage_HTML% utilise</p>

            <div class="storage-values">
                <span>$Utilise_HTML Go utilises</span>
                <span>$Libre_HTML Go libres</span>
            </div>

            <p class="storage-total">
                Capacite totale : $Taille_HTML Go
            </p>
        </div>

        <div class="details">
            <h2>Informations complementaires</h2>
            <p><strong>Stockage :</strong> $Disques_HTML</p>
            <p><strong>Date de l'inventaire :</strong> $Date_HTML</p>
        </div>

        <footer>
            Projet d'inventaire informatique - PowerShell
        </footer>

    </div>
</body>
</html>
"@
# Enregistrement du tableau de bord
$HTML | Set-Content -Path "$Dossier\inventaire.html" -Encoding UTF8

Write-Host "Dashboard HTML genere avec succes !" -ForegroundColor Green

# 11. Confirmation
Write-Host "Les rapports ont ete generes dans C:\Projet" -ForegroundColor Green
