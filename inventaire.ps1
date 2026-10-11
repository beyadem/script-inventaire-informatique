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

# 6. Creation de l'inventaire
$Inventaire = [PSCustomObject]@{
    NomPC          = $NomPC
    Systeme        = $Windows.Caption
    VersionWindows = $Windows.Version
    Processeur     = $Processeur.Name.Trim()
    RAM_Go         = $RAM
    AdresseIP      = $AdresseIP
    AdresseMAC     = $AdresseMAC
    Disques        = $InformationsDisques
    DateInventaire = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
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

# 10. Generation du tableau de bord HTML

# Definition du style CSS
$CSS = @"
<style>
    body {
        background-color: #0f172a;
        color: #f1f5f9;
        font-family: Arial, sans-serif;
        margin: 40px;
    }

    h1 {
        color: #38bdf8;
        text-align: center;
        margin-bottom: 30px;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        background-color: #1e293b;
        border-radius: 10px;
    }

    th {
        background-color: #334155;
        color: #38bdf8;
        padding: 15px;
        text-align: left;
    }

    td {
        padding: 15px;
        border-bottom: 1px solid #475569;
    }

    tr:hover {
        background-color: #334155;
    }
</style>
"@
# Conversion de l'inventaire en HTML
$HTML = $Inventaire | ConvertTo-Html `
    -Title "Projet informatique" `
    -Head $CSS `
    -PreContent "<h1>Tableau de bord informatique</h1>"
# Enregistrement du fichier HTML
$HTML | Out-File -FilePath "$Dossier\inventaire.html" -Encoding UTF8

# 11. Confirmation
Write-Host "Les rapports ont ete generes dans C:\Projet" -ForegroundColor Green
