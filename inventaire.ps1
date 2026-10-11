# Projet : SCRIPT D'INVENTAIRE INFORMATIQUE
# Auteur : beyadem
# ==========================================

# 1. Informations sur le PC et Windows
$NomPC = $env:COMPUTERNAME
$Windows = Get-CimInstance -ClassName Win32_OperatingSystem
$Processeur = Get-CimInstance -ClassName Win32_Processor

# 2. Quantite de RAM en Go
$Ordinateur = Get-CimInstance -ClassName Win32_ComputerSystem
$RAM = [math]::Round($Ordinateur.TotalPhysicalMemory / 1GB, 2)

# 3. Informations reseau : adresse IP et adresse MAC
$ConfigurationReseau = Get-NetIPConfiguration |
    Where-Object {
        $_.IPv4Address -and
        $_.IPv4Address.IPAddress -notlike "169.254.*"
    } |
    Select-Object -First 1
$AdresseIP = $ConfigurationReseau.IPv4Address.IPAddress
$Interface = Get-NetAdapter |
    Where-Object { $_.InterfaceIndex -eq $ConfigurationReseau.InterfaceIndex }
$AdresseMAC = $Interface.MacAddress

# 4. Informations sur les disques
$Disques = Get-CimInstance -ClassName Win32_LogicalDisk `
    -Filter "DriveType=3"
$InformationsDisques = ($Disques | ForEach-Object {
    $Taille = [math]::Round($_.Size / 1GB, 2)
    $Libre = [math]::Round($_.FreeSpace / 1GB, 2)
    "$($_.DeviceID) : $Libre Go libres sur $Taille Go"
}) -join " | "

# 5. Creation de l'inventaire
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

# 6. Affichage des resultats
$Inventaire | Format-List

# 7. Export dans un fichier CSV
$Inventaire | Export-Csv `
    -Path "C:\Inventaire\inventaire.csv" `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host "Inventaire enregistre dans C:\Inventaire\inventaire.csv" -ForegroundColor Green

