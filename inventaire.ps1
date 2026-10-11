# ==========================================
# Projet : Inventaire informatique
# Auteur : beyadem
# Objectif : Recuperer les informations du PC
# ==========================================

# 1. Recuperer le nom de l'ordinateur
$NomPC = $env:COMPUTERNAME

# 2. Recuperer les informations de Windows
$Windows = Get-CimInstance -ClassName Win32_OperatingSystem

# 3. Recuperer les informations du processeur
$Processeur = Get-CimInstance -ClassName Win32_Processor

# 4. Afficher les resultats
Write-Host "===== INVENTAIRE INFORMATIQUE ====="
Write-Host "Nom du PC : $NomPC"
Write-Host "Systeme : $($Windows.Caption)"
Write-Host "Version Windows : $($Windows.Version)"
Write-Host "Processeur : $($Processeur.Name)"


# 5. Creer un objet contenant les informations
$Inventaire = [PSCustomObject]@{
    NomPC          = $NomPC
    Systeme        = $Windows.Caption
    VersionWindows = $Windows.Version
    Processeur     = $Processeur.Name.Trim()
    DateInventaire = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

# 6. Exporter les informations dans un fichier CSV
$Inventaire | Export-Csv -Path "C:\Inventaire\inventaire.csv" -NoTypeInformation -Encoding UTF8

# 7. Confirmer la creation du fichier
Write-Host "Inventaire enregistre dans C:\Inventaire\inventaire.csv"

