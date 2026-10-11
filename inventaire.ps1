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

