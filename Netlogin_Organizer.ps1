# script pour trier les résultats de la commande "show netlogin" extraits en CLI de XMC (XIQ site engine, extreme networks centralisation solution) 
# Le script génère un fichier excel et utilise des appels COM relatifs à EXCEL
# ---- /!\ Si vous n'avez pas EXCEL installé sur le système, les commandes ne fonctionneront pas, seul les CSV seront générés /!\
# ---- /!\ Powershell version 7.4 /!\ ---- 
# Vive MCS
# Tom Médine

# Script to sort the results of the "show netlogin" command extracted from a quantity of XMC (XIQ site engine, extreme networks centralisation solution) 
# The script generates an Excel file and uses COM calls related to EXCEL
# ---- /!\ If you do not have EXCEL installed on the system, the commands will not work; only CSV files will be generated /!\
# ---- /!\ PowerShell version 7.4 /!\ ---- 
# Long live MCS
# Tom Médine

# Check PowerShell version
$requiredVersion = [version]"7.4"
$currentVersion = $PSVersionTable.PSVersion

if ($currentVersion -lt $requiredVersion) {
    Write-Host "Erreur : Ce script nécessite PowerShell version 7.4 ou supérieure. Vous utilisez la version $currentVersion." -ForegroundColor Red
    exit
}

Write-Host "Votre version de PowerShell $currentVersion"

# Demander à l'utilisateur de saisir le chemin du dossier parent
$rootFolder = Read-Host "Veuillez entrer le chemin complet du dossier parent où se trouvent les sous-dossiers ex: C:\Users\tom\Export NAC\"

# Confirmation du chemin
$confirmation = Read-Host "Confirmer le chemin : $rootFolder (oui/non)"

# Tant que la confirmation est "non", redemander le chemin
while ($confirmation -ne "oui") {
    Write-Host "Veuillez ressaisir le chemin."
    $rootFolder = Read-Host "Veuillez entrer le chemin complet du dossier parent où se trouvent les sous-dossiers ex: C:\Users\tom\Export NAC\"
    $confirmation = Read-Host "Confirmer le chemin : $rootFolder (oui/non)"
}

# Vérifier si le dossier existe
if (-not (Test-Path -Path $rootFolder)) {
    Write-Host "Le dossier spécifié n'existe pas. Veuillez réessayer." -ForegroundColor Red
    exit
}

# Créer un nouveau dossier "Exports CSV" dans le dossier racine pour les fichiers CSV
$csvExportFolder = Join-Path -Path $rootFolder -ChildPath "Exports CSV"
if (-not (Test-Path -Path $csvExportFolder)) {
    New-Item -Path $csvExportFolder -ItemType Directory
}

# Fonction pour analyser un fichier texte et extraire les informations de Netlogin
function Analyze-File {
    param (
        [string]$filePath
    )

    # Lire le contenu du fichier texte
    $fileContent = Get-Content -Path $filePath

    # Extraire le nom du switch depuis la première ligne (avant le premier "." ou le caractère "#")
    $switchName = ($fileContent[0] -split '[\.\s#]')[0]

    # Créer un tableau pour stocker les statuts des ports (de 1 à 54)
    $portsStatus = @{ }
    for ($i = 1; $i -le 54; $i++) {
        $portsStatus[$i] = "NAC pas ok"  # Par défaut, tous les ports sont considérés comme "NAC pas ok"
    }

    # Parcourir les lignes du fichier et vérifier si Netlogin est activé sur un port
    foreach ($line in $fileContent) {
        if ($line -match "Port:\s*(\d+),\s*State:\s*Enabled") {
            $portNumber = [int]$matches[1]
            $portsStatus[$portNumber] = "NAC OK"  # Si activé, marquer le port comme "NAC OK"
        }
    }

    return [PSCustomObject]@{
        SwitchName  = $switchName
        PortsStatus = $portsStatus
    }
}

# Tableau pour stocker le chemin des CSVs pour une fusion ultérieure
$csvFiles = @()

# Parcourir chaque sous-dossier du dossier racine
Get-ChildItem -Path $rootFolder -Directory | Where-Object { $_.Name -ne "Exports CSV" } | ForEach-Object {
    $subFolder = $_.FullName
    $folderName = $_.Name  # Nom du dossier utilisé pour nommer le fichier CSV

    # Tableau pour stocker toutes les informations pour le CSV final
    $csvData = @()

    # Parcourir tous les fichiers texte du sous-dossier
    Get-ChildItem -Path $subFolder -Filter "*.txt" | ForEach-Object {
        $filePath = $_.FullName

        # Analyser le fichier texte pour obtenir le nom du switch et le statut des ports
        $switchData = Analyze-File -filePath $filePath

        # Préparer les données pour le CSV
        foreach ($port in $switchData.PortsStatus.Keys) {
            $csvData += [PSCustomObject]@{
                Switch     = $switchData.SwitchName
                Port       = $port
                Netlogin   = $switchData.PortsStatus[$port]
            }
        }
    }

    # Trier les données par nom du switch et par numéro de port
    $sortedData = $csvData | Sort-Object Switch, Port

    # Utiliser le nom du dossier pour nommer le fichier CSV
    $csvFileName = "Netlogin_Status_$folderName.csv"
    $csvPath = Join-Path -Path $csvExportFolder -ChildPath $csvFileName
    $sortedData | Export-Csv -Path $csvPath -NoTypeInformation

    # Ajouter le chemin du fichier CSV à la liste pour la fusion dans Excel
    $csvFiles += $csvPath
}

Write-Host "Tous les fichiers CSV ont été exportés dans le dossier 'Exports CSV'." -ForegroundColor Green
Write-Host "Création du fichier Excel... " -ForegroundColor Green

# Generate a timestamp for the Excel file name
$timestamp = Get-Date -Format "dd-MM-yyyy HH-mm"
$excelFileName = "NAC_Results_Compiled ($timestamp).xlsx"
$excelPath = Join-Path -Path $csvExportFolder -ChildPath $excelFileName

# Check if the Excel file already exists
if (Test-Path $excelPath) {
    $response = Read-Host "Le fichier existe déjà. Voulez-vous le remplacer ? (oui/non)"
    
    if ($response -eq "oui") {
        Remove-Item $excelPath -Force  # Delete the existing file
        Write-Host "Le fichier existant a été supprimé."
    } else {
        Write-Host "Le fichier existant ne sera pas remplacé. Arrêt du script."
        exit
    }
}

# Create an Excel application COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false  # Set to $true if you want to see the Excel window
$workbook = $excel.Workbooks.Add()

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $csvExportFolder -Filter "*.csv"

# Check if any CSV files were found
if ($csvFiles.Count -eq 0) {
    Write-Host "Erreur : Aucun fichier CSV trouvé dans le dossier spécifié."
    $excel.Quit()
    exit
}

# Variable to track whether at least one sheet has been added
$sheetAdded = $false

# Iterate over each CSV file
foreach ($csvFile in $csvFiles) {
    # Create a new worksheet for each CSV
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($csvFile.FullName)
    $worksheet = $workbook.Sheets.Add()
    $worksheet.Name = $sheetName

    # Mark that a sheet has been added
    $sheetAdded = $true

    # Read CSV file and properly handle quoted fields
    $csvData = Import-Csv -Path $csvFile.FullName

    # Write the header row (CSV column names)
    $header = $csvData[0].PSObject.Properties.Name
    $row = 1
    $column = 1
    foreach ($columnHeader in $header) {
        $worksheet.Cells.Item($row, $column) = $columnHeader
        $column++
    }

    # Write the data rows
    $row++
    foreach ($record in $csvData) {
        $column = 1
        foreach ($value in $record.PSObject.Properties.Value) {
            $worksheet.Cells.Item($row, $column) = $value
            $column++
        }
        $row++
    }
}

# Save the Excel workbook
try {
    $workbook.SaveAs($excelPath)
    Write-Host "Le fichier Excel a été créé avec succès à $excelPath" -ForegroundColor Yellow
} catch {
    Write-Host "Erreur : Impossible d'enregistrer le fichier Excel. Veuillez vérifier le chemin et les permissions du fichier."
} finally {
    # Clean up
    $workbook.Close()
    $excel.Quit()

    # Release COM objects to free up resources
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Force garbage collection
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Profitez bien de ces informations inestimables sur le NAC, bonne journée merci au revoir"
