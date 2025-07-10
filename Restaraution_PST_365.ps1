<#
.DESCRIPTION
    - Ce script détache d'abord tous les fichiers PST montés dans Outlook,
    puis monte les fichiers PST présents dans un dossier spécifique,
    copie leur contenu dans un dossier "OldArchives" dans la boîte principale,
    et enfin les détache à nouveau.
    Fichier de log disponible dans D:\log.

.VERSION
    1

.AUTHOR
    COUTARD Théo

#>

# Ignore les messages d'erreur pour éviter d'interrompre le script
$ErrorActionPreference = "SilentlyContinue"

# Charge la bibliothèque Outlook et crée une instance de l'application Outlook
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# --- PARTIE 1 : Détacher les fichiers PST déjà montés ---
$Stores = $namespace.Stores
for ($i = $Stores.Count - 1; $i -ge 0; $i--) {
    # Vérifie si le magasin est un fichier PST (ExchangeStoreType = 3)
    if ($Stores[$i].ExchangeStoreType -eq 3) {
        $objFolder = $Stores[$i].GetRootFolder()
        $namespace.RemoveStore($objFolder)  # Détache le fichier PST
    }
}

# --- PARTIE 2 : Monter, copier et détacher les nouveaux fichiers PST ---

# Récupère le nom d'utilisateur courant
$user = $env:USERNAME
# Définit le chemin de base où se trouvent les fichiers PST
$pstBasePath = "D:\Utilisateurs\$user"
$pstFolderPath = ""

# Recherche un dossier nommé "pst" dans le répertoire utilisateur
$possibleFolders = Get-ChildItem -Path $pstBasePath -Directory | Where-Object { $_.Name.ToLower() -eq "pst" }

# Si un dossier "pst" est trouvé, on récupère son chemin
if ($possibleFolders.Count -gt 0) {
    $pstFolderPath = $possibleFolders[0].FullName
} else {
    Write-Output "Dossier PST introuvable pour l'utilisateur $user"
    exit  # Arrête le script si aucun dossier n'est trouvé
}

# Prépare le dossier et le fichier de log
$logDir = "D:\logs"
if (-not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null  # Crée le dossier s'il n'existe pas
}
$logDate = Get-Date -Format "yyyy-MM-dd"
$logFile = "$logDir\Log_migration_archives_$logDate.log"

# Fonction pour écrire dans le fichier de log
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$timestamp] $message"
}

# Log de démarrage
Write-Log "Script démarré pour l'utilisateur : $user"
Write-Log "Chemin PST détecté : $pstFolderPath"

# Récupère le dossier racine de la boîte principale
$defaultStore = $namespace.DefaultStore
$rootFolder = $defaultStore.GetRootFolder()

# Vérifie si le dossier "OldArchives" existe, sinon le crée
$oldArchivesFolder = $rootFolder.Folders | Where-Object { $_.Name -eq "OldArchives" }
if (-not $oldArchivesFolder) {
    $oldArchivesFolder = $rootFolder.Folders.Add("OldArchives")
    Write-Log "Dossier 'OldArchives' créé."
} else {
    Write-Log "Dossier 'OldArchives' déjà existant."
}

# Fonction récursive pour copier tous les dossiers et sous-dossiers
function Copy-AllFoldersRecursively {
    param ($sourceFolder, $destinationFolder)

    # Ignore les dossiers qui ne contiennent pas des mails
    if ($sourceFolder.DefaultItemType -ne 0) {
        Write-Log "Ignoré (non-mail) : $($sourceFolder.Name)"
        return
    }

    # Copie tous les éléments du dossier
    foreach ($item in @($sourceFolder.Items)) {
        try {
            $copiedItem = $item.Copy()
            $copiedItem.Move($destinationFolder) | Out-Null
        } catch {
            Write-Log "Erreur lors de la copie d'un élément dans $($sourceFolder.Name)"
        }
    }

    # Copie récursivement les sous-dossiers
    foreach ($subFolder in $sourceFolder.Folders) {
        try {
            if ($subFolder.DefaultItemType -eq 0) {
                $newSubFolder = $destinationFolder.Folders.Add($subFolder.Name)
                Copy-AllFoldersRecursively -sourceFolder $subFolder -destinationFolder $newSubFolder
            } else {
                Write-Log "Ignoré (non-mail) : $($subFolder.Name)"
            }
        } catch {
            Write-Log "Erreur lors de la copie du sous-dossier : $($subFolder.Name)"
        }
    }
}

# Seuil minimum de taille pour traiter un fichier PST (265 Ko)
$minSize = 265KB

# Récupère tous les fichiers PST valides dans le dossier
$pstFiles = Get-ChildItem -Path $pstFolderPath -Filter *.pst | Where-Object { $_.Length -gt $minSize }

# Pour chaque fichier PST trouvé
foreach ($pst in $pstFiles) {
    try {
        Write-Log "Montage de : $($pst.FullName)"
        $namespace.AddStore($pst.FullName)  # Monte le fichier PST

        # Récupère le magasin correspondant
        $store = $namespace.Stores | Where-Object { $_.FilePath -eq $pst.FullName }
        if ($store) {
            $pstRoot = $store.GetRootFolder()
            $targetFolderName = [System.IO.Path]::GetFileNameWithoutExtension($pst.Name)
            $targetFolder = $oldArchivesFolder.Folders.Add($targetFolderName)

            # Copie tous les dossiers de premier niveau
            foreach ($topFolder in $pstRoot.Folders) {
                if ($topFolder.DefaultItemType -eq 0) {
                    $newTopFolder = $targetFolder.Folders.Add($topFolder.Name)
                    Copy-AllFoldersRecursively -sourceFolder $topFolder -destinationFolder $newTopFolder
                } else {
                    Write-Log "Ignoré (non-mail) : $($topFolder.Name)"
                }
            }

            # Détache le fichier PST après traitement
            $namespace.RemoveStore($pstRoot)
            Write-Log "PST '$($pst.Name)' traité et détaché."
        }
    } catch {
        Write-Log "Erreur avec le fichier : $($pst.Name)"
    }
}

# Fin du script
Write-Log "Script terminé."