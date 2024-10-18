# Charger l'assemblée pour utiliser les formulaires Windows
Add-Type -AssemblyName System.Windows.Forms

# Attendre 5 secondes pour que l'environnement soit complètement initialisé
Start-Sleep -Seconds 1

# Suppression de toutes les imprimantes installées
Get-Printer | ForEach-Object {
    try {
        Remove-Printer -Name $_.Name
        Write-Host "L'imprimante $($_.Name) a été supprimée."
    } catch {
        Write-Host "Erreur lors de la suppression de l'imprimante $($_.Name)."
    }
}

# Si le script est exécuté sur un ordinateur en LOCAL alors nom du PC
$computerName = (Get-WmiObject -Class Win32_ComputerSystem).Name

# Si le script est exécuté sur un serveur RDS, récupère le nom du client à partir de la variable CLIENTNAME
if ($computerName -like "RDS*") {
    $sessionID = (Get-Process -PID $pid).SessionID
    $sessionCLIENTNAME = (Get-ItemProperty -path ("HKCU:\Volatile Environment\" + $sessionID) -name "CLIENTNAME").CLIENTNAME
    $computerName = $sessionCLIENTNAME

    # DEBOG : AFFICHAGE NOM DU PC
    #[System.Windows.Forms.MessageBox]::Show("Le PC s'appelle : $computerName")
}

# Fonction pour configurer l'imprimante
function Set-Printer {
    param (
        [string]$printerName
    )

    # Vérifie si l'imprimante existe avec la cmdlet Get-Printer
    $printer = Get-Printer -Name $printerName -ErrorAction SilentlyContinue

    # Si l'imprimante n'est pas trouvée, elle est ajoutée pour l'utilisateur
    if (-not $printer) {
        try {
            # Ajoute l'imprimante réseau à l'utilisateur
            Add-Printer -ConnectionName $printerName
            Write-Host "L'imprimante $printerName a été ajoutée."
        } catch {
            Write-Host "Impossible d'ajouter l'imprimante $printerName. Vérifiez que l'imprimante est accessible."
            exit
        }
    }

    # Définit l'imprimante par défaut avec rundll32
    $printerPath = $printerName.Replace('\', '\\')  # Escape des backslashes pour la commande
    $command = "rundll32 printui.dll,PrintUIEntry /y /n `"$printerName`""
    Invoke-Expression $command

    #[System.Windows.Forms.MessageBox]::Show("L'imprimante par défaut a été configurée pour $printerName.")
}
    $printerName = "Microsoft Print to PDF"
    $printer = Get-Printer | Where-Object { $_.Name -eq $printerName }
    if ($null -eq $printer) {
        Write-Host "Ajout de l'imprimante Microsoft Print to PDF..."
        Add-Printer -Name $printerName -DriverName "Microsoft Print To PDF" -PortName "PORTPROMPT:"
        Write-Host "L'imprimante Microsoft Print to PDF a été ajoutée avec succès."
    } else {
        Write-Host "L'imprimante Microsoft Print to PDF est déjà installée."
    }
# Vérification et configuration des imprimantes selon le préfixe du nom de l'ordinateur
if ($computerName -like "ADMDIR*") {
    # Vérifie si l'imprimante "Microsoft Print to PDF" est déjà installée
    $printerName = "Microsoft Print to PDF"
    $printer = Get-Printer | Where-Object { $_.Name -eq $printerName }
    if ($null -eq $printer) {
        Write-Host "Ajout de l'imprimante Microsoft Print to PDF..."
        Add-Printer -Name $printerName -DriverName "Microsoft Print To PDF" -PortName "PORTPROMPT:"
        Write-Host "L'imprimante Microsoft Print to PDF a été ajoutée avec succès."
    } else {
        Write-Host "L'imprimante Microsoft Print to PDF est déjà installée."
    }
    if ($computername -eq "ADMIDIRP009" -or $computerName -eq "ADMDIRP007") {
    Set-Printer "\\SRVchcaIMP\COPchcaRH"
    }
    else {
    Set-Printer "\\SRVchcaIMP\COPchcaADM"
    }
} elseif ($computerName -like "PORTDIRADJ1") {
    Set-Printer "\\SRVchcaIMP\COPchcaADM"
} elseif ($computerName -like "ADMACC*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaACC"
} elseif ($computerName -like "ADMANI*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaANI"
} elseif ($computerName -like "ADMCUI*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaCUI"
} elseif ($computerName -like "ADMBIO*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaMAG"
} elseif ($computerName -like "ADMSPEB*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaMJPM"
} elseif ($computerName -like "ADMTEC*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaATE"
} elseif ($computerName -like "ADMLOG*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaMAG"
} elseif ($computerName -like "ADMSYN*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaSYN"
} elseif ($computerName -like "QLF*") {
    # Vérification pour QLFCAD
    if ($computerName -like "QLFCAD*") {
        Set-Printer "\\SRVchcaIMP\IMPchcaIDCcad"
    } else {
        Set-Printer "\\SRVchcaIMP\IMPchcaQLFsds"  # Exemple par défaut pour QLF
    }
} elseif ($computerName -like "QGR*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaIDNsds"
} elseif ($computerName -like "PHA*") {
    Set-Printer "\\SRVchcaIMP\COPchcaPHA"
} elseif ($computerName -like "NOMAD*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaACC"
} elseif ($computerName -like "JDM*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaQLFsds"
} elseif ($computerName -like "PASA*") {
    Set-Printer "\\SRVchcaIMP\IMPchcaCS"
}
elseif ($computerName -like "IDN*") {
    # Vérification pour IDNCAD
    if ($computerName -like "IDNCAD*") {
        Set-Printer "\\SRVchcaIMP\IMPchcaIDNcad"
    } else {
        Set-Printer "\\SRVchcaIMP\IMPchcaIDNsds"  # Exemple par défaut pour IDN
    }
} elseif ($computerName -like "IDC*") {
    # Vérification des préfixes spécifiques pour IDC
    if ($computerName -like "IDCSOI*") {
        Set-Printer "\\SRVchcaIMP\IMPchcaIDCsds"
    } elseif ($computerName -like "IDCSPE*") {
        Set-Printer "\\SRVchcaIMP\IMPchcaANI"
    } else {
        Set-Printer "\\SRVchcaIMP\IMPchcaIDNsds"  # Exemple par défaut
    }
} elseif ($computerName -match "^MSSR.*") {
    $printerName = $null

    # Vérification pour les suffixes
    if ($computerName -match "(1\d{2})$") {
        $printerName = "\\SRVchcaIMP\IMPchcaMED"
    } elseif ($computerName -match "(2\d{2})$") {
        $printerName = "\\SRVchcaIMP\IMPchcaOCE"
    }

    # Vérification pour "CAD" dans le nom
    if ($computerName -match "CAD") {
        $printerName = "\\SRVchcaIMP\IMPchcaMEDcad"
    }

    if ($computerName -match "HDJ") {
        $printerName ="\\SRVchcaIMP\IMPchcaERGO"
    }

    # Vérification pour "SEC" dans le nom
    if ($computerName -match "SEC") {
        $printerName = "\\SRVchcaIMP\COPchcaMED"
    }
    
    # Vérification pour "SEC" dans le nom
    if ($computerName -match "SPEB2") {
        $printerName = "\\SRVchcaIMP\IMPchcaERGO"
    }

    # Assurez-vous qu'une imprimante a été définie avant d'appeler Set-Printer
    if ($printerName -ne $null) {
        Set-Printer $printerName
    } else {
        Write-Host "Aucune imprimante n'a été définie pour $computerName."
    }
}

# Charger les composants .NET pour interagir avec Active Directory
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Récupérer le nom d'utilisateur courant
$userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Se connecter au domaine Active Directory
$ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Domain)

# Récupérer l'utilisateur courant
$user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctx, $userName)

# Récupérer les groupes auxquels l'utilisateur appartient
$groups = $user.GetAuthorizationGroups()

# Créer une liste pour stocker les noms des groupes
$groupNames = @()

# Parcourir les groupes et ajouter uniquement leurs noms dans la liste
foreach ($group in $groups) {
    $groupNames += $group.Name
}

# Définir un tableau d'associations entre groupes et imprimantes
$groupPrinterMapping = @{
    "GRP_IMPchcaCS" = "\\SRVchcaIMP\IMPchcaCS"
    "GRP_COPchcaADM" = "\\SRVchcaIMP\COPchcaADM"
    "GRP_COPchcaMED" = "\\SRVchcaIMP\COPchcaMED"
    "GRP_COPchcaPHA" = "\\SRVchcaIMP\COPchcaPHA"
    "GRP_COPchcaRH" = "\\SRVchcaIMP\COPchcaRH"
    "GRP_IMPchcaACC" = "\\SRVchcaIMP\IMPchcaACC"
    "GRP_IMPchcaANI" = "\\SRVchcaIMP\IMPchcaANI"
    "GRP_IMPchcaATE" = "\\SRVchcaIMP\IMPchcaATE"
    "GRP_IMPchcaCUI" = "\\SRVchcaIMP\IMPchcaCUI"
    "GRP_IMPchcaERGO" = "\\SRVchcaIMP\IMPchcaERGO"
    "GRP_IMPchcaIDCcad" = "\\SRVchcaIMP\IMPchcaIDCcad"
    "GRP_IMPchcaIDCsds" = "\\SRVchcaIMP\IMPchcaIDCsds"
    "GRP_IMPchcaIDNcad" = "\\SRVchcaIMP\IMPchcaIDNcad"
    "GRP_IMPchcaIDNsds" = "\\SRVchcaIMP\IMPchcaIDNsds"
    "GRP_IMPchcaMAG" = "\\SRVchcaIMP\IMPchcaMAG"
    "GRP_IMPchcaMED" = "\\SRVchcaIMP\IMPchcaMED"
    "GRP_IMPchcaMEDcad" = "\\SRVchcaIMP\IMPchcaMEDcad"
    "GRP_IMPchcaMJPM" = "\\SRVchcaIMP\IMPchcaMJPM"
    "GRP_IMPchcaOCE" = "\\SRVchcaIMP\IMPchcaOCE"
    "GRP_IMPchcaQLFsds" = "\\SRVchcaIMP\IMPchcaQLFsds"
    "GRP_IMPchcaSYN" = "\\SRVchcaIMP\IMPchcaSYN"
    "G_IMP_IMPchcaPHAeti" = "\\SRVchcaIMP\IMPchcaPHAeti"
}

# Parcourir chaque groupe dans le tableau et vérifier l'appartenance
foreach ($group in $groupPrinterMapping.Keys) {
    if ($groupNames -contains $group) {
        #[System.Windows.Forms.MessageBox]::Show("Vous êtes dans le groupe $group.")

        # Obtenir le nom de l'imprimante correspondante
        $printerName = $groupPrinterMapping[$group]

        # Vérifier si l'imprimante est déjà ajoutée
        $printer = Get-Printer -Name $printerName -ErrorAction SilentlyContinue

        if (-not $printer) {
            try {
                # Ajoute l'imprimante réseau à l'utilisateur
                Add-Printer -ConnectionName $printerName
                #[System.Windows.Forms.MessageBox]::Show("L'imprimante $printerName a été ajoutée.")
            } catch {
                #[System.Windows.Forms.MessageBox]::Show("Impossible d'ajouter l'imprimante $printerName. Vérifiez que l'imprimante est accessible.")
            }
        } else {
            #[System.Windows.Forms.MessageBox]::Show("L'imprimante $printerName est déjà installée.")
        }
    }
}

# Si l'utilisateur n'appartient à aucun des groupes spécifiés
#[System.Windows.Forms.MessageBox]::Show("Toutes les imprimantes ont étés installées !")
