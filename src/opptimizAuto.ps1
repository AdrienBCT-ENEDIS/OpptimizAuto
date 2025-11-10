param(
    [Parameter(Mandatory = $true)]
    [string]$FolderPath
)

# Ajouter le type User32 uniquement s'il n'existe pas déjà
# Nécessaire pour gérer la gestion des fenêtres 
if (-not ([System.Type]::GetType("User32"))) {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class User32 {
            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern IntPtr GetConsoleWindow();
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);
        }
"@
}

# Pour les SendKeys
Add-Type -AssemblyName System.Windows.Forms

# Minimiser le terminal PowerShell pour éviter les bugs lors des interactions simulées
$consoleHandle = [User32]::GetConsoleWindow()
if ($consoleHandle -ne [System.IntPtr]::Zero) {
    [User32]::ShowWindow($consoleHandle, 6) 
} else {
    Write-Host "Impossible de recuperer le handle de la console." -ForegroundColor Yellow
}

# Vérifier si le dossier existe
if (-not (Test-Path $FolderPath)) {
    Write-Host "Erreur : le dossier $FolderPath n'existe pas." -ForegroundColor Red
    exit 1
}

# Vérifier s'il y a des fichiers .pptx dans le dossier et ses sous-dossiers
$pptxFiles = Get-ChildItem -Path $FolderPath -Filter *.pptx -Recurse
if (-not $pptxFiles) {
    Write-Host "Aucun fichier .pptx trouve dans le dossier $FolderPath." -ForegroundColor Yellow
    exit 0
}

# Initialiser PowerPoint
$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.Visible = 1
$pptApp.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMaximized

# Forcer PowerPoint au premier plan
try {
    $hwnd = [System.IntPtr]::new($pptApp.HWND)
    if (-not [User32]::SetForegroundWindow($hwnd)) {
        Write-Host "Impossible de mettre PowerPoint au premier plan." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Erreur lors de la tentative de mise au premier plan de PowerPoint : $_" -ForegroundColor Red
}

# Récupérer la version complète depuis le registre
$officeRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$fullVersion = (Get-ItemProperty -Path $officeRegPath).VersionToReport

# Exemple : "16.0.17928.20708"
$buildNumber = [int]($fullVersion.Split(".")[2]) # Exemple : extrait "17928"
$thresholdBuild = 17928

# Déterminer si on est supérieur à 2408
$versionSup2408 = $buildNumber -gt $thresholdBuild
Write-Host "VERSION : $fullVersion\n $buildNumber"

Write-Host "Version supérieure à 2408 ? $versionSup2408"



try {
    $pptxFiles | ForEach-Object {
        $pptFile = $_.FullName
        Write-Host "Traitement du fichier : $pptFile" -ForegroundColor Cyan

        # Vérifier que le fichier existe
        if (-not (Test-Path $pptFile)) {
            Write-Host "Fichier introuvable : $pptFile" -ForegroundColor Red
            $hasErrors = $true
            return
        }

        try {
            Start-Sleep -Milliseconds 500 
            $presentation = $pptApp.Presentations.Open($pptFile)

            $presentation.Windows(1).Activate()
            Start-Sleep -Milliseconds 500 

            # Simuler le lancement de l'outil Optimizer avec les touches
            [System.Windows.Forms.SendKeys]::SendWait("%")
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
            Start-Sleep -Milliseconds 200

            # Ajuster le nombre de LEFT selon la version
            if ($versionSup2408) {
                [System.Windows.Forms.SendKeys]::SendWait("{LEFT}{LEFT}{LEFT}")  # 3 LEFT si > 2408
            } else {
                [System.Windows.Forms.SendKeys]::SendWait("{LEFT}{LEFT}")        # 2 LEFT sinon
            }

            [System.Windows.Forms.SendKeys]::SendWait(" ")

            Start-Sleep -Milliseconds 400

            [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{RIGHT}")
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB} ")
            [System.Windows.Forms.SendKeys]::SendWait(" ")
            Start-Sleep -Milliseconds 2250
            [System.Windows.Forms.SendKeys]::SendWait(" ")

            # Fermeture du power point actuel

            $presentation.Close()
                      

            Write-Host "Fichier optimise et ferme : $pptFile" -ForegroundColor Yellow
        } catch {
            Write-Host "Erreur lors du traitement du fichier : $pptFile" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            $hasErrors = $true
        }
    }
}
finally {
    Write-Host "Fermeture de PowerPoint..." -ForegroundColor Yellow

    if (-not $hasErrors) {
        Write-Host "Traitement termine avec succes !" -ForegroundColor Green
    } else {
        Write-Host "Traitement termine avec des erreurs. Verifiez les logs." -ForegroundColor Red
    }
    $pptApp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    # Terminer les processus PowerPoint restants
    Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}

