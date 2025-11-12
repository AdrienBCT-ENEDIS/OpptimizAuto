@echo off
REM Ouvre une boîte de dialogue pour selectionner un dossier, calcule la taille avant/apres et lance le script PowerShell

powershell -NoProfile -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$folder = (New-Object -ComObject Shell.Application).BrowseForFolder(0, 'Sélectionnez le dossier contenant les fichiers PPTX', 0);" ^
  "if ($folder) {" ^
  "  $folder = $folder.Self.Path;" ^
  "  $sizeBefore = (Get-ChildItem -LiteralPath $folder -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum;" ^
  "  $sizeBeforeMB = if ($sizeBefore) { [math]::Round(($sizeBefore/1MB),2) } else { 0 };" ^
  "  & '.\opptimizAuto.ps1' -FolderPath $folder;" ^
  "  $sizeAfter = (Get-ChildItem -LiteralPath $folder -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum;" ^
  "  $sizeAfterMB = if ($sizeAfter) { [math]::Round(($sizeAfter/1MB),2) } else { 0 };" ^
  "  $savedMB = [math]::Round(($sizeBeforeMB - $sizeAfterMB),2);" ^
  "  Add-Type -AssemblyName System.Windows.Forms;" ^
  "  $form = New-Object System.Windows.Forms.Form;" ^
  "  $form.Text = 'Resultats de l''optimisation';" ^
  "  $form.Width = 450; $form.Height = 250;" ^
  "  $form.StartPosition = 'CenterScreen';" ^
  "  $form.FormBorderStyle = 'FixedDialog';" ^
  "  $label = New-Object System.Windows.Forms.Label;" ^
  "  $label.Text = \"Taille avant optimisation : $sizeBeforeMB MB`n`nTaille apres optimisation : $sizeAfterMB MB`n`nEspace libere : $savedMB MB\";" ^
  "  $label.AutoSize = $true;" ^
  "  $label.Top = 50; $label.Left = 50;" ^
  "  $form.Controls.Add($label);" ^
  "  $button = New-Object System.Windows.Forms.Button;" ^
  "  $button.Text = 'OK'; $button.Top = 150; $button.Left = 170; $button.height = 35;" ^
  "  $button.Add_Click({ $form.Close() });" ^
  "  $form.Controls.Add($button);" ^
  "  [System.Windows.Forms.Application]::Run($form);" ^
  "} else {" ^
  "  Write-Host 'Aucun dossier sélectionné.';" ^
  "}" ^
  exit
