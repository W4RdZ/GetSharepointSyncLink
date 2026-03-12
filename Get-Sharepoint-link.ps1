#Requires -Version 5.0
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -NoExit -File `"$PSCommandPath`"" -Wait
    exit
}

# ===========================
# Script 1 : Generer liens SharePoint synchronises via OneDrive
# + Export automatique dans OneDrive\Documents\Backup Sharepoint
# ===========================

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }

try {

    #region ETAPE 0 : Copie des fichiers dans un dossier temporaire
    $settingsRoot = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings"
    $tempDir      = Join-Path $scriptDir ("IniTemp_" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    Write-Host "Recherche des dossiers Business*..." -ForegroundColor Gray
    $businessFolders = Get-ChildItem -Path $settingsRoot -Directory -Filter "Business*" -ErrorAction SilentlyContinue
    if (-not $businessFolders) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Aucun dossier Business* trouve dans $settingsRoot", "Erreur", "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }

    $copiedFiles = foreach ($folder in $businessFolders) {
        Get-ChildItem -Path $folder.FullName -Filter "*.ini" -File -ErrorAction SilentlyContinue |
            ForEach-Object {
                $dest = Join-Path $tempDir ($folder.Name + "_" + $_.Name)
                Copy-Item $_.FullName $dest -Force
                Get-Item $dest
            }
    }

    Write-Host "$($copiedFiles.Count) fichier(s) .ini copies" -ForegroundColor Gray

    if (-not $copiedFiles) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Aucun .ini trouve dans les dossiers Business*", "Erreur", "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }
    #endregion

    #region ETAPE 1 : Trouver le fichier GUID d'etat
    $guidPattern = '^Business\d+_[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\.ini$'

    $stateFile = $copiedFiles |
        Where-Object { $_.Name -match $guidPattern } |
        Where-Object { (Get-Content $_.FullName -Encoding Unicode) -match "libraryScope|libraryFolder" } |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $stateFile) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Fichier d'etat introuvable.", "Erreur", "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }

    Write-Host "Fichier d'etat trouve : $stateFile" -ForegroundColor Green
    #endregion

    #region ETAPE 2 : Pre-charger les policy files
    $templateBySiteTitle = @{}

    $copiedFiles | Where-Object { $_.Name -match '(?i)ClientPolicy' } | ForEach-Object {
        $content = Get-Content $_.FullName -Encoding Unicode -ErrorAction SilentlyContinue

        $titleLine    = $content | Where-Object { $_ -match '^\s*SiteTitle\s*=' }             | Select-Object -First 1
        $templateLine = $content | Where-Object { $_ -match '^\s*ViewOnlineUrlTemplate\s*=' } | Select-Object -First 1

        if ($titleLine -and $templateLine) {
            $siteTitle = ($titleLine    -replace '^\s*SiteTitle\s*=\s*"?', '').Trim().TrimEnd('"')
            $template  = ($templateLine -replace '^\s*ViewOnlineUrlTemplate\s*=\s*', '').Trim()
            $templateBySiteTitle[$siteTitle] = $template
        }
    }

    Write-Host "$($templateBySiteTitle.Count) policy file(s) charge(s)" -ForegroundColor Gray
    #endregion

    #region ETAPE 3 : Parser libraryScope et libraryFolder
    $allLines = Get-Content $stateFile -Encoding Unicode

    $scopeByIndex = @{}

    $allLines | Where-Object { $_ -match '^\s*libraryScope\s*=' } | ForEach-Object {
        $line       = $_
        $scopeIndex = $null
        if ($line -match '=\s*(\d+)\s+') { $scopeIndex = [int]$Matches[1] }

        $resourceId = $null
        if ($line -match '=\s*\d+\s+([0-9a-fA-F]+)') { $resourceId = $Matches[1] }

        $quoted      = [regex]::Matches($line, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }
        $siteTitle   = $quoted[0]
        $libraryName = $quoted[1]
        $siteUrl     = $quoted | Where-Object { $_ -match '^https://' } | Select-Object -First 1
        $localPath   = $quoted | Where-Object { $_ -match '^[A-Za-z]:\\' } | Select-Object -First 1

        if ($scopeIndex -ne $null) {
            $scopeByIndex[$scopeIndex] = [PSCustomObject]@{
                ScopeIndex  = $scopeIndex
                ResourceId  = $resourceId
                SiteTitle   = $siteTitle
                LibraryName = $libraryName
                SiteUrl     = $siteUrl
                LocalPath   = $localPath
            }
        }
    }

    $folders = [System.Collections.Generic.List[PSCustomObject]]::new()

    $allLines | Where-Object { $_ -match '^\s*libraryFolder\s*=' } | ForEach-Object {
        $line = $_

        if ($line -notmatch '=\s*\d+\s+(\d+)\s+([0-9a-fA-F]+)') { return }
        $scopeIndex = [int]$Matches[1]
        $resourceId = $Matches[2]

        $quoted     = [regex]::Matches($line, '"([^"]*)"') | ForEach-Object { $_.Groups[1].Value }
        $localPath  = $quoted | Where-Object { $_ -match '^[A-Za-z]:\\' } | Select-Object -First 1
        $folderName = $quoted | Where-Object { $_ -notmatch '^[A-Za-z]:\\' -and $_ -ne '' } | Select-Object -First 1

        $folders.Add([PSCustomObject]@{
            ScopeIndex  = $scopeIndex
            ResourceId  = $resourceId
            FolderName  = $folderName
            LocalPath   = $localPath
        })
    }

    Write-Host "$($scopeByIndex.Count) libraryScope(s) et $($folders.Count) libraryFolder(s) trouves" -ForegroundColor Gray
    #endregion

    #region ETAPE 4 : Construire les liens finaux
    $links = [System.Collections.Generic.List[PSCustomObject]]::new()

    $scopeIndexesWithFolder = $folders | Select-Object -ExpandProperty ScopeIndex -Unique

    function Resolve-SharePointUrl {
        param($SiteTitle, $ResourceId)
        $template = $templateBySiteTitle[$SiteTitle]
        if (-not $template) { return $null }
        if ($template -notmatch [Regex]::Escape('{ResourceID}')) { return $null }
        return $template -replace [Regex]::Escape('{ResourceID}'), $ResourceId
    }

    foreach ($scope in $scopeByIndex.Values | Sort-Object ScopeIndex) {
        if ($scope.SiteUrl -match 'sharepoint\.com/personal/') { continue }
        if ($scopeIndexesWithFolder -contains $scope.ScopeIndex) { continue }

        $finalUrl = Resolve-SharePointUrl -SiteTitle $scope.SiteTitle -ResourceId $scope.ResourceId
        if (-not $finalUrl) { continue }

        $links.Add([PSCustomObject]@{
            Type      = "Bibliotheque"
            SiteTitle = $scope.SiteTitle
            Name      = $scope.LibraryName
            LocalPath = $scope.LocalPath
            Url       = $finalUrl
        })
        Write-Host "  [LIB] $($scope.SiteTitle) - $($scope.LibraryName)" -ForegroundColor Cyan
        Write-Host "        $finalUrl" -ForegroundColor Gray
    }

    foreach ($folder in $folders) {
        $parentScope = $scopeByIndex[$folder.ScopeIndex]
        if (-not $parentScope) { continue }
        if ($parentScope.SiteUrl -match 'sharepoint\.com/personal/') { continue }

        $finalUrl = Resolve-SharePointUrl -SiteTitle $parentScope.SiteTitle -ResourceId $folder.ResourceId
        if (-not $finalUrl) { continue }

        $links.Add([PSCustomObject]@{
            Type      = "Dossier"
            SiteTitle = $parentScope.SiteTitle
            Name      = $folder.FolderName
            LocalPath = $folder.LocalPath
            Url       = $finalUrl
        })
        Write-Host "  [DIR] $($parentScope.SiteTitle) > $($folder.FolderName)" -ForegroundColor Yellow
        Write-Host "        $finalUrl" -ForegroundColor Gray
    }

    Write-Host "`n$($links.Count) lien(s) trouve(s)" -ForegroundColor Green
    #endregion

    #region ETAPE 5 : Definir le dossier d'export dans OneDrive personnel
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $oneDriveRoot = $env:OneDrive
    if (-not $oneDriveRoot -or -not (Test-Path $oneDriveRoot)) {
        [System.Windows.Forms.MessageBox]::Show("OneDrive personnel introuvable.", "Erreur", "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }

    $exportFolder = Join-Path $oneDriveRoot "Documents\Backup Sharepoint"
    $datasFolder  = Join-Path $exportFolder "datas"
    New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null
    New-Item -ItemType Directory -Path $datasFolder  -Force | Out-Null

    Write-Host "Dossier d'export : $exportFolder" -ForegroundColor Green
    #endregion

    #region ETAPE 6 : Generer le .csv dans datas\
    $csvPath = Join-Path $datasFolder "liens.csv"
    $links | Select-Object Type, SiteTitle, Name, Url | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV genere : $csvPath" -ForegroundColor Gray
    #endregion

    #region ETAPE 7 : Generer le Resynchroniser_SharePoint.ps1 dans datas\
    $ps1Path    = Join-Path $datasFolder "Resynchroniser_SharePoint.ps1"
    $ps1Content = @'
#Requires -Version 5.0
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$PSCommandPath`"" -Wait
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptDir2 = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$CsvPath    = Join-Path $scriptDir2 "liens.csv"

if (-not (Test-Path $CsvPath)) {
    [System.Windows.Forms.MessageBox]::Show("Fichier liens.csv introuvable : $CsvPath", "Erreur", "OK", "Error")
    exit
}

$liens = Import-Csv -Path $CsvPath -Encoding UTF8

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Resynchronisation SharePoint via OneDrive"
$form.Size            = New-Object System.Drawing.Size(620, 520)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::WhiteSmoke

$labelTitre           = New-Object System.Windows.Forms.Label
$labelTitre.Text      = "Selectionnez les bibliotheques a ouvrir :"
$labelTitre.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$labelTitre.Location  = New-Object System.Drawing.Point(20, 15)
$labelTitre.Size      = New-Object System.Drawing.Size(580, 30)
$form.Controls.Add($labelTitre)

$panel                = New-Object System.Windows.Forms.Panel
$panel.Location       = New-Object System.Drawing.Point(20, 55)
$panel.Size           = New-Object System.Drawing.Size(575, 360)
$panel.AutoScroll     = $true
$panel.BorderStyle    = "FixedSingle"
$panel.BackColor      = [System.Drawing.Color]::White
$form.Controls.Add($panel)

$checkboxes = @()
$yPos       = 10

foreach ($lien in $liens) {
    $cb          = New-Object System.Windows.Forms.CheckBox
    $cb.Text     = "[$($lien.Type)] $($lien.SiteTitle) - $($lien.Name)"
    $cb.Location = New-Object System.Drawing.Point(10, $yPos)
    $cb.Size     = New-Object System.Drawing.Size(545, 22)
    $cb.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
    $cb.Tag      = $lien.Url
    $cb.Checked  = $true
    $panel.Controls.Add($cb)
    $checkboxes += $cb
    $yPos       += 28
}

$btnAll          = New-Object System.Windows.Forms.Button
$btnAll.Text     = "Tout cocher"
$btnAll.Location = New-Object System.Drawing.Point(20, 430)
$btnAll.Size     = New-Object System.Drawing.Size(120, 32)
$btnAll.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$btnAll.Add_Click({ $checkboxes | ForEach-Object { $_.Checked = $true } })
$form.Controls.Add($btnAll)

$btnNone          = New-Object System.Windows.Forms.Button
$btnNone.Text     = "Tout decocher"
$btnNone.Location = New-Object System.Drawing.Point(150, 430)
$btnNone.Size     = New-Object System.Drawing.Size(120, 32)
$btnNone.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$btnNone.Add_Click({ $checkboxes | ForEach-Object { $_.Checked = $false } })
$form.Controls.Add($btnNone)

$btnSync           = New-Object System.Windows.Forms.Button
$btnSync.Text      = "Ouvrir les Sharepoint"
$btnSync.Location  = New-Object System.Drawing.Point(400, 428)
$btnSync.Size      = New-Object System.Drawing.Size(200, 34)
$btnSync.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnSync.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnSync.ForeColor = [System.Drawing.Color]::White
$btnSync.FlatStyle = "Flat"
$btnSync.Add_Click({
    $selected = $checkboxes | Where-Object { $_.Checked }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Veuillez selectionner au moins une bibliotheque.", "Attention", "OK", "Warning")
        return
    }
    foreach ($cb in $selected) {
        Start-Process $cb.Tag
        Start-Sleep -Milliseconds 500
    }
    [System.Windows.Forms.MessageBox]::Show("$($selected.Count) SharePoint(s) ouvert(s) !", "Succes", "OK", "Information")
    $form.Close()
})
$form.Controls.Add($btnSync)

$form.ShowDialog() | Out-Null
'@

    [System.IO.File]::WriteAllText($ps1Path, $ps1Content, [System.Text.Encoding]::UTF8)
    Write-Host "Script 2 genere : $ps1Path" -ForegroundColor Gray
    #endregion

    #region ETAPE 8 : Generer le raccourci .lnk avec icone shell32.dll,132
    $lnkPath  = Join-Path $exportFolder "restaurer les sharepoint.lnk"
    $shell    = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath       = "powershell.exe"
    $shortcut.Arguments        = "-ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$ps1Path`""
    $shortcut.WorkingDirectory = $datasFolder
    $shortcut.IconLocation     = "shell32.dll,132"
    $shortcut.Save()
    Write-Host "Raccourci .lnk genere : $lnkPath" -ForegroundColor Gray
    #endregion

    #region ETAPE 9 : Masquer le dossier datas et son contenu
    $datasInfo = Get-Item $datasFolder
    $datasInfo.Attributes = $datasInfo.Attributes -bor [System.IO.FileAttributes]::Hidden

    Get-ChildItem $datasFolder | ForEach-Object {
        $_.Attributes = $_.Attributes -bor [System.IO.FileAttributes]::Hidden
    }
    Write-Host "Dossier datas et contenu masques" -ForegroundColor Gray
    #endregion

    # Nettoyage
    Remove-Item $tempDir -Recurse -Force

    # Confirmation finale
    [System.Windows.Forms.MessageBox]::Show(
        "Backup termine !`n`n$($links.Count) lien(s) sauvegarde(s) dans :`n$exportFolder`n`nSur la nouvelle machine, ouvrez ce dossier depuis votre OneDrive et double-cliquez sur :`nrestaurer les sharepoint.lnk",
        "SharePoint Resync - Backup reussi",
        "OK",
        "Information"
    )

} catch {
    Write-Host "`nERREUR : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Ligne  : $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host "Detail : $($_.InvocationInfo.Line.Trim())" -ForegroundColor Yellow
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    Read-Host "`nAppuyez sur Entree pour fermer"
}
