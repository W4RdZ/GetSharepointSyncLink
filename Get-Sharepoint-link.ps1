#Requires -Version 5.0
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -NoExit -File `"$PSCommandPath`"" -Wait
    exit
}

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }

#region I18N : Detection de la langue et chargement des traductions
$uiLang = [System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName.ToLower()

$i18n = @{}

if ($uiLang -eq 'fr') {
    $i18n = @{
        # Console - script 1
        SearchingBusiness       = "Recherche des dossiers Business*..."
        IniCopied               = "{0} fichier(s) .ini copies"
        StateFileFound          = "Fichier d'etat trouve : {0}"
        PolicyLoaded            = "{0} policy file(s) charge(s)"
        ScopeFolderFound        = "{0} libraryScope(s) et {1} libraryFolder(s) trouves"
        EntriesGenerated        = "{0} entree(s) generee(s)"
        ExportFolder            = "Dossier d'export : {0}"
        CsvGenerated            = "CSV genere : {0}"
        Script2Generated        = "Script 2 genere : {0}"
        LnkGenerated            = "Raccourci .lnk genere : {0}"
        DatasHidden             = "Dossier datas et contenu masques"
        ErrorLine               = "Ligne  : {0}"
        ErrorDetail             = "Detail : {0}"
        PressEnter              = "Appuyez sur Entree pour fermer"
        # MessageBox - script 1
        ErrNoBusinessFolder     = "Aucun dossier Business* trouve dans {0}"
        ErrNoIni                = "Aucun .ini trouve dans les dossiers Business*"
        ErrNoStateFile          = "Fichier d'etat introuvable."
        ErrNoOneDrive           = "OneDrive personnel introuvable."
        ErrTitle                = "Erreur"
        SuccessMsg              = "Backup termine !`n`n{0} lien(s) sauvegarde(s) dans :`n{1}`n`nSur la nouvelle machine, ouvrez ce dossier depuis votre OneDrive et double-cliquez sur :`nrestaurer les sharepoint.lnk"
        SuccessTitle            = "SharePoint Resync - Backup reussi"
        # GUI - script 2 genere
        FormTitle               = "Resynchronisation SharePoint via OneDrive"
        LabelTitre              = "Selectionnez uniquement les SharePoint que vous souhaitez resynchroniser."
        BtnAll                  = "Tout cocher"
        BtnNone                 = "Tout decocher"
        BtnCancel               = "Annuler"
        BtnSync                 = "Ouvrir les SharePoint"
        ErrCsvNotFound          = "Fichier liens.csv introuvable : {0}"
        WarnSelectOne           = "Veuillez selectionner au moins une bibliotheque."
        WarnTitle               = "Attention"
    }
} else {
    # Fallback : anglais
    $i18n = @{
        # Console - script 1
        SearchingBusiness       = "Searching for Business* folders..."
        IniCopied               = "{0} .ini file(s) copied"
        StateFileFound          = "State file found: {0}"
        PolicyLoaded            = "{0} policy file(s) loaded"
        ScopeFolderFound        = "{0} libraryScope(s) and {1} libraryFolder(s) found"
        EntriesGenerated        = "{0} entry/entries generated"
        ExportFolder            = "Export folder: {0}"
        CsvGenerated            = "CSV generated: {0}"
        Script2Generated        = "Script 2 generated: {0}"
        LnkGenerated            = "Shortcut .lnk generated: {0}"
        DatasHidden             = "Datas folder and contents hidden"
        ErrorLine               = "Line   : {0}"
        ErrorDetail             = "Detail : {0}"
        PressEnter              = "Press Enter to close"
        # MessageBox - script 1
        ErrNoBusinessFolder     = "No Business* folder found in {0}"
        ErrNoIni                = "No .ini file found in Business* folders"
        ErrNoStateFile          = "State file not found."
        ErrNoOneDrive           = "Personal OneDrive not found."
        ErrTitle                = "Error"
        SuccessMsg              = "Backup complete!`n`n{0} link(s) saved to:`n{1}`n`nOn the new machine, open this folder from your OneDrive and double-click:`nrestaurer les sharepoint.lnk"
        SuccessTitle            = "SharePoint Resync - Backup successful"
        # GUI - script 2 genere
        FormTitle               = "SharePoint Resynchronization via OneDrive"
        LabelTitre              = "Select only the SharePoint libraries you want to resynchronize."
        BtnAll                  = "Select all"
        BtnNone                 = "Deselect all"
        BtnCancel               = "Cancel"
        BtnSync                 = "Open SharePoint"
        ErrCsvNotFound          = "File liens.csv not found: {0}"
        WarnSelectOne           = "Please select at least one library."
        WarnTitle               = "Warning"
    }
}
#endregion

try {

    #region ETAPE 0 : Copie des fichiers dans un dossier temporaire
    $settingsRoot = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings"
    $tempDir      = Join-Path $scriptDir ("IniTemp_" + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    Write-Host $i18n.SearchingBusiness -ForegroundColor Gray
    $businessFolders = Get-ChildItem -Path $settingsRoot -Directory -Filter "Business*" -ErrorAction SilentlyContinue
    if (-not $businessFolders) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(($i18n.ErrNoBusinessFolder -f $settingsRoot), $i18n.ErrTitle, "OK", "Error")
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

    Write-Host ($i18n.IniCopied -f $copiedFiles.Count) -ForegroundColor Gray

    if (-not $copiedFiles) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($i18n.ErrNoIni, $i18n.ErrTitle, "OK", "Error")
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
        [System.Windows.Forms.MessageBox]::Show($i18n.ErrNoStateFile, $i18n.ErrTitle, "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }

    Write-Host ($i18n.StateFileFound -f $stateFile) -ForegroundColor Green
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

    Write-Host ($i18n.PolicyLoaded -f $templateBySiteTitle.Count) -ForegroundColor Gray
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
                IsSynced    = ($null -ne $localPath -and $localPath -ne '')
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
            DisplayName = $folderName
            LocalPath   = $localPath
        })
    }

    Write-Host ($i18n.ScopeFolderFound -f $scopeByIndex.Count, $folders.Count) -ForegroundColor Gray
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

        if ($scopeIndexesWithFolder -contains $scope.ScopeIndex) {
            $links.Add([PSCustomObject]@{
                Type              = "Header"
                SiteTitle         = $scope.SiteTitle
                Name              = $scope.LibraryName
                ParentLibraryName = ''
                DisplayName       = $scope.SiteTitle
                LocalPath         = $null
                Url               = $null
                IsSynced          = $false
                IsLast            = $false
            })
            Write-Host "  [HDR] $($scope.SiteTitle)" -ForegroundColor Magenta

            $children = @($folders | Where-Object { $_.ScopeIndex -eq $scope.ScopeIndex })
            for ($i = 0; $i -lt $children.Count; $i++) {
                $folder   = $children[$i]
                $isLast   = ($i -eq $children.Count - 1)
                $finalUrl = Resolve-SharePointUrl -SiteTitle $scope.SiteTitle -ResourceId $folder.ResourceId
                if (-not $finalUrl) { continue }

                $links.Add([PSCustomObject]@{
                    Type              = "Dossier"
                    SiteTitle         = $scope.SiteTitle
                    Name              = $folder.FolderName
                    ParentLibraryName = $scope.LibraryName
                    DisplayName       = $folder.FolderName
                    LocalPath         = $folder.LocalPath
                    Url               = $finalUrl
                    IsSynced          = $true
                    IsLast            = $isLast
                })
                Write-Host "    $($folder.FolderName)" -ForegroundColor Yellow
            }

        } elseif ($scope.IsSynced) {
            $finalUrl    = Resolve-SharePointUrl -SiteTitle $scope.SiteTitle -ResourceId $scope.ResourceId
            if (-not $finalUrl) { continue }
            $displayName = Split-Path $scope.LocalPath -Leaf

            $links.Add([PSCustomObject]@{
                Type              = "Bibliotheque"
                SiteTitle         = $scope.SiteTitle
                Name              = $scope.LibraryName
                ParentLibraryName = ''
                DisplayName       = $displayName
                LocalPath         = $scope.LocalPath
                Url               = $finalUrl
                IsSynced          = $true
                IsLast            = $false
            })
            Write-Host "  [LIB] $displayName" -ForegroundColor Cyan
            Write-Host "        $finalUrl" -ForegroundColor Gray
        }
    }

    Write-Host ("`n" + ($i18n.EntriesGenerated -f $links.Count)) -ForegroundColor Green
    #endregion

    #region ETAPE 5 : Definir le dossier d'export dans OneDrive personnel
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $oneDriveRoot = $env:OneDrive
    if (-not $oneDriveRoot -or -not (Test-Path $oneDriveRoot)) {
        [System.Windows.Forms.MessageBox]::Show($i18n.ErrNoOneDrive, $i18n.ErrTitle, "OK", "Error")
        Remove-Item $tempDir -Recurse -Force
        exit
    }

    $exportFolder = Join-Path $oneDriveRoot "Documents\Backup Sharepoint"
    $datasFolder  = Join-Path $exportFolder "datas"
    New-Item -ItemType Directory -Path $exportFolder -Force | Out-Null

    if (Test-Path $datasFolder) {
        Get-ChildItem $datasFolder -Force | ForEach-Object { $_.Attributes = [System.IO.FileAttributes]::Normal }
        $datasInfo = Get-Item $datasFolder -Force
        $datasInfo.Attributes = [System.IO.FileAttributes]::Normal
        Remove-Item $datasFolder -Recurse -Force
    }

    New-Item -ItemType Directory -Path $datasFolder -Force | Out-Null
    Write-Host ($i18n.ExportFolder -f $exportFolder) -ForegroundColor Green
    #endregion

    #region ETAPE 6 : Generer le .csv dans datas\
    $csvPath = Join-Path $datasFolder "liens.csv"
    $links | Where-Object { $_.IsSynced } |
        Select-Object Type, SiteTitle, Name, ParentLibraryName, DisplayName, Url, IsLast |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host ($i18n.CsvGenerated -f $csvPath) -ForegroundColor Gray
    #endregion

    #region ETAPE 7 : Generer le Resynchroniser_SharePoint.ps1 dans datas\
    # Les chaines traduites a embarquer dans le script genere sont interpolees ici,
    # au moment de la generation, depuis la hashtable $i18n deja chargee.
    $ps1Path = Join-Path $datasFolder "Resynchroniser_SharePoint.ps1"

    $g_FormTitle     = $i18n.FormTitle
    $g_LabelTitre    = $i18n.LabelTitre
    $g_BtnAll        = $i18n.BtnAll
    $g_BtnNone       = $i18n.BtnNone
    $g_BtnCancel     = $i18n.BtnCancel
    $g_BtnSync       = $i18n.BtnSync
    $g_ErrCsv        = $i18n.ErrCsvNotFound -replace '"', '`"'
    $g_ErrTitle      = $i18n.ErrTitle
    $g_WarnSelectOne = $i18n.WarnSelectOne
    $g_WarnTitle     = $i18n.WarnTitle

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('#Requires -Version 5.0')
    $lines.Add('if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne ''STA'') {')
    $lines.Add('    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -WindowStyle Hidden -File `"$PSCommandPath`"" -Wait')
    $lines.Add('    exit')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('Add-Type -AssemblyName System.Windows.Forms')
    $lines.Add('Add-Type -AssemblyName System.Drawing')
    $lines.Add('')
    $lines.Add('$scriptDir2 = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }')
    $lines.Add('$CsvPath    = Join-Path $scriptDir2 "liens.csv"')
    $lines.Add('')
    $lines.Add('if (-not (Test-Path $CsvPath)) {')
    $lines.Add('    [System.Windows.Forms.MessageBox]::Show(("' + $g_ErrCsv + '" -f $CsvPath), "' + $g_ErrTitle + '", "OK", "Error")')
    $lines.Add('    exit')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('$liens = Import-Csv -Path $CsvPath -Encoding UTF8')
    $lines.Add('')
    $lines.Add('$form                 = New-Object System.Windows.Forms.Form')
    $lines.Add('$form.Text            = "' + $g_FormTitle + '"')
    $lines.Add('$form.Size            = New-Object System.Drawing.Size(620, 560)')
    $lines.Add('$form.MinimumSize     = New-Object System.Drawing.Size(620, 440)')
    $lines.Add('$form.StartPosition   = "CenterScreen"')
    $lines.Add('$form.FormBorderStyle = "Sizable"')
    $lines.Add('$form.MaximizeBox     = $true')
    $lines.Add('$form.BackColor       = [System.Drawing.Color]::WhiteSmoke')
    $lines.Add('')
    $lines.Add('$labelTitre           = New-Object System.Windows.Forms.Label')
    $lines.Add('$labelTitre.Text      = "' + $g_LabelTitre + '"')
    $lines.Add('$labelTitre.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)')
    $lines.Add('$labelTitre.Location  = New-Object System.Drawing.Point(20, 15)')
    $lines.Add('$labelTitre.Size      = New-Object System.Drawing.Size(580, 50)')
    $lines.Add('$form.Controls.Add($labelTitre)')
    $lines.Add('')
    $lines.Add('$panel             = New-Object System.Windows.Forms.Panel')
    $lines.Add('$panel.Location    = New-Object System.Drawing.Point(20, 75)')
    $lines.Add('$panel.Size        = New-Object System.Drawing.Size(575, 390)')
    $lines.Add('$panel.AutoScroll  = $true')
    $lines.Add('$panel.BorderStyle = "FixedSingle"')
    $lines.Add('$panel.BackColor   = [System.Drawing.Color]::White')
    $lines.Add('$form.Controls.Add($panel)')
    $lines.Add('')
    $lines.Add('$checkboxes = @()')
    $lines.Add('$yPos       = 10')
    $lines.Add('')
    $lines.Add('$bibliotheques = $liens | Where-Object { $_.Type -eq "Bibliotheque" }')
    $lines.Add('$grouped = [ordered]@{}')
    $lines.Add('foreach ($lien in ($liens | Where-Object { $_.Type -eq "Dossier" })) {')
    $lines.Add('    $key = "$($lien.SiteTitle)|$($lien.ParentLibraryName)"')
    $lines.Add('    if (-not $grouped.Contains($key)) { $grouped[$key] = [System.Collections.Generic.List[object]]::new() }')
    $lines.Add('    $grouped[$key].Add($lien)')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('foreach ($lien in $bibliotheques) {')
    $lines.Add('    $cb          = New-Object System.Windows.Forms.CheckBox')
    $lines.Add('    $cb.Text     = $lien.DisplayName')
    $lines.Add('    $cb.Location = New-Object System.Drawing.Point(10, $yPos)')
    $lines.Add('    $cb.Size     = New-Object System.Drawing.Size(545, 22)')
    $lines.Add('    $cb.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)')
    $lines.Add('    $cb.Tag      = $lien.Url')
    $lines.Add('    $cb.Checked  = $false')
    $lines.Add('    $panel.Controls.Add($cb)')
    $lines.Add('    $checkboxes += $cb')
    $lines.Add('    $yPos += 30')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('foreach ($key in $grouped.Keys) {')
    $lines.Add('    $items = $grouped[$key]')
    $lines.Add('    $parts = $key -split ''\|''')
    $lines.Add('    $siteTitle = $parts[0]')
    $lines.Add('')
    $lines.Add('    $lbl           = New-Object System.Windows.Forms.Label')
    $lines.Add('    $lbl.Text      = $siteTitle')
    $lines.Add('    $lbl.Location  = New-Object System.Drawing.Point(10, $yPos)')
    $lines.Add('    $lbl.Size      = New-Object System.Drawing.Size(545, 22)')
    $lines.Add('    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)')
    $lines.Add('    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(0, 90, 160)')
    $lines.Add('    $panel.Controls.Add($lbl)')
    $lines.Add('    $yPos += 26')
    $lines.Add('')
    $lines.Add('    for ($i = 0; $i -lt $items.Count; $i++) {')
    $lines.Add('        $lien        = $items[$i]')
    $lines.Add('        $cb          = New-Object System.Windows.Forms.CheckBox')
    $lines.Add('        $cb.Text     = $lien.DisplayName')
    $lines.Add('        $cb.Location = New-Object System.Drawing.Point(20, $yPos)')
    $lines.Add('        $cb.Size     = New-Object System.Drawing.Size(535, 22)')
    $lines.Add('        $cb.Font     = New-Object System.Drawing.Font("Segoe UI", 9)')
    $lines.Add('        $cb.Tag      = $lien.Url')
    $lines.Add('        $cb.Checked  = $false')
    $lines.Add('        $panel.Controls.Add($cb)')
    $lines.Add('        $checkboxes += $cb')
    $lines.Add('        $yPos += 26')
    $lines.Add('    }')
    $lines.Add('    $yPos += 8')
    $lines.Add('}')
    $lines.Add('')
    $lines.Add('$btnAll          = New-Object System.Windows.Forms.Button')
    $lines.Add('$btnAll.Text     = "' + $g_BtnAll + '"')
    $lines.Add('$btnAll.Location = New-Object System.Drawing.Point(20, 480)')
    $lines.Add('$btnAll.Size     = New-Object System.Drawing.Size(120, 32)')
    $lines.Add('$btnAll.Font     = New-Object System.Drawing.Font("Segoe UI", 9)')
    $lines.Add('$btnAll.Add_Click({ $checkboxes | ForEach-Object { $_.Checked = $true } })')
    $lines.Add('$form.Controls.Add($btnAll)')
    $lines.Add('')
    $lines.Add('$btnNone          = New-Object System.Windows.Forms.Button')
    $lines.Add('$btnNone.Text     = "' + $g_BtnNone + '"')
    $lines.Add('$btnNone.Location = New-Object System.Drawing.Point(150, 480)')
    $lines.Add('$btnNone.Size     = New-Object System.Drawing.Size(120, 32)')
    $lines.Add('$btnNone.Font     = New-Object System.Drawing.Font("Segoe UI", 9)')
    $lines.Add('$btnNone.Add_Click({ $checkboxes | ForEach-Object { $_.Checked = $false } })')
    $lines.Add('$form.Controls.Add($btnNone)')
    $lines.Add('')
    $lines.Add('$btnCancel           = New-Object System.Windows.Forms.Button')
    $lines.Add('$btnCancel.Text      = "' + $g_BtnCancel + '"')
    $lines.Add('$btnCancel.Location  = New-Object System.Drawing.Point(290, 478)')
    $lines.Add('$btnCancel.Size      = New-Object System.Drawing.Size(100, 34)')
    $lines.Add('$btnCancel.Font      = New-Object System.Drawing.Font("Segoe UI", 10)')
    $lines.Add('$btnCancel.FlatStyle = "Flat"')
    $lines.Add('$btnCancel.Add_Click({ $form.Close() })')
    $lines.Add('$form.Controls.Add($btnCancel)')
    $lines.Add('')
    $lines.Add('$btnSync           = New-Object System.Windows.Forms.Button')
    $lines.Add('$btnSync.Text      = "' + $g_BtnSync + '"')
    $lines.Add('$btnSync.Location  = New-Object System.Drawing.Point(400, 478)')
    $lines.Add('$btnSync.Size      = New-Object System.Drawing.Size(200, 34)')
    $lines.Add('$btnSync.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)')
    $lines.Add('$btnSync.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)')
    $lines.Add('$btnSync.ForeColor = [System.Drawing.Color]::White')
    $lines.Add('$btnSync.FlatStyle = "Flat"')
    $lines.Add('$btnSync.Add_Click({')
    $lines.Add('    $selected = $checkboxes | Where-Object { $_.Checked }')
    $lines.Add('    if ($selected.Count -eq 0) {')
    $lines.Add('        [System.Windows.Forms.MessageBox]::Show("' + $g_WarnSelectOne + '", "' + $g_WarnTitle + '", "OK", "Warning")')
    $lines.Add('        return')
    $lines.Add('    }')
    $lines.Add('    foreach ($cb in $selected) {')
    $lines.Add('        Start-Process $cb.Tag')
    $lines.Add('        Start-Sleep -Milliseconds 500')
    $lines.Add('    }')
    $lines.Add('    $form.Close()')
    $lines.Add('})')
    $lines.Add('$form.Controls.Add($btnSync)')
    $lines.Add('')
    $lines.Add('$form.Add_Resize({')
    $lines.Add('    $panel.Width    = $form.ClientSize.Width - 40')
    $lines.Add('    $panel.Height   = $form.ClientSize.Height - 165')
    $lines.Add('    $btnAll.Top     = $form.ClientSize.Height - 75')
    $lines.Add('    $btnNone.Top    = $form.ClientSize.Height - 75')
    $lines.Add('    $btnCancel.Top  = $form.ClientSize.Height - 77')
    $lines.Add('    $btnCancel.Left = $form.ClientSize.Width - 310')
    $lines.Add('    $btnSync.Top    = $form.ClientSize.Height - 77')
    $lines.Add('    $btnSync.Left   = $form.ClientSize.Width - 220')
    $lines.Add('})')
    $lines.Add('')
    $lines.Add('$form.ShowDialog() | Out-Null')

    [System.IO.File]::WriteAllLines($ps1Path, $lines, [System.Text.UTF8Encoding]::new($false))
    Write-Host ($i18n.Script2Generated -f $ps1Path) -ForegroundColor Gray
    #endregion

    #region ETAPE 8 : Generer le raccourci .lnk
    # Les arguments embarquent un bloc PowerShell inline qui teste successivement
    # $env:OneDrive puis $env:OneDriveCommercial pour localiser le ps1 cible,
    # sans aucun chemin absolu ni fichier intermediaire.
    $subPath = "Documents\Backup Sharepoint\datas\Resynchroniser_SharePoint.ps1"
    $lnkArgs = "-ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command " +
               "& { " +
               "`$sub = '$subPath'; " +
               "`$candidates = @(`$env:OneDrive, `$env:OneDriveCommercial) | Where-Object { `$_ -and (Test-Path (Join-Path `$_ `$sub)) }; " +
               "if (`$candidates) { & (Join-Path `$candidates[0] `$sub) } " +
               "else { [System.Windows.Forms.MessageBox]::Show('Impossible de localiser le script de restauration. Verifiez que OneDrive est connecte.', 'Erreur', 'OK', 'Error') }" +
               " }"

    $lnkPath  = Join-Path $exportFolder "restaurer les sharepoint.lnk"
    $shell    = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($lnkPath)
    $shortcut.TargetPath       = "powershell.exe"
    $shortcut.Arguments        = $lnkArgs
    $shortcut.WorkingDirectory = ""
    $shortcut.IconLocation     = "shell32.dll,132"
    $shortcut.Save()
    Write-Host ($i18n.LnkGenerated -f $lnkPath) -ForegroundColor Gray
    #endregion

    #region ETAPE 9 : Masquer le dossier datas et son contenu
    $datasInfo = Get-Item $datasFolder
    $datasInfo.Attributes = $datasInfo.Attributes -bor [System.IO.FileAttributes]::Hidden
    Get-ChildItem $datasFolder | ForEach-Object {
        $_.Attributes = $_.Attributes -bor [System.IO.FileAttributes]::Hidden
    }
    Write-Host $i18n.DatasHidden -ForegroundColor Gray
    #endregion

    Remove-Item $tempDir -Recurse -Force

    $syncedCount = ($links | Where-Object { $_.IsSynced }).Count
    [System.Windows.Forms.MessageBox]::Show(
        ($i18n.SuccessMsg -f $syncedCount, $exportFolder),
        $i18n.SuccessTitle,
        "OK",
        "Information"
    )

} catch {
    Write-Host "`nERREUR : $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ($i18n.ErrorLine   -f $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Yellow
    Write-Host ($i18n.ErrorDetail -f $_.InvocationInfo.Line.Trim()) -ForegroundColor Yellow
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
    Read-Host ("`n" + $i18n.PressEnter)
}
