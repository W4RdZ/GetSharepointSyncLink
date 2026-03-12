# ===========================
# Script : Générer liens SharePoint synchronisés via OneDrive
# ===========================

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }

#region ETAPE 0 : Copie des fichiers dans un dossier temporaire
$settingsRoot = Join-Path $env:LOCALAPPDATA "Microsoft\OneDrive\settings"
$tempDir      = Join-Path $scriptDir ("IniTemp_" + [guid]::NewGuid())
New-Item -ItemType Directory -Path $tempDir | Out-Null

$businessFolders = Get-ChildItem -Path $settingsRoot -Directory -Filter "Business*" -ErrorAction SilentlyContinue
if (-not $businessFolders) {
    Write-Warning "Aucun dossier Business* trouvé dans $settingsRoot"
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

if (-not $copiedFiles) {
    Write-Warning "Aucun .ini trouvé dans les dossiers Business*"
    Remove-Item $tempDir -Recurse -Force
    exit
}
#endregion

#region ETAPE 1 : Trouver le fichier GUID d'état
$guidPattern = '^Business\d+_[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\.ini$'

$stateFile = $copiedFiles |
    Where-Object { $_.Name -match $guidPattern } |
    Where-Object { (Get-Content $_.FullName -Encoding Unicode) -match "libraryScope|libraryFolder" } |
    Select-Object -First 1 -ExpandProperty FullName

if (-not $stateFile) {
    Write-Warning "Fichier d'état introuvable"
    Remove-Item $tempDir -Recurse -Force
    exit
}

Write-Host "Fichier d'état trouvé : $stateFile" -ForegroundColor Green
#endregion

#region ETAPE 2 : Pré-charger les policy files — SiteTitle -> ViewOnlineUrlTemplate
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

Write-Host "$($templateBySiteTitle.Count) policy file(s) chargé(s)" -ForegroundColor Gray
#endregion

#region ETAPE 3 : Parser libraryScope et libraryFolder
$allLines = Get-Content $stateFile -Encoding Unicode

# --- libraryScope ---
$scopeByIndex = @{}

$allLines | Where-Object { $_ -match '^\s*libraryScope\s*=' } | ForEach-Object {
    $line       = $_
    $scopeIndex = $null
    if ($line -match '=\s*(\d+)\s+') { $scopeIndex = [int]$Matches[1] }

    # Extraire le resourceId : 2e token après = (hex pur avant le +n éventuel)
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

# --- libraryFolder ---
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
#endregion

#region ETAPE 4 : Construire les liens finaux
$links = [System.Collections.Generic.List[PSCustomObject]]::new()

# Scope indexes qui ont au moins un libraryFolder associé
$scopeIndexesWithFolder = $folders | Select-Object -ExpandProperty ScopeIndex -Unique

# Fonction helper partagée pour résoudre un lien via ViewOnlineUrlTemplate
function Resolve-SharePointUrl {
    param($SiteTitle, $ResourceId)

    $template = $templateBySiteTitle[$SiteTitle]
    if (-not $template) {
        Write-Warning "Policy introuvable pour SiteTitle='$SiteTitle'"
        return $null
    }
    if ($template -notmatch [Regex]::Escape('{ResourceID}')) {
        Write-Warning "Pas de {ResourceID} dans le template pour '$SiteTitle'"
        return $null
    }
    return $template -replace [Regex]::Escape('{ResourceID}'), $ResourceId
}

# -- Bibliothèques racines (libraryScope) --
# Exclure : OneDrive personnel (/personal/) et ceux qui ont un libraryFolder plus précis
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

# -- Sous-dossiers synchronisés (libraryFolder) --
foreach ($folder in $folders) {
    $parentScope = $scopeByIndex[$folder.ScopeIndex]
    if (-not $parentScope) {
        Write-Warning "Scope parent introuvable pour '$($folder.FolderName)' (scopeIndex=$($folder.ScopeIndex))"
        continue
    }

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

Write-Host "`n$($links.Count) lien(s) total ($($scopeIndexesWithFolder.Count) dossier(s), $($scopeByIndex.Count - $scopeIndexesWithFolder.Count) bibliotheque(s) racine)" -ForegroundColor Green
#endregion

#region ETAPE 5 : Générer le fichier batch
if ($links.Count -gt 0) {
    $batchPath  = Join-Path $scriptDir "OpenAllLinks.bat"
    $batchLines = [System.Collections.Generic.List[string]]::new()
    $batchLines.Add("@echo off")
    $batchLines.Add("REM Bibliotheques et dossiers SharePoint synchronises via OneDrive")
    $batchLines.Add("")

    foreach ($l in $links) {
        $batchLines.Add("REM [$($l.Type)] $($l.SiteTitle) - $($l.Name)")
        $batchLines.Add("start `"`" `"$($l.Url)`"")
        $batchLines.Add("")
    }

    [System.IO.File]::WriteAllLines($batchPath, $batchLines, [System.Text.UTF8Encoding]::new($false))
    Write-Host "Batch généré : $batchPath" -ForegroundColor Green
} else {
    Write-Warning "Aucun lien généré, batch non créé."
}
#endregion

# Nettoyage
Remove-Item $tempDir -Recurse -Force
