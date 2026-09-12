param(
    [switch]$NoVersionBump,
    [switch]$SkipZip,
    [switch]$Help
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $OutputEncoding = [System.Text.Encoding]::UTF8
}

if ($Help) {
    Write-Host 'AP1-Konfigurator Build' -ForegroundColor Cyan
    Write-Host '  .\src\build.ps1'
    Write-Host '  .\src\build.ps1 -NoVersionBump'
    Write-Host '  .\src\build.ps1 -SkipZip'
    exit 0
}

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

$pythonExe = Join-Path $ProjectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $pythonExe)) {
    throw 'Python-Umgebung fehlt. Bitte zuerst .\src\setup.ps1 ausführen.'
}

$buildInfoPath = Join-Path $ProjectRoot 'src\build_info.py'
$content = Get-Content $buildInfoPath -Raw -Encoding UTF8
if ($content -notmatch "'version':\s*'([0-9]+)\.([0-9]+)\.([0-9]+)'") {
    throw 'Versionsnummer in src/build_info.py nicht gefunden.'
}

$major = [int]$Matches[1]
$minor = [int]$Matches[2]
$patch = [int]$Matches[3]
if (-not $NoVersionBump) {
    $patch++
}
$newVersion = "$major.$minor.$patch"
$newDate = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
$content = $content -replace "'version':\s*'[0-9]+\.[0-9]+\.[0-9]+'", "'version': '$newVersion'"
$content = $content -replace "'build_date':\s*'[^']+'", "'build_date': '$newDate'"
Set-Content -Path $buildInfoPath -Value $content -Encoding UTF8
Write-Host "Version: $newVersion" -ForegroundColor Green

& $pythonExe (Join-Path $ProjectRoot 'src\update_docs_versions.py')
if ($LASTEXITCODE -ne 0) { throw 'Doku-Versionsaktualisierung fehlgeschlagen.' }

& $pythonExe (Join-Path $ProjectRoot 'src\fix_markdown.py')
if ($LASTEXITCODE -ne 0) { throw 'Markdown-Normalisierung fehlgeschlagen.' }

$distRoot = Join-Path $ProjectRoot 'dist'
$buildRoot = Join-Path $ProjectRoot 'build'
$releaseRoot = Join-Path $ProjectRoot 'release'
$exeName = 'AP1-Konfigurator'
$iconPath = Join-Path $ProjectRoot 'src\app_icon.ico'
$versionTag = "v$newVersion"

$pyInstallerArgs = @(
    '-m', 'PyInstaller',
    '--noconfirm',
    '--onefile',
    '--noconsole',
    '--name', $exeName,
    '--icon', $iconPath,
    '--specpath', (Join-Path $ProjectRoot 'src'),
    '--distpath', $distRoot,
    '--workpath', $buildRoot,
    '--add-data', ((Join-Path $ProjectRoot 'src\AP1-Konfigurator.ps1') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\AP1-Konfigurator.bat') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\Proxy-Deaktivieren.bat') + ';.'),
    '--add-data', ($iconPath + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\build_info.py') + ';.'),
    '--add-data', ((Join-Path $ProjectRoot 'src\Skript-Module') + ';Skript-Module'),
    '--add-data', ((Join-Path $ProjectRoot 'data') + ';data'),
    '--add-data', ((Join-Path $ProjectRoot 'docs') + ';docs'),
    (Join-Path $ProjectRoot 'src\main.py')
)

& $pythonExe @pyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller fehlgeschlagen.' }

& $pythonExe (Join-Path $ProjectRoot 'src\post_build.py')
if ($LASTEXITCODE -ne 0) { throw 'Post-Build-Paketierung fehlgeschlagen.' }

if (-not (Test-Path $releaseRoot)) {
    New-Item -Path $releaseRoot -ItemType Directory | Out-Null
}

$releaseNotesPath = Join-Path $releaseRoot ("RELEASE_NOTES_{0}.md" -f $versionTag)
$releaseNotesTemplatePath = Join-Path $PSScriptRoot 'RELEASE_NOTES_TEMPLATE.md'
$releaseZipPath = Join-Path $releaseRoot ("{0}-{1}.zip" -f $exeName, $versionTag)

if (-not (Test-Path $releaseNotesTemplatePath)) {
    throw "Release-Notes-Template fehlt: $releaseNotesTemplatePath"
}
$releaseNotes = Get-Content $releaseNotesTemplatePath -Raw -Encoding UTF8
$releaseNotes = $releaseNotes.Replace('vX.Y.Z', $versionTag).Replace('X.Y.Z', $newVersion)
$releaseNotes = $releaseNotes.Replace('AP1-Konfigurator-Portable-' + $versionTag, $exeName + '-' + $versionTag)
[System.IO.File]::WriteAllText($releaseNotesPath, $releaseNotes, [System.Text.UTF8Encoding]::new($false))
Write-Host "Release Notes erstellt/aktualisiert: $releaseNotesPath" -ForegroundColor Green

# Ältere Release-Artefakte aus dem Hauptordner in _Archiv verschieben.
$releaseArchive = Join-Path $releaseRoot '_Archiv'
if (-not (Test-Path $releaseArchive)) {
    New-Item -Path $releaseArchive -ItemType Directory -Force | Out-Null
}
$currentReleaseNames = @(
    "$exeName-$versionTag",
    "$exeName-$versionTag.zip",
    "RELEASE_NOTES_$versionTag.md"
)
$versionedReleasePattern = '^(AP1-Konfigurator(?:-Portable)?-v\d+\.\d+\.\d+(?:\.zip)?|RELEASE_NOTES_v\d+\.\d+\.\d+\.md)$'
$olderReleaseItems = @(Get-ChildItem -Path $releaseRoot -Force -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match $versionedReleasePattern -and $_.Name -notin $currentReleaseNames
    })
foreach ($releaseItem in $olderReleaseItems) {
    $archiveTarget = Join-Path $releaseArchive $releaseItem.Name
    if (Test-Path $archiveTarget) {
        Remove-Item -LiteralPath $archiveTarget -Recurse -Force -ErrorAction SilentlyContinue
    }
    Move-Item -LiteralPath $releaseItem.FullName -Destination $archiveTarget -Force -ErrorAction SilentlyContinue
    if (Test-Path $archiveTarget) {
        Write-Host "Älteres Release archiviert: $($releaseItem.Name)" -ForegroundColor DarkGray
    } else {
        Write-Warning "Älteres Release konnte nicht archiviert werden: $($releaseItem.Name)"
    }
}

if ($SkipZip) {
    Write-Host 'Hinweis: -SkipZip ist im aktuellen Buildablauf ohne Wirkung, da die Paketierung im Post-Build erfolgt.' -ForegroundColor Yellow
}

Write-Host 'Build abgeschlossen.' -ForegroundColor Green
