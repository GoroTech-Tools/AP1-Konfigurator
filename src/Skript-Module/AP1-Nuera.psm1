# AP1-Nuera.psm1
function Get-LatestNueraCandidate {
    param(
        [string[]]$FileNames = @(
            "nuera2026_h.zip",
            "nuera2026_f.zip",
            "nuera2025_h.zip",
            "nuera2025_f.zip",
            "nuera2024_h.zip",
            "nuera2024_f.zip"
        )
    )

    $ordered = $FileNames | ForEach-Object {
        $match = [regex]::Match($_, 'nuera(?<year>\d{4})_(?<kind>[fh])\.zip$', 'IgnoreCase')
        if (-not $match.Success) {
            return [pscustomobject]@{ Name = $_; Year = 0; KindOrder = 0 }
        }

        $kind = $match.Groups['kind'].Value.ToLowerInvariant()
        # Bei gleichem Jahr ist die Herbstversion (h) die aktuellere Ausgabe.
        $kindOrder = if ($kind -eq 'h') { 2 } else { 1 }
        [pscustomobject]@{ Name = $_; Year = [int]$match.Groups['year'].Value; KindOrder = $kindOrder }
    } | Sort-Object @{ Expression = 'Year'; Descending = $true }, @{ Expression = 'KindOrder'; Descending = $true }

    return @($ordered | Select-Object -ExpandProperty Name)
}

function Get-LatestNueraFile {
    param([string]$DownloadPath = (Join-Path $script:ScriptRoot $script:NueraFolderName))
    Write-Info '[DEBUG] Get-LatestNueraFile: Funktionsstart'
    $BaseUrl = "https://www.ihk-aka.de/fileadmin/AkA/Download/Nuera/"

    $headers = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell' }
    Write-Info "[DEBUG] DownloadPath: $DownloadPath"
    if (-not (Test-Path $DownloadPath)) {
        Write-Info "[DEBUG] DownloadPath existiert NICHT!"
    }

    New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
    $candidateNames = Get-LatestNueraCandidate
    foreach ($fileName in $candidateNames) {
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $folderPath = Join-Path $DownloadPath $folderName
        $zipPath = Join-Path $DownloadPath $fileName
        if ((Test-Path $folderPath) -and (Get-ChildItem $folderPath -Recurse -File -ErrorAction SilentlyContinue)) {
            Write-Info "Verwende lokale Nuera-Datei: $fileName"
            Get-ChildItem -Path $DownloadPath -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '^nuera\d{4}_[fh](\.zip)?$' -and $_.Name -notin @($fileName, $folderName) } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            return $folderPath
        }
        if (Test-Path $zipPath) {
            if (Test-ExpandArchive -ZipPath $zipPath -Destination $DownloadPath) {
                if (Test-Path $folderPath) {
                    Write-Info "Entpacke lokale Nuera-Datei erneut: $fileName"
                    return $folderPath
                }
            }
        }
    }
    Get-ChildItem -Path $DownloadPath -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^nuera\d{4}_[fh](\.zip)?$' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Info 'Alte Nüra-Dateien und entpackte Ordner wurden entfernt.'

    $fileNames = $candidateNames
    $found = $null
    foreach ($fileName in $fileNames) {
        $url = "$BaseUrl$fileName"
        $zipPath = Join-Path $DownloadPath $fileName
        try {
            Write-Info "Pruefe und lade: $url"
            Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
            $found = $fileName
            break
        } catch {
            Write-Warning "Nicht gefunden oder Download fehlgeschlagen: $url ($($_.Exception.Message))"
        }
    }
    if (-not $found) {
        Write-Warning 'Keine Nuera-Datei gefunden.'
        return $null
    }
    if (Test-ExpandArchive -ZipPath (Join-Path $DownloadPath $found) -Destination $DownloadPath) {
        Write-Info "Nuera-Dateien erfolgreich in '$DownloadPath' extrahiert"
        $folderName = [System.IO.Path]::GetFileNameWithoutExtension($found)
        $extractedPath = Join-Path $DownloadPath $folderName
        if (Test-Path $extractedPath) {
            return $extractedPath
        } else {
            Write-Warning "Entpackter Ordner nicht gefunden: $extractedPath"
            return $null
        }
    } else {
        Write-Warning "Extraktion der Nuera-Dateien fehlgeschlagen"
        return $null
    }
}

function Copy-NueraToDesktop {
    param([string]$NueraSourcePath, [string]$ZipFileName)
    if (-not (Test-Path $NueraSourcePath)) {
        Write-Error "Kein Nuera-Ordner gefunden: $NueraSourcePath"
        return $null
    }
    $desktop = Get-DesktopPath
    $folderName = (Split-Path $NueraSourcePath -Leaf)
    $target = Join-Path $desktop $folderName
    Write-Info "Ziel-Desktop-Pfad: $target"
    try {
        if (Test-Path $target) {
            Remove-Item $target -Recurse -Force -ErrorAction Stop
        }
        # Kopiere den Ordner direkt auf den Desktop
        Copy-Item -Path $NueraSourcePath -Destination $target -Recurse -Force -ErrorAction Stop
        Write-Info "Auf Desktop kopiert: $target"
        return $target
    } catch {
        Write-Error "Kopieren auf Desktop fehlgeschlagen: $_"
        return $null
    }
}

Export-ModuleMember -Function Get-LatestNueraCandidate,Get-LatestNueraFile,Copy-NueraToDesktop
