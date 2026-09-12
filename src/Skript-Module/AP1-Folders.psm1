# AP1-Folders.psm1
function New-CandidateFoldersFromExcel {
	param([Parameter(Mandatory)] [string]$WorkbookPath, [int]$MaxRows = 500)
	$rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
	New-EnsuredPath $rootPath
	$currentUser = [Environment]::UserName.Trim()
	Get-ChildItem -Path $rootPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
	$excel = $null; $wb = $null
	try {
		$excel = New-Object -ComObject Excel.Application -ErrorAction Stop
		$excel.Visible = $false; $excel.DisplayAlerts = $false
		$wb = $excel.Workbooks.Open($WorkbookPath)
		$sheet = $wb.Sheets.Item(1)
		$used = $sheet.UsedRange
		$rowCount = [Math]::Min($used.Rows.Count, $MaxRows)
		Write-Info "Suche Benutzer '$currentUser' in Excel (max. $rowCount Zeilen)..."
		for ($r = 1; $r -le $rowCount; $r++) {
			$a = $sheet.Cells.Item($r,1).Text
			$b = $sheet.Cells.Item($r,2).Text
			if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
			if (-not [string]::Equals($a.Trim(), $currentUser, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
			$safeA = ($a -replace '[\\/:*?"<>|]', '_').Trim()
			$safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
			if ([string]::IsNullOrWhiteSpace($safeB)) { continue }
			$path = Join-Path $rootPath $safeB
			New-EnsuredPath $path
			Write-Info "Benutzer '$currentUser' gefunden; Ordner '$safeB' wird bereitgestellt."
			break
		}
		Write-Info "Ordner fuer Pruefungskandidaten wurden angelegt."
		return $rootPath
	} catch {
		Write-Warning "Excel-COM-Lesezugriff fehlgeschlagen: $($_.Exception.Message) - nutze Python/pyopenxl-Fallback."
	} finally {
		if ($wb) { 
			try { $wb.Close($false) | Out-Null } catch {}
			Clear-ComObject $wb
		}
		if ($excel) { 
			try { $excel.DisplayAlerts = $true } catch {}
			Clear-ComObject $excel
		}
		Stop-NamedProcess EXCEL
	}

	$pythonCommand = $null
	foreach ($candidate in @('python', 'py')) {
		$cmd = Get-Command $candidate -ErrorAction SilentlyContinue
		if ($cmd) {
			$pythonCommand = $candidate
			break
		}
	}
	if (-not $pythonCommand) {
		throw "Excel-Datei konnte nicht verarbeitet werden: Weder COM noch Python/py verfügbar."
	}

	try {
		& $pythonCommand -c "import openpyxl" | Out-Null
		if ($LASTEXITCODE -ne 0) { throw "missing openpyxl" }
	} catch {
		Write-Info "openpyxl fehlt; installiere Python-Abhaengigkeit..."
		& $pythonCommand -m pip install openpyxl
		if ($LASTEXITCODE -ne 0) {
			throw "openpyxl konnte nicht installiert werden."
		}
	}

	$pythonScript = @"
import os, re, sys
from pathlib import Path
import openpyxl

wb_path = Path(r'$WorkbookPath')
root_path = Path(r'$rootPath')
root_path.mkdir(parents=True, exist_ok=True)
for child in root_path.iterdir():
	if child.is_dir():
		import shutil
		shutil.rmtree(child)
	else:
		child.unlink()
current_user = os.environ.get('USERNAME', '').strip().casefold()

wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
ws = wb.worksheets[0]
for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, $MaxRows), values_only=True):
    if len(row) < 2:
        continue
    a = '' if row[0] is None else str(row[0]).strip()
    b = '' if row[1] is None else str(row[1]).strip()
	if not a or not b or a.casefold() != current_user:
        continue
    safe_b = re.sub(r'[\\/:*?\"<>|]', '_', b).strip()
	if not safe_b:
        continue
	(root_path / safe_b).mkdir(parents=True, exist_ok=True)
	break

wb.close()
print(f'Python-Excel-Fallback: Ordner angelegt unter {root_path}')
"@

	try {
		$null = & $pythonCommand -c $pythonScript
		if ($LASTEXITCODE -ne 0) {
			throw "Python-Excel-Fallback fehlgeschlagen (ExitCode $LASTEXITCODE)."
		}
		Write-Info "Ordner fuer Pruefungskandidaten wurden mit Python-Fallback angelegt."
		return $rootPath
	} catch {
		throw "Excel-Datei konnte weder mit COM noch mit Python-Fallback verarbeitet werden: $($_.Exception.Message)"
	}
}
function New-CandidateFoldersFromCsv {
	param([Parameter(Mandatory)][string]$CsvPath, [int]$MaxRows=500)
	if (-not (Test-Path $CsvPath)) { throw "CSV nicht gefunden: $CsvPath" }
	$rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
	New-EnsuredPath $rootPath
	$currentUser = [Environment]::UserName.Trim()
	Get-ChildItem -Path $rootPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
	$rows = Import-Csv -Path $CsvPath -Delimiter ';' -Header 'Account','Kandidat'
	$i = 0
	foreach ($row in $rows) {
		if ($i -ge $MaxRows) { break }
		$a = $row.Account; $b = $row.Kandidat
		if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
		if (-not [string]::Equals($a.Trim(), $currentUser, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
		$safeB = ($b -replace '[\\/:*?"<>|]', '_').Trim()
		if ([string]::IsNullOrWhiteSpace($safeB)) { continue }
		$path = Join-Path $rootPath $safeB
		New-EnsuredPath $path
		Write-Info "Benutzer '$currentUser' gefunden; Ordner '$safeB' wird bereitgestellt."
		$i++
		break
	}
	Write-Info "Ordner aus CSV angelegt (${i}) Zeilen."
	return $rootPath
}
function Copy-CandidateFolderToDesktop {
	param([string]$SourceRoot)
	$desktop = Get-DesktopPath
	if (-not (Test-Path $SourceRoot)) {
		Write-Warning "Quellordner fuer Kandidaten existiert nicht: $SourceRoot"
		return
	}

	$participantFolders = Get-ChildItem -Path $SourceRoot -Directory -ErrorAction SilentlyContinue
	if (-not $participantFolders) {
		Write-Warning "Keine Kandidatenordner im Quellverzeichnis gefunden: $SourceRoot"
		return
	}

	foreach ($participantFolder in $participantFolders) {
		$target = Join-Path $desktop $participantFolder.Name
		try {
			if (Test-Path $target) {
				Remove-Item $target -Recurse -Force -ErrorAction Stop
			}
			Copy-Item -Path $participantFolder.FullName -Destination $target -Recurse -Force
			Write-Info "Kandidaten-Ordner auf Desktop bereitgestellt: $target"
			return $target
		} catch {
			Write-Warning "Kandidaten-Ordner konnte nicht kopiert werden: $($_.Exception.Message)"
		}
	}
	return $null
}
function New-EnsuredPath {
	param([string]$Path)
	$Path = $Path -replace 'NÃ¼', 'Nü' -replace 'Ã¤', 'ä' -replace 'Ã¶', 'ö' -replace 'ÃÖ', 'Ö' -replace 'ÃŸ', 'ß'
	if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}
function Join-PathSafe {
	param([string]$Path, [string]$ChildPath)
	$result = Join-Path -Path $Path -ChildPath $ChildPath
	$result = $result -replace 'NÃ¼', 'Nü' -replace 'Ã¤', 'ä' -replace 'Ã¶', 'ö' -replace 'ÃÖ', 'Ö' -replace 'ÃŸ', 'ß'
	return $result
}
function Get-DesktopPath { [Environment]::GetFolderPath('Desktop') }

# Alle Funktionsdefinitionen bleiben unverändert
Export-ModuleMember -Function New-CandidateFoldersFromExcel,New-CandidateFoldersFromCsv,Copy-CandidateFolderToDesktop,New-EnsuredPath,Join-PathSafe,Get-DesktopPath
