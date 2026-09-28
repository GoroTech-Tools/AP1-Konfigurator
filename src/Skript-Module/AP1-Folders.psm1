# AP1-Folders.psm1
function New-CandidateFoldersFromExcel {
	param([Parameter(Mandatory)] [string]$WorkbookPath, [int]$MaxRows = 500)
	$rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
	New-EnsuredPath $rootPath
	$currentUser = [Environment]::UserName.Trim()
	Get-ChildItem -Path $rootPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
	Remove-EmptyCandidateRoot -RootPath $rootPath
	New-EnsuredPath $rootPath
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
		Write-Warning "Excel-COM-Lesezugriff fehlgeschlagen: $($_.Exception.Message) - nutze Python-Fallback."
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

	$pythonScript = @"
import os, re, sys, zipfile, posixpath
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    import openpyxl
except ImportError:
    openpyxl = None

MAIN_NS = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
PACKAGE_NS = {'pkg': 'http://schemas.openxmlformats.org/package/2006/relationships'}
REL_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
NON_RECOVERABLE_ERRORS = (FileNotFoundError, IsADirectoryError, NotADirectoryError, PermissionError)

def load_shared_strings(archive):
    try:
        root = ET.fromstring(archive.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    values = []
    for item in root.findall('main:si', MAIN_NS):
        values.append(''.join(node.text or '' for node in item.findall('.//main:t', MAIN_NS)))
    return values

def resolve_first_sheet_path(archive):
    workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    relationships = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    first_sheet = workbook.find('main:sheets/main:sheet', MAIN_NS)
    if first_sheet is None:
        raise RuntimeError('Keine Tabellenblaetter in der Excel-Datei gefunden.')
    rel_id = first_sheet.attrib.get(REL_NS)
    for relation in relationships.findall('pkg:Relationship', PACKAGE_NS):
        if relation.attrib.get('Id') != rel_id:
            continue
        target = relation.attrib.get('Target', '')
        if not target:
            break
        if target.startswith('/'):
            return target.lstrip('/')
        return posixpath.normpath(posixpath.join('xl', target))
    raise RuntimeError('Erstes Tabellenblatt konnte in der Excel-Datei nicht aufgeloest werden.')

def column_index(cell_reference):
    letters = ''.join(ch for ch in cell_reference if ch.isalpha()).upper()
    value = 0
    for letter in letters:
        value = (value * 26) + (ord(letter) - ord('A') + 1)
    return value

def cell_text(cell, shared_strings):
    cell_type = cell.attrib.get('t')
    if cell_type == 'inlineStr':
        return ''.join(node.text or '' for node in cell.findall('.//main:t', MAIN_NS))
    value_node = cell.find('main:v', MAIN_NS)
    if value_node is None or value_node.text is None:
        return ''
    value = value_node.text
    if cell_type == 's':
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return ''
    return value

def iter_rows_with_zip(workbook_path, max_rows):
    with zipfile.ZipFile(workbook_path) as archive:
        shared_strings = load_shared_strings(archive)
        sheet_path = resolve_first_sheet_path(archive)
        worksheet = ET.fromstring(archive.read(sheet_path))
    row_count = 0
    for row in worksheet.findall('.//main:sheetData/main:row', MAIN_NS):
        row_count += 1
        if row_count > max_rows:
            break
        first_value = ''
        second_value = ''
        for cell in row.findall('main:c', MAIN_NS):
            index = column_index(cell.attrib.get('r', ''))
            if index == 1:
                first_value = cell_text(cell, shared_strings)
            elif index == 2:
                second_value = cell_text(cell, shared_strings)
        yield first_value, second_value

def iter_rows(workbook_path, max_rows):
    if openpyxl is not None:
        try:
            wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
            try:
                ws = wb.worksheets[0]
                for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True):
                    if len(row) < 2:
                        continue
                    yield row[0], row[1]
            finally:
                wb.close()
            return
        except Exception as exc:
            if isinstance(exc, NON_RECOVERABLE_ERRORS):
                raise
            print(f'openpyxl-Lesezugriff fehlgeschlagen, nutze ZIP/XML-Fallback: {exc}', file=sys.stderr)
    yield from iter_rows_with_zip(workbook_path, max_rows)

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

for first_value, second_value in iter_rows(wb_path, $MaxRows):
    a = '' if first_value is None else str(first_value).strip()
    b = '' if second_value is None else str(second_value).strip()
    if not a or not b or a.casefold() != current_user:
        continue
    safe_b = re.sub(r'[\\/:*?\"<>|]', '_', b).strip()
    if not safe_b:
        continue
    (root_path / safe_b).mkdir(parents=True, exist_ok=True)
    break

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
	Remove-EmptyCandidateRoot -RootPath $rootPath
	New-EnsuredPath $rootPath
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
function Remove-EmptyCandidateRoot {
	param([string]$RootPath)
	if (-not $RootPath) { return }
	if (Test-Path $RootPath) {
		$childItems = @(Get-ChildItem -Path $RootPath -Force -ErrorAction SilentlyContinue)
		if ($childItems.Count -eq 0) {
			Remove-Item -Path $RootPath -Recurse -Force -ErrorAction SilentlyContinue
		}
	}
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
Export-ModuleMember -Function New-CandidateFoldersFromExcel,New-CandidateFoldersFromCsv,Copy-CandidateFolderToDesktop,Remove-EmptyCandidateRoot,New-EnsuredPath,Join-PathSafe,Get-DesktopPath
