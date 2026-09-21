param(
  [ValidateSet('On','Off','Skip')]
  [string]$Proxy = 'Skip',
  [string]$ProxyServer = '192.168.0.1:8080',
  [string]$ProxyBypass = '*.office365.com; *.cloudappsecurity.com; *.onmicrosoft.com; *.office.net; *.office.com; *.microsoft.com; *.microsoftonline.com; *.live.com; *.azure.net; *.gfx.ms; *.onestore.ms; *.msecnd.net; *.outlookgroups.ms; *.linkedin.com; *.msocdn.com; *.live.net; ihk-aka.de',
  [string]$ExcelListPath = '',
  [string]$CsvFallbackPath = '',
  [int]$MaxRows = 500,
  [switch]$Quiet,
  [switch]$RegistryOnly,
  [switch]$UseCom
)

$script:AppVersion = '1.0.30'

if ($PSScriptRoot) {
  $script:BaseRoot = $PSScriptRoot
} else {
  $script:BaseRoot = (Get-Location).Path
}

$dataRootCandidate = Join-Path $script:BaseRoot 'data'
if (Test-Path $dataRootCandidate) {
  $script:DataRoot = $dataRootCandidate
} else {
  $parentDataRootCandidate = Join-Path (Split-Path -Path $script:BaseRoot -Parent) 'data'
  if (Test-Path $parentDataRootCandidate) {
    $script:DataRoot = $parentDataRootCandidate
  } else {
    $script:DataRoot = $script:BaseRoot
  }
}

### --- Office-Version 16.0 für alle 2026 laufenden Versionen ---
if (-not $script:OfficeVersion) { $script:OfficeVersion = '16.0' }

### --- Encoding für Umlaute und Sonderzeichen setzen ---
try {
  [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
  $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

# --- Module laden (direkt am Anfang, damit alle Funktionen global verfügbar sind) ---
try {
  $modulePath = Join-Path $script:BaseRoot 'Skript-Module'
  if (-not (Test-Path $modulePath)) {
    $modulePath = Join-Path $script:DataRoot 'Skript-Module'
  }
  Get-ChildItem -Path $modulePath -Filter '*.psm1' | ForEach-Object {
    Import-Module $_.FullName -Force
  }
} catch {
  Write-Host "[FEHLER] Modul konnte nicht geladen werden: $($_.Exception.Message)" -ForegroundColor Red
  exit 1
}

# Setzt PersonalTemplates für Word immer auf den Desktop des aktuellen Benutzers
try {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $wordRegPath = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
    Set-ItemProperty -Path $wordRegPath -Name "PersonalTemplates" -Value $desktop -Force
    Write-Info "PersonalTemplates für Word wurde auf den Desktop gesetzt: $desktop"
} catch {
    Write-Warning "Konnte PersonalTemplates nicht setzen: $($_.Exception.Message)"
}

# Optionaler Komfortschritt; der stabile Standardbetrieb benoetigt kein Shell-COM.
if ($UseCom) {
  try { Add-DesktopToQuickAccess } catch { Write-Warning "Schnellzugriff konnte nicht per COM gesetzt werden: $($_.Exception.Message)" }
} else {
  Write-Info "Registry-only-Modus aktiv: Office-/Shell-COM wird nicht verwendet."
}



# === Desktop-Verknüpfung ausblenden (optional) ===
# $shortcutName = Read-Host 'Name der auszublendenden Desktop-Verknüpfung (ohne .lnk)'
# $antwort = Read-Host "Soll die Verknüpfung '$shortcutName' wirklich ausgeblendet werden? (J/N)"
# if ($antwort -eq 'J') {
#     Hide-DesktopShortcut -ShortcutName $shortcutName
# } else {
#     Write-Host "Abgebrochen. Die Verknüpfung bleibt sichtbar." -ForegroundColor Yellow
# }

# ==================
# Region: Hauptlauf
# ==================

# ==========================================


function Set-OfficeFirstRunMarkers {
  # Setzt Marker, damit zukuenftige Laeufe nicht mehr als First-Run erkannt werden
  try {
    $wordRegPath = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
    $excelRegPath = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
    
    if (-not (Test-Path $wordRegPath)) { New-Item -Path $wordRegPath -Force | Out-Null }
    if (-not (Test-Path $excelRegPath)) { New-Item -Path $excelRegPath -Force | Out-Null }
    
    New-ItemProperty -Path $wordRegPath -Name "FirstRun" -Value 0 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $excelRegPath -Name "FirstRun" -Value 0 -PropertyType DWord -Force | Out-Null
    
    Write-Info "First-Run-Marker gesetzt"
  } catch {
    WriteWarn "Konnte First-Run-Marker nicht setzen: $($_.Exception.Message)"
  }
}

function Set-OfficeRegistrySettings {
  $regWord   = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
  $regExcel  = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
  $regWinAdv = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

  $wordSettings = @{
    'DeveloperTools'                   = 1
    'Ruler'                            = 1
    'ShowAllFormatting'                = 1
    'VisiDrawTableDrs'                 = 1
    'DisableBootToOfficeStart'         = 1
    'DisableBackstageOpenKeyShortcuts' = 1
  }
  $excelSettings = @{
    'DeveloperTools'                   = 1
    'DisableBootToOfficeStart'         = 1
  }
  $windowsSettings = @{ 'HideFileExt' = 0 }

  Write-Info "Setze Office-/Windows-Optionen (HKCU)..."
  Set-RegistryValues -RegPath $regWord   -Settings $wordSettings
  Set-RegistryValues -RegPath $regExcel  -Settings $excelSettings
  Set-RegistryValues -RegPath $regWinAdv -Settings $windowsSettings
}

function Set-DefaultSavePaths {
  param([string]$Path, [switch]$UseCom)
  if ([string]::IsNullOrWhiteSpace($Path)) {
    $Path = [Environment]::GetFolderPath('Desktop')
  }
  if (-not (Test-Path $Path)) {
    try {
      New-Item -ItemType Directory -Path $Path -Force | Out-Null
      Write-Info "[DEBUG] Set-DefaultSavePaths: Speicherordner wurde angelegt: $Path"
    } catch {
      Write-Info "[DEBUG] Set-DefaultSavePaths: Speicherordner konnte nicht angelegt werden: $($_.Exception.Message)"
    }
  }
  $regWord   = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
  $regExcel  = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options" -replace "\\\\", "\\")
  Write-Info "Setze Standard-Speicherpfade: $Path"
  Set-RegistryValues -RegPath $regWord  -Settings @{ 'DOC-PATH'   = $Path }
  Set-RegistryValues -RegPath $regExcel -Settings @{ 'DefaultPath' = $Path }

  if ($UseCom) {
    # Word COM
    Stop-NamedProcess WINWORD
    try {
      $word = New-Object -ComObject Word.Application -ErrorAction Stop
      $word.Visible = $false
      $wdDocumentsPath = 0 # Enum-Ersatz
      $word.Options.DefaultFilePath($wdDocumentsPath) = $Path
    } catch { WriteWarn "Word COM nicht gesetzt: $($_.Exception.Message)" }
    finally {
      Clear-ComObject $word
    }

    # Excel COM
    Stop-NamedProcess EXCEL
    try {
      $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
      $excel.Visible = $false
      $excel.DefaultFilePath = $Path
    } catch { WriteWarn "Excel COM nicht gesetzt: $($_.Exception.Message)" }
    finally {
      Clear-ComObject $excel
    }
  } else {
    Write-Info "COM uebersprungen (Registry-only)."
  }
}

function Set-WordTemplatesPath {
  param([string]$TemplatePath)
  
  $regWord = ("HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options" -replace "\\\\", "\\")
  Write-Info "Setze Word Standard-Speicherort für persönliche Vorlagen: $TemplatePath"
  
  # Korrekter Registry-Schlüssel für persönliche Vorlagen in modernen Office-Versionen
  $templateSettings = @{ 
    'PersonalTemplates' = $TemplatePath  # Moderner Office-Schlüssel für persönliche Vorlagen
    'USER-DOT-PATH' = $TemplatePath      # Fallback für ältere Versionen
  }
  
  Set-RegistryValues -RegPath $regWord -Settings $templateSettings
  
  # Zusätzlich auch über COM setzen (falls verfügbar)
  try {
    Stop-NamedProcess WINWORD
    $word = New-Object -ComObject Word.Application -ErrorAction Stop
    $word.Visible = $false
    
    # Setze beide Template-Pfade über COM
    $wdUserTemplatesPath = 1      # User Templates Path
    $wdWorkgroupTemplatesPath = 2 # Workgroup Templates Path
    
    $word.Options.DefaultFilePath($wdUserTemplatesPath) = $TemplatePath
    $word.Options.DefaultFilePath($wdWorkgroupTemplatesPath) = $TemplatePath
    
    Write-Info "Word Vorlagen-Pfade auch über COM gesetzt (User + Workgroup)."
  } catch { 
    WriteWarn "Word COM für Vorlagen-Pfad nicht verfügbar: $($_.Exception.Message)" 
  } finally {
    Clear-ComObject $word
  }
  
  Write-Info "Hinweis: Word muss neu gestartet werden, um die neuen Vorlagen-Pfade anzuzeigen."
}

function Set-ExcelAutoCorrectRegistry {
  $regExcel = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
  $settings = @{ 'CorrectSentenceCap' = 0 }
  Write-Info "Setze Excel Autokorrektur (HKCU)..."
  Set-RegistryValues -RegPath $regExcel -Settings $settings
}

# ===================================
# Region: Vorlagen & Schnellzugriff
# ===================================
function Copy-QuickAccessToolbarFiles {
  if (-not $script:ScriptRoot) {
    $script:ScriptRoot = $script:DataRoot
  }
  $sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Symbolleiste Schnellzugriff'
  $targetPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office'
  New-EnsuredPath $targetPath
  foreach ($file in @('Excel.officeUI','Word.officeUI')) {
    $src = Join-Path $sourcePath $file
    $dst = Join-Path $targetPath $file
    if (Test-Path $src) {
      try {
        Stop-NamedProcess WINWORD; Stop-NamedProcess EXCEL
        Copy-Item -LiteralPath $src -Destination $dst -Force
        Write-Info "Schnellzugriff uebernommen: $file"
      } catch { WriteWarn "Schnellzugriff $file nicht kopiert: $($_.Exception.Message)" }
    } else {
      Write-Info "Schnellzugriff-Datei fehlt: $file"
    }
  }
}

function Copy-WordTemplate {
  $targetPath = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
  $sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Word\Normal.dotm'
  if (-not (Test-Path $sourcePath)) { Write-Info "Normal.dotm-Quelle fehlt ($sourcePath), ueberspringe."; return }
  New-EnsuredPath (Split-Path $targetPath -Parent)
  try {
    Stop-NamedProcess WINWORD
    if (Test-Path $targetPath) {
      Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Write-Info "Normal.dotm kopiert."
  } catch { WriteWarn "Normal.dotm konnte nicht kopiert werden: $($_.Exception.Message)" }
}

function Copy-ExcelTemplate {
  $targetPath = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART\Mappe.xltx'
  $sourcePath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Excel\Mappe.xltx'
  if (-not (Test-Path $sourcePath)) { Write-Info "Mappe.xltx-Quelle fehlt ($sourcePath), ueberspringe."; return }
  New-EnsuredPath (Split-Path $targetPath -Parent)
  try {
    Stop-NamedProcess EXCEL
    if (Test-Path $targetPath) {
      Copy-Item -LiteralPath $targetPath -Destination ($targetPath + '.bak') -Force -ErrorAction SilentlyContinue
    }
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Write-Info "Mappe.xltx kopiert."
  } catch { WriteWarn "Mappe.xltx konnte nicht kopiert werden: $($_.Exception.Message)" }
}

# ============================================
# Region: Ordnererzeugung & Desktop-Deployment
# ============================================
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

function New-CandidateFoldersFromExcel {
  param([Parameter(Mandatory)] [string]$WorkbookPath, [int]$MaxRows = 500, [switch]$NoCom)
  $rootPath = Join-Path $script:ScriptRoot '2. Bei Bedarf anpassen\Ordner'
  New-EnsuredPath $rootPath
  $currentUser = [Environment]::UserName.Trim()
  Get-ChildItem -Path $rootPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  Remove-EmptyCandidateRoot -RootPath $rootPath
  New-EnsuredPath $rootPath

  $excel = $null; $wb = $null
  try {
    if ($NoCom) { throw 'COM im Standardmodus deaktiviert' }
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
import os, re
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

  return $rootPath
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


# ===========================
# Region: Proxy & Taskbar
# ===========================

function Set-Proxy {
  param(
    [ValidateSet('On','Off','Skip')] [string]$State,
    [string]$Server,
    [string]$BypassList
  )
  if ($State -eq 'Skip') { Write-Info "Proxy-Konfiguration uebersprungen."; return }

  $reg = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
  $enable = if ($State -eq 'On') { 1 } else { 0 }

  # Nutzerabfrage (nur im interaktiven Modus)
  if (-not $Quiet) {
    $msg = "Proxy wird auf '$State' gesetzt (Server: $Server, Bypass: $BypassList). Fortfahren? [J/N]"
    $confirm = Read-Host $msg
    if ($confirm -notin @('J','j','Y','y','')) {
      Write-Info "Proxy-Konfiguration abgebrochen durch Nutzer."
      return
    }
  }

  Write-Info "Setze Proxy: $State"
  Set-RegistryValues -RegPath $reg -Settings @{
    'ProxyEnable'  = $enable
    'ProxyServer'  = $Server
    'ProxyOverride'= $BypassList
    'AutoDetect'   = 0
  }
}

function Set-TaskbarSettings {
  param([ValidateSet('Left','Center')] [string]$Alignment='Left', [ValidateSet('Hidden','Icon','Box')] [string]$Search='Icon')
  $regAdv = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
  $regSea = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'
  $taskbarValue = if ($Alignment -eq 'Left') { 0 } else { 1 }
  $searchValue = switch ($Search) { 'Hidden' {0} 'Icon' {1} 'Box' {2} }
  Write-Info "Setze Taskbar: Alignment=$Alignment, Search=$Search"
  Set-RegistryValues -RegPath $regAdv -Settings @{
    'TaskbarAl' = $taskbarValue
    'TaskbarGlomLevel' = 2
  }
  Set-RegistryValues -RegPath $regSea -Settings @{ 'SearchboxTaskbarMode' = $searchValue }

  try {
    $notify = Add-Type -Name Win32TaskbarRefresh -Namespace Win32 -PassThru -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr result);
'@
    $hwnd = [IntPtr]0xffff
    $msg = 0x001A
    $result = [IntPtr]::Zero
    $notify::SendMessageTimeout($hwnd, $msg, [IntPtr]::Zero, "TrayNotify", 0x0002, 1000, [ref]$result) | Out-Null
  } catch {
    Write-Info "Taskbar-Refresh ohne Explorer-Neustart übersprungen: $($_.Exception.Message)"
  }
}

# ==================
# Region: Hauptlauf
# ==================


function Start-AP1Konfiguration {
    param(
        [ValidateSet('On','Off','Skip')]
        [string]$Proxy = 'Skip',
        [string]$ProxyServer = '192.168.0.1:8080',
        [string]$ProxyBypass = '*.office365.com; *.cloudappsecurity.com; *.onmicrosoft.com; *.office.net; *.office.com; *.microsoft.com; *.microsoftonline.com; *.live.com; *.azure.net; *.gfx.ms; *.onestore.ms; *.msecnd.net; *.outlookgroups.ms; *.linkedin.com; *.msocdn.com; *.live.net; ihk-aka.de',
        [string]$ExcelListPath = '',
        [string]$CsvFallbackPath = '',
        [int]$MaxRows = 500,
        [switch]$Quiet,
        [switch]$RegistryOnly,
        [switch]$UseCom
    )

    try {
          Write-Info "[DEBUG] Start Hauptlauf (Version $script:AppVersion)"
          Write-SafeOutput "Dieses Skript richtet den Pruefungsrechner fuer die AP 1 ein." -ForegroundColor Yellow
        $desktopPath = Get-DesktopPath
          Write-Info "[DEBUG] DesktopPath: $desktopPath"
        # COM-Autodetektion
        $comOkWord = $false
        $comOkExcel = $false
        if ($UseCom -and -not $RegistryOnly) {
              Write-Info "[DEBUG] Pruefe COM-Verfuegbarkeit..."
            $comOkWord   = Test-ComAvailable -ProgId 'Word.Application'
            $comOkExcel  = Test-ComAvailable -ProgId 'Excel.Application'
              Write-Info "[DEBUG] COM Word: $comOkWord, COM Excel: $comOkExcel"
            if (-not ($comOkWord -and $comOkExcel)) {
                WriteWarn "COM-Start von Word/Excel fehlgeschlagen - wechsle in Registry-only-Modus."
                $RegistryOnly = $true
            } else {
                Write-Info "COM fuer Word/Excel verfuegbar - normaler Modus bleibt aktiv."
            }
        }
        # Nuera-Datei laden (immer)
          Write-Info "[DEBUG] Starte Nuera-Download..."
        Write-Info "Suche und lade neueste Nuera-Dateien..."
        $nueraDownloadPath = Join-Path $script:ScriptRoot '3. Nuera-Dateien'
        $nueraPath = $null
        $nueraZip = $null
        $fileInfos = @()
        # Erneut die Dateinamen wie im Modul
        $fileNames = @(
            "nuera2026_h.zip",
          "nuera2026_f.zip",
            "nuera2025_h.zip",
          "nuera2025_f.zip",
          "nuera2024_h.zip",
          "nuera2024_f.zip"
        )
        foreach ($fileName in $fileNames) {
            $testPath = Join-Path $nueraDownloadPath $fileName
            if (Test-Path $testPath) {
                $fileInfos += [pscustomobject]@{ FileName=$fileName; Path=$testPath; LastModified=(Get-Item $testPath).LastWriteTime }
            }
        }
        if ($fileInfos.Count -gt 0) {
            $latest = $fileInfos | Sort-Object LastModified -Descending | Select-Object -First 1
            $nueraZip = $latest.FileName
        }
        $nueraPath = Get-LatestNueraFile -DownloadPath $nueraDownloadPath
        if ($nueraPath) {
            Write-Info "[DEBUG] Kopiere Nuera-Ordner auf Desktop..."
            Write-Info "Kopiere Nuera-Ordner auf Desktop..."
          $desktopResult = Copy-NueraToDesktop -NueraSourcePath $nueraPath
            if ($desktopResult) {
                Write-Info "Nuera-Ordner erfolgreich auf Desktop kopiert: $desktopResult"
            } else {
                WriteWarn "Nuera-Ordner konnte nicht auf Desktop kopiert werden."
            }
        } else {
            WriteWarn "Nuera-Ordner konnte nicht bereitgestellt werden."
        }
        # Office vorbereiten: COM nur auf ausdrücklichen Wunsch starten.
          Write-Info "[DEBUG] Initialisiere Office..."
        if ($UseCom -and -not $RegistryOnly) {
          Write-Info "Initialisiere Office per COM..."
          Initialize-OfficeApps
        } else {
          $RegistryOnly = $true
          Write-Info "Registry-only-Modus: Office wird nicht per COM gestartet."
        }
          Write-Info "[DEBUG] Setze Office/Windows-Optionen..."
        Write-Info "Setze Office/Windows-Optionen..."
        Set-OfficeRegistrySettings
          Write-Info "[DEBUG] Setze Standard-Speicherpfade..."
        Write-Info "Setze Standard-Speicherpfad auf den Benutzerordner..."
          Write-Info "[DEBUG] Uebernehme Autokorrektur..."
        Write-Info "Uebernehme Autokorrektur-Einstellungen..."
        Set-WordAutoCorrectRegistry
        Set-ExcelAutoCorrectRegistry
          Write-Info "[DEBUG] Uebernehme Schnellzugriff..."
        Write-Info "Uebernehme Schnellzugriff-Symbolleisten..."
        Copy-QuickAccessToolbarFiles
          Write-Info "[DEBUG] Kopiere Vorlagen..."
        Write-Info "Kopiere Vorlagen (Normal.dotm und Mappe.xltx) mit Backup..."
        Copy-WordTemplate
        Copy-ExcelTemplate
        # Ordnererzeugung
        $rootPath = $null
        try {
            Write-Info "[DEBUG] Starte Ordnererzeugung..."
          if (-not $ExcelListPath) {
            $ExcelListPath = Join-Path $script:ScriptRoot '1. Anpassen\AP1-TN.xlsx'
            Write-Info "[DEBUG] ExcelListPath automatisch gesetzt: $ExcelListPath"
          }
          Write-Info "[DEBUG] Pruefe Existenz der Excel-Datei: $ExcelListPath"
          if (Test-Path $ExcelListPath) {
            Write-Info "Erzeuge Benutzerordner aus Excel (COM oder Python-Fallback)..."
            $rootPath = New-CandidateFoldersFromExcel -WorkbookPath $ExcelListPath -MaxRows $MaxRows -NoCom:$RegistryOnly
          } elseif ($CsvFallbackPath) {
            Write-Info "COM nicht verfuegbar - nutze CSV-Fallback..."
            $rootPath = New-CandidateFoldersFromCsv -CsvPath $CsvFallbackPath -MaxRows $MaxRows
          } else {
            throw "Excel-Datei nicht gefunden: $ExcelListPath"
          }
            Write-Info "[DEBUG] rootPath: $rootPath"
        } catch {
            Write-Info "[DEBUG] Fehler bei Ordnererzeugung: $($_.Exception.Message)"
          if ($CsvFallbackPath) {
            WriteWarn "Excel-Verarbeitung fehlgeschlagen. Versuche CSV-Fallback: $($_.Exception.Message)"
            $rootPath = New-CandidateFoldersFromCsv -CsvPath $CsvFallbackPath -MaxRows $MaxRows
          } else {
            throw
          }
        }
        # Ordner bereitstellen
        if ($rootPath) {
              Write-Info "[DEBUG] Kopiere Kandidaten-Ordner auf Desktop..."
            Write-Info "Lege Kandidaten-Ordner auf Desktop des aktuellen Nutzers..."
          $userFolderPath = Copy-CandidateFolderToDesktop -SourceRoot $rootPath
          if ($userFolderPath) {
            Write-Info "Setze Standard-Speicherpfad auf Benutzerordner: $userFolderPath"
            Set-DefaultSavePaths -Path $userFolderPath -UseCom:(!$RegistryOnly)
          } else {
            WriteWarn "Kein passender Benutzerordner gefunden; verwende den Desktop als Speicherort."
            Set-DefaultSavePaths -Path $desktopPath -UseCom:(!$RegistryOnly)
          }
            Write-Info "Raeume temporaeren Ordner auf..."
            try {
                Get-ChildItem $rootPath -ErrorAction SilentlyContinue |
                    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                if (-not (Get-ChildItem -Path $rootPath -Force -ErrorAction SilentlyContinue)) {
                    Remove-Item -Path $rootPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            } catch {}
        }
        # Taskbar und Proxy
          Write-Info "[DEBUG] Setze Taskbar/Proxy..."
        Write-Info "Taskbar-Einstellungen uebernehmen..."
        Set-TaskbarSettings -Alignment 'Left' -Search 'Icon'
        Write-Info "Proxy konfigurieren (falls angegeben)..."
        Set-Proxy -State $Proxy -Server $ProxyServer -BypassList $ProxyBypass
          Write-SafeOutput 'Fertig. Der Rechner ist fuer die AP 1 vorbereitet.' -ForegroundColor Green
    } catch {
          Write-SafeOutput "FEHLER beim Ausfuehren des Skripts: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.InvocationInfo.PositionMessage) {
              Write-SafeOutput "[DEBUG] Fehlerstelle: $($_.InvocationInfo.PositionMessage)" -ForegroundColor Red
        }
    } finally {
          Write-Info "[DEBUG] Stop-PrepTranscript aufgerufen"
        Stop-PrepTranscript
    }
}

function Main {
    # ScriptRoot initialisieren (für Logging und Pfade)
    if (-not $script:ScriptRoot) {
      $script:ScriptRoot = $script:DataRoot
    }
    $global:ScriptRoot = $script:ScriptRoot
    # Logging starten
    Start-PrepTranscript
    # ====== Funktions-Selbsttest: Prüft, ob alle Kernfunktionen geladen sind ======
    if ($MyInvocation.ScriptName) {
      # Selftest entfernt: Module und Funktionen werden direkt genutzt
    }
    # Hauptlauf
    $startParams = @{
      Proxy           = $Proxy
      ProxyServer     = $ProxyServer
      ProxyBypass     = $ProxyBypass
      ExcelListPath   = $ExcelListPath
      CsvFallbackPath = $CsvFallbackPath
      MaxRows         = $MaxRows
      Quiet           = $Quiet
      RegistryOnly    = $RegistryOnly
      UseCom          = $UseCom
    }
    Start-AP1Konfiguration @startParams
}

Main
