# AP1-Logging.psm1
function Write-Info {
	param([string]$Message)
	# Log-Ausgaben müssen auch im interaktiven/PowerShell-Lauf sichtbar sein.
	# Das Start-Transcript erfasst dieselben Meldungen zusätzlich auf Datei.
	WriteHostSafe $Message
}
function Write-SafeOutput {
	param(
		[Parameter(Mandatory)][string]$Message,
		[string]$ForegroundColor
	)
	WriteHostSafe -Message $Message -ForegroundColor $ForegroundColor
	Write-Output $Message
}

function WriteHostSafe {
	param(
		[Parameter(Mandatory)][string]$Message,
		[string]$ForegroundColor
	)
	$msg = $Message
	if ($PSBoundParameters.ContainsKey('ForegroundColor') -and $ForegroundColor) {
		Write-Host $msg -ForegroundColor $ForegroundColor
	} else {
		Write-Host $msg
	}
}
function WriteWarn {
	param([string]$Message)
	WriteHostSafe "(WARN) $Message" -ForegroundColor Yellow
	Write-Output "[WARN] $Message"
}
function WriteErr {
	param([string]$Message)
	WriteHostSafe "(ERROR) $Message" -ForegroundColor Red
	Write-Output "[ERROR] $Message"
}

function Stop-PrepTranscript { try { Stop-Transcript | Out-Null } catch {} }
function Start-PrepTranscript {
	try {
		# Hauptskript-Verzeichnis bestimmen
		$mainScriptRoot = $null
		if ($global:ScriptRoot) { $mainScriptRoot = $global:ScriptRoot }
		elseif ($PSScriptRoot -and $PSScriptRoot -notlike '*Skript-Module*') { $mainScriptRoot = $PSScriptRoot }
		elseif ($MyInvocation -and $MyInvocation.ScriptName) { $mainScriptRoot = Split-Path -Path $MyInvocation.ScriptName -Parent }
		else { $mainScriptRoot = (Get-Location).Path }
		$logDir = Join-Path $mainScriptRoot '4. Logs'
		if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

		# Log-Rotation: Nur die letzten 5 AP1_Prep_*.log behalten
		$logFiles = Get-ChildItem -Path $logDir -Filter 'AP1_Prep_*.log' | Sort-Object LastWriteTime -Descending
		if ($logFiles.Count -gt 5) {
			$toDelete = $logFiles | Select-Object -Skip 5
			foreach ($oldLog in $toDelete) {
				try { Remove-Item $oldLog.FullName -Force } catch {}
			}
		}

		$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
		$logPath = Join-Path $logDir "AP1_Prep_$ts.log"
		if ($logPath) {
			Start-Transcript -Path $logPath -ErrorAction Stop | Out-Null
			return $logPath
		} else {
			WriteWarn "Log-Pfad konnte nicht bestimmt werden."
			return $null
		}
	} catch {
		WriteWarn "Fehler beim Starten des Transcripts: $($_.Exception.Message)"
		return $null
	}
}

# Alle Funktionsdefinitionen bleiben unveraendert
Export-ModuleMember -Function Write-Info,Write-SafeOutput,WriteHostSafe,WriteWarn,WriteErr,Stop-PrepTranscript,Start-PrepTranscript,Remove-OldLogFiles
