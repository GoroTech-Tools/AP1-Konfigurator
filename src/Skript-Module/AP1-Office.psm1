# AP1-Office.psm1
# Modul fuer Office-/Word-/Excel-Logik

function Test-OfficeFirstRun {
    # Prüft, ob Office FirstRun noch aussteht (Word/Excel)
    $wordRegPath = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Word\Options"
    $excelRegPath = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Excel\Options"
    $officeProfilePath = "HKCU:\Software\Microsoft\Office\$script:OfficeVersion\Common\General"

    $wordFirstRun = $false
    $excelFirstRun = $false
    $profileExists = Test-Path $officeProfilePath

    # Word prüfen
    if (Test-Path $wordRegPath) {
        $wordProps = Get-ItemProperty -Path $wordRegPath -Name "FirstRun" -ErrorAction SilentlyContinue
        if ($null -ne $wordProps -and $wordProps.PSObject.Properties["FirstRun"]) {
            if ($wordProps.FirstRun -eq 1) {
                Write-Info "[FirstRun] Word: Wert ist 1 (FirstRun steht aus)"
                $wordFirstRun = $true
            } else {
                Write-Info "[FirstRun] Word: Wert ist $($wordProps.FirstRun) (kein FirstRun)"
            }
        } else {
            Write-Info "[FirstRun] Word: Wert fehlt, FirstRun steht aus."
            $wordFirstRun = $true
        }
    } else {
        Write-Info "[FirstRun] Word: Registry-Key fehlt, FirstRun steht aus."
        $wordFirstRun = $true
    }

    # Excel prüfen
    if (Test-Path $excelRegPath) {
        $excelProps = Get-ItemProperty -Path $excelRegPath -Name "FirstRun" -ErrorAction SilentlyContinue
        if ($null -ne $excelProps -and $excelProps.PSObject.Properties["FirstRun"]) {
            if ($excelProps.FirstRun -eq 1) {
                Write-Info "[FirstRun] Excel: Wert ist 1 (FirstRun steht aus)"
                $excelFirstRun = $true
            } else {
                Write-Info "[FirstRun] Excel: Wert ist $($excelProps.FirstRun) (kein FirstRun)"
            }
        } else {
            Write-Info "[FirstRun] Excel: Wert fehlt, FirstRun steht aus."
            $excelFirstRun = $true
        }
    } else {
        Write-Info "[FirstRun] Excel: Registry-Key fehlt, FirstRun steht aus."
        $excelFirstRun = $true
    }

    if (-not $profileExists) {
        Write-Info "[FirstRun] Office-Profil fehlt (wird aber ignoriert, wenn Word/Excel initialisiert sind)."
    }

    return ($wordFirstRun -or $excelFirstRun)
}

function Confirm-OfficeWelcomeDialogs {
    param(
        [string[]]$ProcessNames = @('WINWORD', 'EXCEL'),
        [int]$Attempts = 5,
        [int]$DelaySeconds = 2
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        WriteWarn "System.Windows.Forms konnte nicht geladen werden; Tastendruck zum Bestätigen wurde übersprungen."
        return
    }

    $shell = New-Object -ComObject WScript.Shell -ErrorAction SilentlyContinue
    if (-not $shell) {
        WriteWarn "WScript.Shell nicht verfügbar; Office-Willkommensdialoge konnten nicht automatisch bestätigt werden."
        return
    }

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        foreach ($processName in $ProcessNames) {
            $proc = Get-Process $processName -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
            if (-not $proc) { continue }

            try {
                $null = $shell.AppActivate($proc.Id)
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
                Start-Sleep -Milliseconds 250
                [System.Windows.Forms.SendKeys]::SendWait(' ')
                Write-Info "Office-Willkommensdialog für $processName automatisch bestätigt (Versuch $attempt/$Attempts)."
            } catch {
                WriteWarn "Bestätigung für $processName fehlgeschlagen: $($_.Exception.Message)"
            }
        }
        if ($attempt -lt $Attempts) { Start-Sleep -Seconds $DelaySeconds }
    }
}

function Initialize-OfficeApps {
    Write-Info "Initialisiere Word/Excel (robust)..."
    $maxTries = 3
    $try = 0
    $success = $false
    $isActualFirstRun = Test-OfficeFirstRun
    Stop-NamedProcess WINWORD
    Stop-NamedProcess EXCEL
    Get-Process OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    while (-not $success -and $try -lt $maxTries) {
        $try++
        $processStarted = $false
        if (-not (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
            Start-Process WINWORD -WindowStyle Minimized
            $processStarted = $true
        }
        if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) {
            Start-Process EXCEL -WindowStyle Minimized
            $processStarted = $true
        }
        if ($isActualFirstRun -and $processStarted) {
            Write-Host -ForegroundColor Yellow "`nErster Office-Start erkannt!"
            Write-Host -ForegroundColor Yellow "Versuche, Office-Willkommensfenster automatisch zu bestaetigen..."
            Confirm-OfficeWelcomeDialogs -ProcessNames @('WINWORD', 'EXCEL') -Attempts 6 -DelaySeconds 2
            Write-Host -ForegroundColor Yellow "Wenn ein Hinweisfenster offen bleibt, kann das Skript nach kurzer Zeit automatisch weiterlaufen."
        } elseif ($processStarted) {
            Write-Info "Office-Prozesse gestartet (kein First-Run erkannt)"
            Start-Sleep -Seconds 2
        } else {
            Write-Info "Office-Prozesse liefen bereits"
        }
        Stop-NamedProcess WINWORD
        Stop-NamedProcess EXCEL
        Get-Process OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        $comOkWord = $false
        $comOkExcel = $false
        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
            $comOkWord = $true
        } catch {
            WriteWarn "Word COM konnte nicht gestartet werden (Versuch $try/$maxTries): $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            $comOkExcel = $true
        } catch {
            WriteWarn "Excel COM konnte nicht gestartet werden (Versuch $try/$maxTries): $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
        if ($comOkWord -and $comOkExcel) {
            $success = $true
        } else {
            Write-Info "Warte 3 Sekunden und versuche erneut..."
            Start-Sleep -Seconds 3
        }
    }
    if (-not $success) {
        WriteWarn "Word/Excel COM konnte nach $maxTries Versuchen nicht initialisiert werden. Wechsle in Registry-only-Modus."
        $script:RegistryOnly = $true
    } else {
        Write-Info "Word/Excel COM erfolgreich initialisiert."
        # Setze First-Run-Marker für zukünftige Laeufe
        Set-OfficeFirstRunMarkers
    }
}

Export-ModuleMember -Function Test-OfficeFirstRun,Confirm-OfficeWelcomeDialogs,Initialize-OfficeApps
