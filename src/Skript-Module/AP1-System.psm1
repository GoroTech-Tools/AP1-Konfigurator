
function Set-Proxy {
	param(
		[ValidateSet('On','Off','Skip')] [string]$State,
		[string]$Server, # Wird über die aufrufende Batch-Datei übergeben, z. B. 192.168.0.1:8080
		[string]$BypassList # wird über die aufrufende Batch-Datei übergeben, z. B. "*.office365.com; *.cloudappsecurity.com; *.onmicrosoft.com; *.office.net; *.office.com; *.microsoft.com; *.microsoftonline.com; *.live.com; *.azure.net; *.gfx.ms; *.onestore.ms; *.msecnd.net; *.outlookgroups.ms; *.linkedin.com; *.msocdn.com; *.live.net; ihk-aka.de"
	)
	if ($State -eq 'Skip') { Write-Info "Proxy-Konfiguration übersprungen."; return }
	$reg = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
	$enable = if ($State -eq 'On') { 1 } else { 0 }
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
}
function Stop-NamedProcess {
	param([string]$Name)
	try { Get-Process -Name $Name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch {}
}
function Expand-ZipWithShell {
	param(
		[Parameter(Mandatory)] [string]$ZipPath,
		[Parameter(Mandatory)] [string]$Destination
	)
	New-EnsuredPath $Destination
	try {
		$shell = New-Object -ComObject Shell.Application
		$zip = $shell.NameSpace($ZipPath)
		$dest = $shell.NameSpace($Destination)
		if (-not $zip -or -not $dest) { throw "Shell.Application konnte Pfade nicht oeffnen." }
		$dest.CopyHere($zip.Items(), 16)
		Start-Sleep -Seconds 2
		Write-Info "ZIP mit Explorer extrahiert: $ZipPath"
		return $true
	} catch {
		Write-Warning "Explorer-Extraktion fehlgeschlagen: $($_.Exception.Message)"
		return $false
	}
}
function Test-ExpandArchive {
	param([Parameter(Mandatory)] [string]$ZipPath, [Parameter(Mandatory)] [string]$Destination)
	New-EnsuredPath $Destination
	if (Expand-ZipWithShell -ZipPath $ZipPath -Destination $Destination) { return $true }
	$sevenZip = $null
	foreach ($path in $env:Path.Split(';')) {
		$candidate = Join-Path $path '7z.exe'
		if (Test-Path $candidate) { $sevenZip = $candidate; break }
	}
	if ($sevenZip) {
		try {
			Write-Info "Entpacke mit 7-Zip (UTF-8)..."
			& $sevenZip x -y -mcp=UTF-8 -o"$Destination" "$ZipPath" | Out-Null
			return $true
		} catch {
			Write-Warning "7-Zip-Entpacken fehlgeschlagen: $($_.Exception.Message)"
		}
	}
	try {
		Write-Info "Entpacke mit Expand-Archive..."
		Expand-Archive -LiteralPath $ZipPath -DestinationPath $Destination -Force
		return $true
	} catch {
		try {
			Write-Info "Entpacke mit tar..."
			& tar -xf $ZipPath -C $Destination
			return $true
		} catch {
			Write-Warning "Archiv konnte nicht extrahiert werden: $($_.Exception.Message)"
			Write-Warning "Bitte pruefen Sie das entpackte Nuera-Verzeichnis manuell auf korrekte Umlaute und Vollstaendigkeit!"
			return $false
		}
	}
}

Export-ModuleMember -Function Stop-NamedProcess,Expand-ZipWithShell,Test-ExpandArchive,Set-Proxy,Set-TaskbarSettings