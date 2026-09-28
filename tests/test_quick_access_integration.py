from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_desktop_quick_access_function_is_integrated() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')

    assert "function Add-DesktopToQuickAccess" in module
    assert "[Environment]::GetFolderPath('Desktop')" in module
    assert "Export-ModuleMember -Function Test-Shortcut,New-Shortcut,Add-DesktopToQuickAccess" in module
    assert "if ($UseCom)" in script
    assert "Add-DesktopToQuickAccess" in script
