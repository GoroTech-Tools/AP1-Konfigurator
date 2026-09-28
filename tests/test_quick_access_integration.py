from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_desktop_quick_access_function_is_integrated() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')

    assert re.search(
        r"function Add-DesktopToQuickAccess\s*\{[\s\S]*?\[Environment\]::GetFolderPath\('Desktop'\)",
        module,
    )
    assert re.search(r"Export-ModuleMember[^\n]*Add-DesktopToQuickAccess", module)
    assert re.search(r"if \(\$UseCom\)\s*\{\s*try \{\s*Add-DesktopToQuickAccess", script)
