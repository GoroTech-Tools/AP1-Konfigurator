from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_desktop_quick_access_function_is_integrated() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')

    function_match = re.search(r"function Add-DesktopToQuickAccess\s*\{(?P<body>.*?)\n\}", module, re.DOTALL)
    assert function_match is not None
    assert "[Environment]::GetFolderPath('Desktop')" in function_match.group('body')

    export_line = next((line for line in module.splitlines() if 'Export-ModuleMember -Function' in line), '')
    assert 'Add-DesktopToQuickAccess' in export_line

    use_com_block = re.search(r"if \(\$UseCom\)\s*\{(?P<body>.*?)\n\}", script, re.DOTALL)
    assert use_com_block is not None
    assert 'Add-DesktopToQuickAccess' in use_com_block.group('body')
