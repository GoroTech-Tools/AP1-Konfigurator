from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_desktop_quick_access_function_is_integrated() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')
    normalized_script = re.sub(r'\s+', '', script).lower()
    normalized_module = re.sub(r'\s+', '', module).lower()

    assert 'functionadd-desktoptoquickaccess{' in normalized_module
    assert re.search(r"getfolderpath\((\"|')desktop(\"|')\)", normalized_module)
    assert 'export-modulemember-function' in normalized_module
    assert 'add-desktoptoquickaccess' in normalized_module
    assert 'if($usecom){try{add-desktoptoquickaccess' in normalized_script
