from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _extract_brace_block(source: str, block_start: int) -> str:
    depth = 0
    body_start = None
    for index in range(block_start, len(source)):
        char = source[index]
        if char == '{':
            depth += 1
            if depth == 1:
                body_start = index + 1
        elif char == '}':
            depth -= 1
            if depth == 0 and body_start is not None:
                return source[body_start:index]
    raise AssertionError('Klammerblock konnte nicht vollständig gelesen werden.')


def test_desktop_quick_access_function_is_integrated() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')

    function_match = re.search(r'function\s+Add-DesktopToQuickAccess\s*\{', module, re.IGNORECASE)
    assert function_match is not None
    function_body = _extract_brace_block(module, function_match.start())
    assert re.search(r"GetFolderPath\((\"|')Desktop(\"|')\)", function_body, re.IGNORECASE)

    assert re.search(
        r'^\s*Export-ModuleMember\s+-Function\s+[^\r\n#]*\bAdd-DesktopToQuickAccess\b',
        module,
        re.IGNORECASE | re.MULTILINE,
    )

    use_com_match = re.search(r'if\s*\(\$UseCom\)\s*\{', script, re.IGNORECASE)
    assert use_com_match is not None
    use_com_body = _extract_brace_block(script, use_com_match.start())
    assert re.search(r'^\s*(?:try\s*\{\s*)?Add-DesktopToQuickAccess\b', use_com_body, re.IGNORECASE | re.MULTILINE)
