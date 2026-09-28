from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _extract_brace_block(source: str, block_start: int) -> str:
    depth = 0
    body_start = None
    in_single = False
    in_double = False
    in_comment = False
    for index in range(block_start, len(source)):
        char = source[index]
        if in_comment:
            if char == '\n':
                in_comment = False
            continue
        if in_single:
            if char == "'":
                in_single = False
            continue
        if in_double:
            if char == '"' and (index == 0 or source[index - 1] != '`'):
                in_double = False
            continue
        if char == '#':
            in_comment = True
            continue
        if char == "'":
            in_single = True
            continue
        if char == '"':
            in_double = True
            continue
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
    assert re.search(r'\[switch\]\$SortAlphabetically', function_body, re.IGNORECASE)
    assert re.search(r'\[switch\]\$SortDryRun', function_body, re.IGNORECASE)
    assert re.search(r'GetFolderPath\(\[Environment\+SpecialFolder\]::DesktopDirectory\)', function_body, re.IGNORECASE)
    assert 'Remove-StaleDesktopQuickAccessEntries' in function_body
    assert 'Invoke-PinToQuickAccessByPath' in function_body
    assert 'Reorder-QuickAccessAlphabetically' in function_body

    assert re.search(
        r'^\s*Export-ModuleMember\s+-Function\s+[^\r\n#]*\bAdd-DesktopToQuickAccess\b',
        module,
        re.IGNORECASE | re.MULTILINE,
    )

    use_com_matches = re.finditer(r'if\s*\(\$UseCom\)\s*\{', script, re.IGNORECASE)
    assert any(
        re.search(
            r'^\s*(?:try\s*\{\s*)?Add-DesktopToQuickAccess\b[^\r\n#]*-SortAlphabetically:\$SortAlphabetically\b[^\r\n#]*-SortDryRun:\$SortDryRun\b',
            _extract_brace_block(script, match.start()),
            re.IGNORECASE | re.MULTILINE,
        )
        for match in use_com_matches
    )


def test_quick_access_module_contains_sorting_and_stale_desktop_helpers() -> None:
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-QuickAccess.psm1').read_text(encoding='utf-8')

    for function_name in (
        'Get-NormalizedPath',
        'Get-ShellVerb',
        'Invoke-PinToQuickAccessByPath',
        'Get-QuickAccessFilesystemEntries',
        'Test-StringSequenceEqual',
        'Reorder-QuickAccessAlphabetically',
        'Remove-StaleDesktopQuickAccessEntries',
    ):
        assert re.search(rf'function\s+{re.escape(function_name)}\s*\{{', module, re.IGNORECASE)

    assert 'pintohome' in module
    assert 'unpinfromhome' in module
    assert 'shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}' in module
