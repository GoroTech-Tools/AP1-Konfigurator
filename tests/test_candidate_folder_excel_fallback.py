from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _extract_python_fallback(script_path: Path) -> str:
    source = script_path.read_text(encoding='utf-8')
    start_marker = '$pythonScript = @"'
    start = source.find(start_marker)
    assert start >= 0, f'Python-Fallback-Start nicht gefunden in {script_path}'
    start += len(start_marker)
    if start < len(source) and source[start] == '\r':
        start += 1
    if start < len(source) and source[start] == '\n':
        start += 1

    end = source.find('\n"@', start)
    if end < 0:
        end = source.find('\r\n"@', start)
    assert end >= 0, f'Python-Fallback-Ende nicht gefunden in {script_path}'
    return source[start:end]


def _python_literal(path: Path) -> str:
    return str(path).replace("\\", "\\\\").replace("'", "\\'")


def _create_minimal_xlsx(path: Path, username: str, folder_name: str, *, include_cell_refs: bool = True) -> None:
    first_row = (
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
        if include_cell_refs
        else '<row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c></row>'
    )
    second_row = (
        '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row>'
        if include_cell_refs
        else '<row r="2"><c t="s"><v>2</v></c><c t="s"><v>3</v></c></row>'
    )
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr(
            '[Content_Types].xml',
            """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>""",
        )
        archive.writestr(
            'xl/workbook.xml',
            """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Tabelle1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )
        archive.writestr(
            'xl/_rels/workbook.xml.rels',
            """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        archive.writestr(
            'xl/sharedStrings.xml',
            f"""<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>anderer-user</t></si>
  <si><t>Ignorieren</t></si>
  <si><t>{username}</t></si>
  <si><t>{folder_name}</t></si>
</sst>""",
        )
        archive.writestr(
            'xl/worksheets/sheet1.xml',
            f"""<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    {first_row}
    {second_row}
  </sheetData>
</worksheet>""",
        )


def test_excel_python_fallback_works_without_openpyxl(tmp_path: Path) -> None:
    script = _extract_python_fallback(ROOT / 'src' / 'AP1-Konfigurator.ps1')
    workbook = tmp_path / 'AP1-TN.xlsx'
    root_path = tmp_path / 'Ordner'
    _create_minimal_xlsx(workbook, username='candidate.user', folder_name='Kandidat Eins')

    script = script.replace("$WorkbookPath", _python_literal(workbook))
    script = script.replace("$rootPath", _python_literal(root_path))
    script = script.replace("$MaxRows", "10")

    blocker = """
import builtins
_real_import = builtins.__import__
def _blocked_import(name, *args, **kwargs):
    if name == 'openpyxl':
        raise ImportError('blocked by test')
    return _real_import(name, *args, **kwargs)
builtins.__import__ = _blocked_import
"""

    env = os.environ.copy()
    env['USERNAME'] = 'candidate.user'
    result = subprocess.run(
        [sys.executable, '-c', blocker + script],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert (root_path / 'Kandidat Eins').is_dir()


def test_excel_python_fallback_works_when_openpyxl_read_fails(tmp_path: Path) -> None:
    script = _extract_python_fallback(ROOT / 'src' / 'AP1-Konfigurator.ps1')
    workbook = tmp_path / 'AP1-TN.xlsx'
    root_path = tmp_path / 'Ordner'
    _create_minimal_xlsx(workbook, username='candidate.user', folder_name='Kandidat Zwei')

    script = script.replace("$WorkbookPath", _python_literal(workbook))
    script = script.replace("$rootPath", _python_literal(root_path))
    script = script.replace("$MaxRows", "10")

    failing_openpyxl = """
import builtins
import types
import sys
_real_import = builtins.__import__
def _patched_import(name, *args, **kwargs):
    if name == 'openpyxl':
        module = types.ModuleType('openpyxl')
        def _load_workbook(*_args, **_kwargs):
            raise ValueError('simulated openpyxl failure')
        module.load_workbook = _load_workbook
        sys.modules[name] = module
        return module
    return _real_import(name, *args, **kwargs)
builtins.__import__ = _patched_import
"""

    env = os.environ.copy()
    env['USERNAME'] = 'candidate.user'
    result = subprocess.run(
        [sys.executable, '-c', failing_openpyxl + script],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert (root_path / 'Kandidat Zwei').is_dir()
    assert 'openpyxl-Lesezugriff fehlgeschlagen' in result.stderr


def test_excel_python_fallback_reads_cells_without_references(tmp_path: Path) -> None:
    script = _extract_python_fallback(ROOT / 'src' / 'AP1-Konfigurator.ps1')
    workbook = tmp_path / 'AP1-TN.xlsx'
    root_path = tmp_path / 'Ordner'
    _create_minimal_xlsx(
        workbook,
        username='candidate.user',
        folder_name='Kandidat Vier',
        include_cell_refs=False,
    )

    script = script.replace("$WorkbookPath", _python_literal(workbook))
    script = script.replace("$rootPath", _python_literal(root_path))
    script = script.replace("$MaxRows", "10")

    blocker = """
import builtins
_real_import = builtins.__import__
def _blocked_import(name, *args, **kwargs):
    if name == 'openpyxl':
        raise ImportError('blocked by test')
    return _real_import(name, *args, **kwargs)
builtins.__import__ = _blocked_import
"""

    env = os.environ.copy()
    env['USERNAME'] = 'candidate.user'
    result = subprocess.run(
        [sys.executable, '-c', blocker + script],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )

    assert result.returncode == 0, result.stderr
    assert (root_path / 'Kandidat Vier').is_dir()


def test_excel_python_fallback_does_not_hide_unexpected_openpyxl_errors(tmp_path: Path) -> None:
    script = _extract_python_fallback(ROOT / 'src' / 'AP1-Konfigurator.ps1')
    workbook = tmp_path / 'AP1-TN.xlsx'
    root_path = tmp_path / 'Ordner'
    _create_minimal_xlsx(workbook, username='candidate.user', folder_name='Kandidat Drei')

    script = script.replace("$WorkbookPath", _python_literal(workbook))
    script = script.replace("$rootPath", _python_literal(root_path))
    script = script.replace("$MaxRows", "10")

    broken_openpyxl = """
import builtins
import types
import sys
_real_import = builtins.__import__
def _patched_import(name, *args, **kwargs):
    if name == 'openpyxl':
        module = types.ModuleType('openpyxl')
        sys.modules[name] = module
        return module
    return _real_import(name, *args, **kwargs)
builtins.__import__ = _patched_import
"""

    env = os.environ.copy()
    env['USERNAME'] = 'candidate.user'
    result = subprocess.run(
        [sys.executable, '-c', broken_openpyxl + script],
        capture_output=True,
        text=True,
        env=env,
        check=False,
    )

    assert result.returncode != 0
    assert not (root_path / 'Kandidat Drei').exists()


def test_fallback_script_no_longer_installs_openpyxl_at_runtime() -> None:
    main_script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module_script = (ROOT / 'src' / 'Skript-Module' / 'AP1-Folders.psm1').read_text(encoding='utf-8')

    assert 'pip install openpyxl' not in main_script
    assert 'pip install openpyxl' not in module_script
    assert 'zipfile' in main_script
    assert 'zipfile' in module_script
