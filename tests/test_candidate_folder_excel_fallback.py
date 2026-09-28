from __future__ import annotations

import getpass
import subprocess
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

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


def _create_minimal_inline_xlsx(path: Path, first_cell_ref: str, second_cell_ref: str, username: str, candidate: str) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Tabelle1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""
    sheet = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="{first_cell_ref}" t="inlineStr"><is><t>{username}</t></is></c>
      <c r="{second_cell_ref}" t="inlineStr"><is><t>{candidate}</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', content_types)
        archive.writestr('xl/workbook.xml', workbook)
        archive.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        archive.writestr('xl/worksheets/sheet1.xml', sheet)


def _create_minimal_shared_string_xlsx(path: Path, username: str, candidate: str) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Tabelle1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""
    shared_strings = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>{username}</t></si>
  <si><t>{candidate}</t></si>
</sst>
"""
    sheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>
"""
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', content_types)
        archive.writestr('xl/workbook.xml', workbook)
        archive.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        archive.writestr('xl/worksheets/sheet1.xml', sheet)
        archive.writestr('xl/sharedStrings.xml', shared_strings)


def _run_fallback_function_in_pwsh(xlsx_path: Path, script_root: Path) -> list[str]:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    signature = 'function New-CandidateFoldersFromExcel {'
    start = script.find(signature)
    assert start != -1
    body = _extract_brace_block(script, start)
    function_text = f'{signature}{body}}}'
    script_root.mkdir(parents=True, exist_ok=True)
    escaped_script_root = str(script_root).replace("'", "''")
    escaped_xlsx_path = str(xlsx_path).replace("'", "''")
    harness = f"""
$ErrorActionPreference = 'Stop'
$script:ScriptRoot = '{escaped_script_root}'
function Write-Info {{ param([string]$Message) }}
function Clear-ComObject {{ param($Object) }}
function Stop-NamedProcess {{ param([string]$Name) }}
function Remove-EmptyCandidateRoot {{
  param([string]$RootPath)
  if (Test-Path $RootPath) {{
    $childItems = @(Get-ChildItem -Path $RootPath -Force -ErrorAction SilentlyContinue)
    if ($childItems.Count -eq 0) {{
      Remove-Item -Path $RootPath -Recurse -Force -ErrorAction SilentlyContinue
    }}
  }}
}}
function New-EnsuredPath {{
  param([string]$Path)
  if (-not (Test-Path $Path)) {{ New-Item -Path $Path -ItemType Directory -Force | Out-Null }}
}}
{function_text}
New-CandidateFoldersFromExcel -WorkbookPath '{escaped_xlsx_path}' -NoCom -MaxRows 5 | Out-Null
"""
    subprocess.run(
        ['pwsh', '-NoProfile', '-NonInteractive', '-Command', harness],
        check=True,
        capture_output=True,
        text=True,
    )
    candidate_root = script_root / '2. Bei Bedarf anpassen' / 'Ordner'
    return [item.name for item in candidate_root.iterdir() if item.is_dir()]


def test_excel_fallback_reads_only_exact_a_b_columns() -> None:
    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        xlsx_path = temp_path / 'AP1-TN.xlsx'
        username = getpass.getuser()
        candidate = 'MUSS_IGNORIERT_WERDEN'
        _create_minimal_inline_xlsx(
            xlsx_path,
            first_cell_ref='AA1',
            second_cell_ref='BA1',
            username=username,
            candidate=candidate,
        )

        created_folders = _run_fallback_function_in_pwsh(xlsx_path, temp_path / 'runtime')
        assert candidate not in created_folders


def test_excel_fallback_reads_shared_strings_from_relationship_sheet() -> None:
    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        xlsx_path = temp_path / 'AP1-TN.xlsx'
        username = getpass.getuser()
        candidate = 'Max Mustermann'
        _create_minimal_shared_string_xlsx(
            xlsx_path,
            username=username,
            candidate=candidate,
        )

        created_folders = _run_fallback_function_in_pwsh(xlsx_path, temp_path / 'runtime')
        assert 'Max Mustermann' in created_folders


def test_excel_fallback_uses_embedded_xlsx_xml_reader() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')

    assert 'System.IO.Compression.ZipFile' in script
    assert 'sharedStrings.xml' in script
    assert 'XLSX-XML-Fallback' in script


def test_excel_fallback_no_longer_installs_openpyxl() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')

    assert '-m pip install openpyxl' not in script
    assert 'import openpyxl' not in script
