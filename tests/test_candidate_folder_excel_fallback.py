from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_excel_fallback_uses_embedded_xlsx_xml_reader() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')

    assert 'System.IO.Compression.ZipFile' in script
    assert 'sharedStrings.xml' in script
    assert 'XLSX-XML-Fallback' in script


def test_excel_fallback_no_longer_installs_openpyxl() -> None:
    script = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')

    assert '-m pip install openpyxl' not in script
    assert 'import openpyxl' not in script
