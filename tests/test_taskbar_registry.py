from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_taskbar_glom_level_is_configured() -> None:
    ps1 = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-System.psm1').read_text(encoding='utf-8')

    assert "TaskbarGlomLevel" in ps1
    assert "TaskbarGlomLevel" in module
    assert "= 2" in ps1
    assert "= 2" in module
    assert "SendMessageTimeout" in ps1
    assert "SendMessageTimeout" in module
    assert "Stop-Process -Name explorer" not in ps1
    assert "Stop-Process -Name explorer" not in module
