from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_taskbar_glom_level_is_configured() -> None:
    ps1 = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-System.psm1').read_text(encoding='utf-8')

    assert "'TaskbarAI' = 0" in ps1
    assert "'TaskbarAI' = 0" in module
    assert "TaskbarGlomLevel" in ps1
    assert "TaskbarGlomLevel" in module
    assert "= 2" in ps1
    assert "= 2" in module
    assert "SendMessageTimeout" in ps1
    assert "SendMessageTimeout" in module


def test_explorer_is_restarted_after_configuration() -> None:
    ps1 = (ROOT / 'src' / 'AP1-Konfigurator.ps1').read_text(encoding='utf-8')
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-System.psm1').read_text(encoding='utf-8')

    for script in (ps1, module):
        assert 'function Restart-Explorer {' in script
        assert "Get-Process -Name explorer" in script
        assert "Start-Process -FilePath 'explorer.exe'" in script
    assert 'Export-ModuleMember' in module
    assert 'Restart-Explorer' in module[module.index('Export-ModuleMember'):]

    finally_block = ps1[ps1.index('} finally {'):]
    assert finally_block.index('Restart-Explorer') < finally_block.index('Stop-PrepTranscript')
