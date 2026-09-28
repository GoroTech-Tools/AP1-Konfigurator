from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def test_candidate_folder_copy_ensures_desktop_path_exists() -> None:
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-Folders.psm1').read_text(encoding='utf-8')

    assert '$desktop = Get-DesktopPath' in module
    assert 'New-EnsuredPath $desktop' in module


def test_get_desktop_path_has_userprofile_fallback() -> None:
    module = (ROOT / 'src' / 'Skript-Module' / 'AP1-Folders.psm1').read_text(encoding='utf-8')

    assert "[Environment]::GetFolderPath('Desktop')" in module
    assert "[Environment]::GetFolderPath('UserProfile')" in module
    assert "$userProfile = $env:USERPROFILE" in module
    assert "Join-Path $userProfile 'Desktop'" in module
