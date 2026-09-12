from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from build_info import BUILD_INFO

ROOT = Path(__file__).resolve().parents[1]
VERSION = f"v{BUILD_INFO['version']}"
ARTIFACT_NAME = BUILD_INFO['artifact_name']
LEGACY_ARTIFACT_NAME = 'AP1-Konfigurator'
DIST_ROOT = ROOT / 'dist'
RELEASE_ROOT = ROOT / 'release'
PACKAGE_DIR = DIST_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
RELEASE_DIR = RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
EXE_SOURCE = DIST_ROOT / f'{ARTIFACT_NAME}.exe'

COPY_MAP = {
    ROOT / 'README.md': 'README.md',
    ROOT / 'data': 'data',
    ROOT / 'docs': 'docs',
}

OBSOLETE_PACKAGE_PATHS = (
    Path('AP1-Konfigurator.ps1'),
    Path('AP1-Konfigurator.bat'),
    Path('Proxy-Deaktivieren.bat'),
    Path(f'RELEASE_NOTES_{VERSION}.md'),
    Path('Skript-Module'),
    Path('data/AP1-Konfigurator.bat'),
    Path('data/AP1-Konfigurator.ps1'),
    Path('data/Proxy-Deaktivieren.bat'),
    Path('data/CHANGELOG.md'),
    Path('data/DOCUMENTATION.md'),
    Path('data/Skript-Module'),
)


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def try_reset_dir(path: Path) -> bool:
    try:
        reset_dir(path)
        return True
    except PermissionError as exc:
        print(f'warnung: {path} konnte nicht bereinigt werden ({exc}). Spiegelordner wird übersprungen.')
        return False


def try_remove_path(path: Path) -> None:
    try:
        if path.is_dir():
            shutil.rmtree(path)
        elif path.exists():
            path.unlink()
    except PermissionError as exc:
        print(f'warnung: veraltetes Artefakt {path} konnte nicht entfernt werden ({exc}).')


def prune_obsolete_paths(package_root: Path) -> None:
    for relative_path in OBSOLETE_PACKAGE_PATHS:
        try_remove_path(package_root / relative_path)


def copy_item(source: Path, relative_target: str, package_dir: Path) -> None:
    target = package_dir / relative_target
    target.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir():
        shutil.copytree(source, target, dirs_exist_ok=True)
    else:
        shutil.copy2(source, target)


def ensure_data_mirror(source_data: Path, release_dir: Path) -> None:
    """Ensure the unpacked release mirror contains the complete data tree."""
    target_data = release_dir / 'data'
    target_data.mkdir(parents=True, exist_ok=True)
    for source in source_data.rglob('*'):
        relative_path = source.relative_to(source_data)
        target = target_data / relative_path
        if source.is_dir():
            target.mkdir(parents=True, exist_ok=True)
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy2(source, target)
        except PermissionError as exc:
            print(f'warnung: data-Datei konnte nicht gespiegelt werden {target} ({exc}).')

    expected = {
        path.relative_to(source_data)
        for path in source_data.rglob('*')
        if path.is_file()
    }
    actual = {
        path.relative_to(target_data)
        for path in target_data.rglob('*')
        if path.is_file()
    }
    missing = sorted(expected - actual)
    if missing:
        raise RuntimeError(
            'Release-Spiegelordner unvollständig; fehlende data-Dateien: '
            + ', '.join(str(path) for path in missing)
        )


def create_zip_archive(source_dir: Path, zip_path: Path) -> Path:
    if zip_path.exists():
        zip_path.unlink()

    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(source_dir.rglob('*')):
            relative_path = path.relative_to(source_dir)
            archive_name = relative_path.as_posix()
            if path.is_dir():
                archive.writestr(f'{archive_name}/', '')
            else:
                archive.write(path, archive_name)

    return zip_path


def main() -> int:
    if not EXE_SOURCE.exists():
        raise FileNotFoundError(f'PyInstaller-Ausgabe fehlt: {EXE_SOURCE}')

    legacy_release_dir = RELEASE_ROOT / f'{LEGACY_ARTIFACT_NAME}-{VERSION}'
    legacy_release_zip = RELEASE_ROOT / f'{LEGACY_ARTIFACT_NAME}-{VERSION}.zip'
    try_remove_path(legacy_release_dir)
    try_remove_path(legacy_release_zip)

    with tempfile.TemporaryDirectory(prefix='ap1-konfigurator-') as temp_dir:
        staging_root = Path(temp_dir)
        staging_package_dir = staging_root / f'{ARTIFACT_NAME}-{VERSION}'
        staging_package_dir.mkdir(parents=True, exist_ok=True)

        dist_dir_ready = try_reset_dir(PACKAGE_DIR)
        release_dir_ready = try_reset_dir(RELEASE_DIR)

        shutil.copy2(EXE_SOURCE, staging_package_dir / f'{ARTIFACT_NAME}.exe')
        for source, target in COPY_MAP.items():
            if source.exists():
                copy_item(source, target, staging_package_dir)

        (staging_package_dir / 'data' / '4. Logs').mkdir(parents=True, exist_ok=True)

        if dist_dir_ready:
            shutil.copytree(staging_package_dir, PACKAGE_DIR, dirs_exist_ok=True)
        else:
            print(f'warnung: {PACKAGE_DIR} konnte nicht aktualisiert werden. Dist-Spiegelordner wird übersprungen.')
        prune_obsolete_paths(PACKAGE_DIR)

        if release_dir_ready:
            shutil.copytree(staging_package_dir, RELEASE_DIR, dirs_exist_ok=True)
        ensure_data_mirror(ROOT / 'data', RELEASE_DIR)
        prune_obsolete_paths(RELEASE_DIR)

        zip_path = create_zip_archive(staging_package_dir, RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}.zip')

    print(f'paketiert: {zip_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
