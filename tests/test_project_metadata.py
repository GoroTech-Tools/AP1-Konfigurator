from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _read_version_from_build_info() -> str:
    source = (ROOT / 'src' / 'build_info.py').read_text(encoding='utf-8')
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == 'BUILD_INFO':
                    if isinstance(node.value, ast.Dict):
                        for key_node, value_node in zip(node.value.keys, node.value.values):
                            if isinstance(key_node, ast.Constant) and key_node.value == 'version':
                                if isinstance(value_node, ast.Constant) and isinstance(value_node.value, str):
                                    return value_node.value
    raise AssertionError('Version in src/build_info.py konnte nicht gelesen werden.')


def test_readme_mentions_current_release_version() -> None:
    version = _read_version_from_build_info()
    readme = (ROOT / 'README.md').read_text(encoding='utf-8')
    assert f'v{version}' in readme
    assert 'Aktueller Stand: **v' in readme
    assert 'Der aktuelle Release-Stand ist `v' not in readme


def test_release_notes_exist_for_current_version() -> None:
    version = _read_version_from_build_info()
    release_note = ROOT / 'release' / f'RELEASE_NOTES_v{version}.md'
    assert release_note.exists(), f'Release Notes fehlen für v{version}'
