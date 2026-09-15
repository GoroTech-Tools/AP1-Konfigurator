from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

from build_info import BUILD_INFO

ROOT = Path(__file__).resolve().parents[1]
VERSION = f"v{BUILD_INFO['version']}"
GERMAN_MONTH_NAMES = (
    'Januar',
    'Februar',
    'März',
    'April',
    'Mai',
    'Juni',
    'Juli',
    'August',
    'September',
    'Oktober',
    'November',
    'Dezember',
)
build_date_value = str(BUILD_INFO['build_date']).replace('Z', '+00:00')
build_datetime = datetime.fromisoformat(build_date_value)
DATE_TEXT = f'{build_datetime.day}. {GERMAN_MONTH_NAMES[build_datetime.month - 1]} {build_datetime.year}'

REPLACEMENTS: dict[Path, list[tuple[str, str]]] = {
    ROOT / 'README.md': [
        (r"Aktueller Stand: \*\*v\d+\.\d+\.\d+\*\* · Letzte Aktualisierung: \*\*[^*]+\*\*", f"Aktueller Stand: **{VERSION}** · Letzte Aktualisierung: **{DATE_TEXT}**"),
        (r"release/RELEASE_NOTES_v\d+\.\d+\.\d+\.md", f"release/RELEASE_NOTES_{VERSION}.md"),
    ],
    ROOT / 'docs' / 'DOKUMENTATION_TECHNIK.md': [
        (r"Aktueller Stand: \*\*v\d+\.\d+\.\d+\*\* · Letzte Aktualisierung: \*\*[^*]+\*\*", f"Aktueller Stand: **{VERSION}** · Letzte Aktualisierung: **{DATE_TEXT}**"),
        (r"release/RELEASE_NOTES_v\d+\.\d+\.\d+\.md", f"release/RELEASE_NOTES_{VERSION}.md"),
    ],
    ROOT / 'docs' / 'DOKUMENTATION_ANWENDER.md': [
        (r"Aktueller Stand: \*\*v\d+\.\d+\.\d+\*\* · Letzte Aktualisierung: \*\*[^*]+\*\*", f"Aktueller Stand: **{VERSION}** · Letzte Aktualisierung: **{DATE_TEXT}**"),
    ],
    ROOT / 'docs' / 'DOKUMENTATION_CHECKLISTE.md': [
        (r"Aktueller Stand: \*\*v\d+\.\d+\.\d+\*\* · Letzte Aktualisierung: \*\*[^*]+\*\*", f"Aktueller Stand: **{VERSION}** · Letzte Aktualisierung: **{DATE_TEXT}**"),
    ],
}

for file_path, replacements in REPLACEMENTS.items():
    if not file_path.exists():
        continue
    content = file_path.read_text(encoding='utf-8')
    updated = content
    for pattern, replacement in replacements:
        updated = re.sub(pattern, replacement, updated)
    if updated != content:
        file_path.write_text(updated, encoding='utf-8')
        print(f'aktualisiert: {file_path.relative_to(ROOT)}')
