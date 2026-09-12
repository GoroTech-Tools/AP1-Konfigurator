from __future__ import annotations

import ctypes
from datetime import datetime
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.request
import zipfile
from tkinter import messagebox, ttk
from pathlib import Path

CREATE_NO_WINDOW = 0x08000000
MB_ICONERROR = 0x00000010
MB_OK = 0x00000000
EMBEDDED_FILES = (
    'AP1-Konfigurator.ps1',
    'AP1-Konfigurator.bat',
    'Proxy-Deaktivieren.bat',
    'app_icon.ico',
)
EMBEDDED_DIRS = (
    'Skript-Module',
    'data',
    'docs',
)
PACKAGE_CONTENT_DIRS = (
    'data',
    'docs',
)
DEFAULT_PROXY_SERVER = '192.168.0.1:8080'
DEFAULT_PROXY_BYPASS = '*.office365.com; *.cloudappsecurity.com; *.onmicrosoft.com; *.office.net; *.office.com; *.microsoft.com; *.microsoftonline.com; *.live.com; *.azure.net; *.gfx.ms; *.onestore.ms; *.msecnd.net; *.outlookgroups.ms; *.linkedin.com; *.msocdn.com; *.live.net; ihk-aka.de'
NUERA_BASE_URL = 'https://www.ihk-aka.de/fileadmin/AkA/Download/Nuera/'
NUERA_CANDIDATES = (
    'nuera2026_f.zip',
    'nuera2026_h.zip',
    'nuera2025_f.zip',
    'nuera2025_h.zip',
    'nuera2024_f.zip',
    'nuera2024_h.zip',
)


def order_nuera_candidates(file_names: list[str] | tuple[str, ...]) -> list[str]:
    def sort_key(name: str) -> tuple[int, int]:
        match = re.match(r'^nuera(?P<year>\d{4})_(?P<kind>[fh])\.zip$', name, flags=re.IGNORECASE)
        if not match:
            return (0, 0)
        year = int(match.group('year'))
        kind_rank = 2 if match.group('kind').lower() == 'h' else 1
        return (year, kind_rank)

    return sorted(file_names, key=sort_key, reverse=True)
PRIMARY_BLUE = '#0b5ed7'
PRIMARY_BLUE_HOVER = '#0a58ca'
PRIMARY_BLUE_ACTIVE = '#084298'
PROGRESS_GREEN = '#198754'
PROGRESS_RED = '#dc3545'

AP1_PROGRESS_MARKERS: tuple[tuple[str, int], ...] = (
    ('Suche und lade neueste Nuera-Dateien', 10),
    ('Kopiere Nuera-Ordner auf Desktop', 18),
    ('Initialisiere Office', 28),
    ('Setze Office/Windows-Optionen', 36),
    ('Setze Standard-Speicherpfad auf den Benutzerordner', 44),
    ('Uebernehme Autokorrektur-Einstellungen', 52),
    ('Uebernehme Schnellzugriff-Symbolleisten', 60),
    ('Kopiere Vorlagen (Normal.dotm und Mappe.xltx) mit Backup', 68),
    ('Erzeuge Benutzerordner aus Excel', 76),
    ('COM nicht verfuegbar - nutze CSV-Fallback', 76),
    ('Lege Kandidaten-Ordner auf Desktop des aktuellen Nutzers', 86),
    ('Raeume temporaeren Ordner auf', 89),
    ('Taskbar-Einstellungen uebernehmen', 92),
    ('Proxy konfigurieren (falls angegeben)', 96),
    ('Fertig. Der Rechner ist fuer die AP 1 vorbereitet.', 100),
)


def show_error(message: str, title: str = 'AP1-Konfigurator') -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, title, MB_OK | MB_ICONERROR)
    except Exception:
        sys.stderr.write(f'{title}: {message}\n')


def resolve_runtime_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resolve_source_dir() -> Path:
    return Path(__file__).resolve().parent


def load_build_info() -> dict[str, str]:
    candidates = []
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        candidates.append(Path(getattr(sys, '_MEIPASS')) / 'build_info.py')
    candidates.append(resolve_source_dir() / 'build_info.py')
    candidates.append(resolve_runtime_dir() / 'src' / 'build_info.py')

    for candidate in candidates:
        if not candidate.exists():
            continue
        spec = importlib.util.spec_from_file_location('ap1_build_info', candidate)
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module.BUILD_INFO

    raise ModuleNotFoundError('build_info.py konnte weder als Datei noch im Bundle geladen werden.')


BUILD_INFO = load_build_info()
ARTIFACT_NAME = BUILD_INFO['artifact_name']
VERSION = f"v{BUILD_INFO['version']}"


def resolve_bundle_dir() -> Path:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(getattr(sys, '_MEIPASS')).resolve()
    return Path(__file__).resolve().parents[1]


def resolve_local_appdata_root() -> Path:
    local_appdata = os.environ.get('LOCALAPPDATA')
    if not local_appdata:
        local_appdata = str(Path.home() / 'AppData' / 'Local')
    return Path(local_appdata) / ARTIFACT_NAME


def resolve_local_appdata_dir() -> Path:
    return resolve_local_appdata_root() / VERSION


def resolve_current_app_dir() -> Path:
    return resolve_local_appdata_root() / 'current'


def resolve_nuera_root(app_dir: Path, runtime_dir: Path) -> Path:
    runtime_nuera = runtime_dir / 'data' / '3. Nuera-Dateien'
    if runtime_nuera.exists():
        return runtime_nuera
    return app_dir / 'data' / '3. Nuera-Dateien'


def copy_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def sync_directory(source: Path, target: Path) -> None:
    if not source.exists():
        return
    target.mkdir(parents=True, exist_ok=True)
    for item in source.iterdir():
        destination = target / item.name
        if item.is_dir():
            sync_directory(item, destination)
        else:
            copy_file(item, destination)


def sync_embedded_runtime(bundle_dir: Path, app_dir: Path) -> None:
    for file_name in EMBEDDED_FILES:
        source = bundle_dir / file_name
        if source.exists():
            copy_file(source, app_dir / file_name)

    for dir_name in EMBEDDED_DIRS:
        source = bundle_dir / dir_name
        if source.exists():
            sync_directory(source, app_dir / dir_name)


def sync_release_content(runtime_dir: Path, app_dir: Path) -> None:
    for dir_name in PACKAGE_CONTENT_DIRS:
        source = runtime_dir / dir_name
        if source.exists():
            sync_directory(source, app_dir / dir_name)


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def prune_old_version_dirs(root_dir: Path, current_version: str) -> None:
    if not root_dir.exists():
        return
    for child in root_dir.iterdir():
        if not child.is_dir():
            continue
        if child.name == 'current' or child.name == current_version:
            continue
        if not child.name.startswith('v'):
            continue
        try:
            shutil.rmtree(child)
        except PermissionError:
            continue


def prepare_current_alias(source_dir: Path) -> Path:
    current_dir = resolve_current_app_dir()
    try:
        reset_dir(current_dir)
        sync_directory(source_dir, current_dir)
        return current_dir
    except PermissionError:
        return source_dir


def prepare_app_dir(runtime_dir: Path) -> Path:
    bundle_dir = resolve_bundle_dir()
    root_dir = resolve_local_appdata_root()
    app_dir = resolve_local_appdata_dir()
    app_dir.mkdir(parents=True, exist_ok=True)
    sync_embedded_runtime(bundle_dir, app_dir)
    sync_release_content(runtime_dir, app_dir)
    (app_dir / 'data' / '4. Logs').mkdir(parents=True, exist_ok=True)
    prune_old_version_dirs(root_dir, VERSION)
    return prepare_current_alias(app_dir)


def find_launcher(runtime_dir: Path) -> Path:
    candidates = [
        runtime_dir / 'AP1-Konfigurator.ps1',
        runtime_dir / 'AP1-Konfigurator.bat',
        runtime_dir / 'data' / 'AP1-Konfigurator.ps1',
        runtime_dir / 'data' / 'AP1-Konfigurator.bat',
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError('Kein Startskript gefunden. Erwartet wurden AP1-Konfigurator.bat oder AP1-Konfigurator.ps1.')


def build_command(launcher: Path) -> list[str]:
    return build_command_with_args(launcher, sys.argv[1:])


def build_command_with_args(launcher: Path, extra_args: list[str]) -> list[str]:
    suffix = launcher.suffix.lower()
    if suffix == '.bat':
        return ['cmd.exe', '/c', str(launcher), *extra_args]
    if suffix == '.ps1':
        return [
            'powershell.exe',
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            str(launcher),
            *extra_args,
        ]
    raise RuntimeError(f'Nicht unterstütztes Startskript: {launcher.name}')


def open_path(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f'Pfad nicht gefunden: {path}')
    os.startfile(path)


def open_parent(path: Path) -> None:
    open_path(path.parent)


def read_proxy_state() -> tuple[str, str]:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Internet Settings') as key:
            enabled, _ = winreg.QueryValueEx(key, 'ProxyEnable')
            server, _ = winreg.QueryValueEx(key, 'ProxyServer')
        if enabled:
            return 'Aktiv', str(server)
        return 'Deaktiviert', str(server)
    except Exception:
        return 'Unbekannt', '-'


def set_proxy_registry(state: str, server: str = DEFAULT_PROXY_SERVER, bypass: str = DEFAULT_PROXY_BYPASS) -> None:
    import winreg

    reg_path = r'Software\Microsoft\Windows\CurrentVersion\Internet Settings'
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
        if state == 'On':
            winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, 'ProxyServer', 0, winreg.REG_SZ, server)
            winreg.SetValueEx(key, 'ProxyOverride', 0, winreg.REG_SZ, bypass)
            winreg.SetValueEx(key, 'AutoDetect', 0, winreg.REG_DWORD, 0)
        elif state == 'Off':
            winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, 'AutoDetect', 0, winreg.REG_DWORD, 0)
        else:
            raise ValueError(f'Ungültiger Proxy-Status: {state}')

    try:
        INTERNET_OPTION_SETTINGS_CHANGED = 39
        INTERNET_OPTION_REFRESH = 37
        ctypes.windll.Wininet.InternetSetOptionW(0, INTERNET_OPTION_SETTINGS_CHANGED, 0, 0)
        ctypes.windll.Wininet.InternetSetOptionW(0, INTERNET_OPTION_REFRESH, 0, 0)
    except Exception:
        # Registry wurde gesetzt; ein manueller Netzwerk-Refresh wäre dann ggf. nötig.
        pass


def find_latest_folder(base_dir: Path, prefix: str) -> Path | None:
    if not base_dir.exists():
        return None
    candidates = [p for p in base_dir.iterdir() if p.is_dir() and p.name.lower().startswith(prefix.lower())]
    if not candidates:
        return None
    return sorted(candidates, key=lambda p: (p.stat().st_mtime, p.name.lower()))[-1]


def find_latest_log(log_dir: Path) -> Path | None:
    if not log_dir.exists():
        return None
    candidates = [p for p in log_dir.iterdir() if p.is_file() and p.suffix.lower() == '.log']
    if not candidates:
        return None
    return sorted(candidates, key=lambda p: (p.stat().st_mtime, p.name.lower()))[-1]


def resolve_app_icon(app_dir: Path) -> Path | None:
    candidates = [
        app_dir / 'app_icon.ico',
        resolve_bundle_dir() / 'app_icon.ico',
        resolve_runtime_dir() / 'src' / 'app_icon.ico',
        resolve_source_dir() / 'app_icon.ico',
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _cleanup_nuera_variants(download_dir: Path, keep_name: str | None = None) -> None:
    for existing in download_dir.iterdir():
        match = re.fullmatch(r'nuera\d{4}_[fh](?:\.zip)?', existing.name, flags=re.IGNORECASE)
        if not match or (keep_name and existing.name.casefold() == keep_name.casefold()):
            continue
        if existing.is_dir():
            shutil.rmtree(existing)
        else:
            existing.unlink()


def _use_local_nuera(download_dir: Path) -> tuple[bool, str]:
    for file_name in order_nuera_candidates(NUERA_CANDIDATES):
        folder_name = Path(file_name).stem
        target_zip = download_dir / file_name
        target_folder = download_dir / folder_name
        has_content = target_folder.is_dir() and any(path.is_file() for path in target_folder.rglob('*'))
        if has_content:
            _cleanup_nuera_variants(download_dir, folder_name)
            return True, f'{folder_name} bereits lokal vorhanden (kein Download erforderlich)'
        if target_folder.exists():
            shutil.rmtree(target_folder)
        if target_zip.is_file():
            try:
                with zipfile.ZipFile(target_zip, 'r') as zip_ref:
                    zip_ref.extractall(download_dir)
                if target_folder.is_dir() and any(path.is_file() for path in target_folder.rglob('*')):
                    _cleanup_nuera_variants(download_dir, folder_name)
                    return True, f'{folder_name}.zip lokal vorhanden und neu entpackt'
            except (OSError, zipfile.BadZipFile) as exc:
                continue
    return False, ''


def download_and_extract_latest_nuera(download_dir: Path) -> tuple[bool, str]:
    download_dir.mkdir(parents=True, exist_ok=True)
    local_ok, local_detail = _use_local_nuera(download_dir)
    if local_ok:
        return True, local_detail

    _cleanup_nuera_variants(download_dir)
    headers = {'User-Agent': 'AP1-Konfigurator/1.0 (+Windows)'}

    last_error = ''
    for file_name in order_nuera_candidates(NUERA_CANDIDATES):
        url = f'{NUERA_BASE_URL}{file_name}'
        target_zip = download_dir / file_name
        try:
            request = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(request, timeout=30) as response:
                data = response.read()
            target_zip.write_bytes(data)

            folder_name = Path(file_name).stem
            target_folder = download_dir / folder_name
            if target_folder.exists():
                shutil.rmtree(target_folder)

            with zipfile.ZipFile(target_zip, 'r') as zip_ref:
                zip_ref.extractall(download_dir)

            return True, folder_name
        except urllib.error.HTTPError as exc:
            if exc.code == 404:
                continue
            last_error = f'HTTP {exc.code} bei {file_name}'
        except Exception as exc:
            last_error = f'{file_name}: {exc}'

    if last_error:
        return False, last_error
    return False, 'Keine passende Nüra-Datei online gefunden.'


class AP1ConfiguratorGUI(tk.Tk):
    def __init__(self, app_dir: Path, runtime_dir: Path) -> None:
        super().__init__()
        self.app_dir = app_dir
        self.runtime_dir = runtime_dir
        self.title(f'AP1-Konfigurator – {VERSION}')
        self.geometry('1120x760')
        self.minsize(980, 680)

        self.proxy_mode = tk.StringVar(value='Skip')
        self.status_text = tk.StringVar(value='Bereit.')
        self.refresh_note = tk.StringVar(value='')
        self.progress_text = tk.StringVar(value='0 %')
        self._nuera_update_running = False
        self._ap1_process: subprocess.Popen | None = None
        self._ap1_progress = tk.DoubleVar(value=0)
        self._ap1_log_path: Path | None = None
        self._ap1_monitor_after: str | None = None

        self.paths = {
            'ap1_tn': self.app_dir / 'data' / '1. Anpassen' / 'AP1-TN.xlsx',
            'normal_dotm': self.app_dir / 'data' / '2. Bei Bedarf anpassen' / 'Word' / 'Normal.dotm',
            'mappe_xltx': self.app_dir / 'data' / '2. Bei Bedarf anpassen' / 'Excel' / 'Mappe.xltx',
            'nuera_root': resolve_nuera_root(self.app_dir, self.runtime_dir),
            'logs_root': self.app_dir / 'data' / '4. Logs',
            'script': self.app_dir / 'AP1-Konfigurator.ps1',
        }

        self._apply_window_icon()
        self._build_style()
        self._build_layout()
        self.refresh_status(download_nuera=False)

    def _apply_window_icon(self) -> None:
        icon_path = resolve_app_icon(self.app_dir)
        if not icon_path:
            return
        try:
            self.iconbitmap(default=str(icon_path))
        except tk.TclError:
            pass

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use('vista')
        except tk.TclError:
            pass
        style.configure('Header.TLabel', font=('Segoe UI', 18, 'bold'))
        style.configure('SubHeader.TLabel', font=('Segoe UI', 10))
        style.configure('Status.TLabel', font=('Segoe UI', 10))
        style.configure('StatusValue.TLabel', font=('Segoe UI', 10), foreground='#212529')
        style.configure('StatusError.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#dc3545')
        style.configure(
            'Progress.Running.Horizontal.TProgressbar',
            troughcolor='#e9ecef',
            background=PRIMARY_BLUE,
            lightcolor=PRIMARY_BLUE,
            darkcolor=PRIMARY_BLUE,
            bordercolor='#ced4da',
        )
        style.configure(
            'Progress.Done.Horizontal.TProgressbar',
            troughcolor='#e9ecef',
            background=PROGRESS_GREEN,
            lightcolor=PROGRESS_GREEN,
            darkcolor=PROGRESS_GREEN,
            bordercolor='#ced4da',
        )
        style.configure(
            'Progress.Error.Horizontal.TProgressbar',
            troughcolor='#e9ecef',
            background=PROGRESS_RED,
            lightcolor=PROGRESS_RED,
            darkcolor=PROGRESS_RED,
            bordercolor='#ced4da',
        )

    def _build_layout(self) -> None:
        container = ttk.Frame(self, padding=16)
        container.pack(fill='both', expand=True)

        header = ttk.Frame(container)
        header.pack(fill='x')
        ttk.Label(header, text='AP1-Konfigurator', style='Header.TLabel').pack(anchor='w')
        ttk.Label(
            header,
            text='Dateien prüfen, Vorlagen öffnen, Nüra-Version kontrollieren, Logs ansehen und Proxy/Start aus der GUI steuern.',
            style='SubHeader.TLabel',
        ).pack(anchor='w', pady=(4, 0))

        top = ttk.Frame(container)
        top.pack(fill='x', pady=(14, 8))

        status_frame = ttk.LabelFrame(top, text='Status', padding=12)
        status_frame.pack(side='left', fill='both', expand=True)

        self.status_rows: dict[str, tk.StringVar] = {}
        self.status_value_labels: dict[str, ttk.Label] = {}
        for row, (key, label) in enumerate([
            ('app_dir', 'Arbeitsordner'),
            ('ap1_tn', 'AP1-TN.xlsx'),
            ('normal_dotm', 'Normal.dotm'),
            ('mappe_xltx', 'Mappe.xltx'),
            ('nuera', 'Aktuellste Nüra-Version'),
            ('logs', 'Logs'),
            ('proxy', 'Proxy'),
        ]):
            ttk.Label(status_frame, text=f'{label}:').grid(row=row, column=0, sticky='w', padx=(0, 12), pady=2)
            value = tk.StringVar(value='-')
            self.status_rows[key] = value
            value_label = ttk.Label(status_frame, textvariable=value, style='StatusValue.TLabel')
            value_label.grid(row=row, column=1, sticky='w', pady=2)
            self.status_value_labels[key] = value_label

        right_controls = ttk.Frame(top)
        right_controls.pack(side='right', fill='y', padx=(12, 0))

        proxy_frame = ttk.LabelFrame(right_controls, text='Proxy', padding=12)
        proxy_frame.pack(fill='x')
        ttk.Label(proxy_frame, text='Modus:').pack(anchor='w')
        mode_box = ttk.Combobox(proxy_frame, textvariable=self.proxy_mode, values=['Skip', 'On', 'Off'], state='readonly', width=10)
        mode_box.pack(anchor='w', pady=(4, 10))

        nuera_frame = ttk.LabelFrame(right_controls, text='Nüra-Ordner', padding=12)
        nuera_frame.pack(fill='x', pady=(10, 0))
        self.refresh_button = ttk.Button(nuera_frame, text='Aktualisieren', command=self.refresh_status_with_download)
        self.refresh_button.pack(fill='x', pady=2)

        middle = ttk.Frame(container)
        middle.pack(fill='both', expand=True, pady=(4, 8))

        left_area = ttk.Frame(middle)
        left_area.pack(side='left', fill='both', expand=True)

        steps_frame = ttk.LabelFrame(left_area, text='Schritt-für-Schritt', padding=12)
        steps_frame.pack(fill='both', expand=True)
        steps_text = tk.Text(steps_frame, height=14, wrap='word', relief='flat', background=self.cget('bg'))
        steps_text.pack(fill='both', expand=True)
        steps_text.insert('end',
            '1. AP1-TN.xlsx öffnen und anpassen.\n'
            '2. [optional] Bei Problemen Normal.dotm und Mappe.xltx öffnen und anpassen.\n'
            '3. Proxy-Modus über die Dropdownliste auswählen.\n'
            '4. AP1-Konfigurator mit dem gewünschten Proxy-Modus starten.\n'
        )
        steps_text.configure(state='disabled')

        start_area = ttk.Frame(left_area, padding=(0, 10, 0, 0))
        start_area.pack(fill='x')
        self.start_button = tk.Button(
            start_area,
            text='AP1-Konfiguration starten',
            command=self.start_ap1,
            bg=PRIMARY_BLUE,
            fg='white',
            activebackground=PRIMARY_BLUE_ACTIVE,
            activeforeground='white',
            relief='flat',
            bd=0,
            highlightthickness=0,
            padx=12,
            pady=8,
            font=('Segoe UI', 10, 'bold'),
            cursor='hand2',
        )
        self.start_button.pack(fill='x')
        self.start_button.bind('<Enter>', lambda _: self.start_button.configure(bg=PRIMARY_BLUE_HOVER))
        self.start_button.bind('<Leave>', lambda _: self.start_button.configure(bg=PRIMARY_BLUE))

        actions_frame = ttk.LabelFrame(middle, text='Dateien & Ordner', padding=12)
        actions_frame.pack(side='right', fill='y', padx=(12, 0))
        actions_frame.columnconfigure(0, weight=1)
        actions_frame.columnconfigure(1, weight=1)

        ttk.Button(actions_frame, text='AP1-TN.xlsx öffnen', command=lambda: self.open_item(self.paths['ap1_tn'])).grid(row=0, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(actions_frame, text='Normal.dotm öffnen', command=lambda: self.open_item(self.paths['normal_dotm'])).grid(row=0, column=1, sticky='ew', padx=2, pady=2)
        ttk.Button(actions_frame, text='Mappe.xltx öffnen', command=lambda: self.open_item(self.paths['mappe_xltx'])).grid(row=1, column=0, sticky='ew', padx=2, pady=2)
        ttk.Button(actions_frame, text='Letztes Log öffnen', command=self.open_latest_log).grid(row=1, column=1, sticky='ew', padx=2, pady=2)

        footer = ttk.Frame(container)
        footer.pack(fill='x')
        ttk.Separator(footer).pack(fill='x', pady=(0, 6))

        progress_row = ttk.Frame(footer)
        progress_row.pack(fill='x', pady=(0, 6))
        self.progress_bar = ttk.Progressbar(
            progress_row,
            orient='horizontal',
            mode='determinate',
            maximum=100,
            variable=self._ap1_progress,
            style='Progress.Running.Horizontal.TProgressbar',
        )
        self.progress_bar.pack(side='left', fill='x', expand=True)
        ttk.Label(progress_row, textvariable=self.progress_text, width=14, anchor='e').pack(side='left', padx=(10, 0))

        ttk.Label(footer, textvariable=self.status_text, style='Status.TLabel').pack(anchor='w')
        ttk.Label(footer, textvariable=self.refresh_note, style='SubHeader.TLabel').pack(anchor='w')

    def set_status(self, text: str) -> None:
        self.status_text.set(text)
        self.update_idletasks()

    def set_note(self, text: str) -> None:
        self.refresh_note.set(text)

    def set_progress(self, value: float, text: str | None = None, style: str = 'Progress.Running.Horizontal.TProgressbar') -> None:
        clamped = max(0.0, min(100.0, float(value)))
        self._ap1_progress.set(clamped)
        self.progress_bar.configure(style=style)
        if text is None:
            self.progress_text.set(f'{clamped:.0f} %')
        else:
            self.progress_text.set(text)
        self.update_idletasks()

    def open_item(self, path: Path) -> None:
        try:
            open_path(path)
            self.set_status(f'Geöffnet: {path.name}')
        except Exception as exc:
            messagebox.showerror('AP1-Konfigurator', str(exc), parent=self)

    def open_latest_nuera(self) -> None:
        latest = find_latest_folder(self.paths['nuera_root'], 'nuera')
        if latest is None:
            messagebox.showinfo('AP1-Konfigurator', 'Keine Nüra-Version gefunden.', parent=self)
            return
        self.open_item(latest)

    def open_logs(self) -> None:
        self.open_item(self.paths['logs_root'])

    def open_latest_log(self) -> None:
        latest = find_latest_log(self.paths['logs_root'])
        if latest is None:
            messagebox.showinfo('AP1-Konfigurator', 'Noch keine Logdatei vorhanden.', parent=self)
            return
        self.open_item(latest)

    def open_release_root(self) -> None:
        self.open_item(self.app_dir)

    def run_proxy_action(self, mode: str) -> None:
        try:
            set_proxy_registry(mode)
            self.proxy_mode.set(mode)
            self.refresh_status(download_nuera=False)
            self.set_status(f'Proxy gesetzt: {mode}')
        except Exception as exc:
            messagebox.showerror('AP1-Konfigurator', str(exc), parent=self)

    def activate_proxy(self) -> None:
        self.run_proxy_action('On')

    def deactivate_proxy(self) -> None:
        self.run_proxy_action('Off')

    def start_ap1(self) -> None:
        if self._ap1_process and self._ap1_process.poll() is None:
            messagebox.showinfo('AP1-Konfigurator', 'Die AP1-Konfiguration läuft bereits.', parent=self)
            return

        try:
            launcher = find_launcher(self.app_dir)
            mode = self.proxy_mode.get() or 'Skip'
            command = build_command_with_args(launcher, ['-Proxy', mode, '-Quiet'])
            self._ap1_process = subprocess.Popen(command, cwd=str(self.app_dir), creationflags=CREATE_NO_WINDOW)
            self._ap1_log_path = None
            self.start_button.configure(state='disabled')
            self.set_progress(3, text='Läuft ...', style='Progress.Running.Horizontal.TProgressbar')
            self.set_status(f'AP1-Konfiguration gestartet ({mode}) ...')
            self._schedule_progress_monitor()
        except Exception as exc:
            messagebox.showerror('AP1-Konfigurator', str(exc), parent=self)

    def _schedule_progress_monitor(self) -> None:
        if self._ap1_monitor_after is not None:
            self.after_cancel(self._ap1_monitor_after)
        self._ap1_monitor_after = self.after(800, self._poll_ap1_progress)

    def _find_current_log_path(self) -> Path | None:
        if self._ap1_log_path and self._ap1_log_path.exists():
            return self._ap1_log_path
        latest = find_latest_log(self.paths['logs_root'])
        if latest and latest.exists():
            self._ap1_log_path = latest
            return latest
        return None

    def _derive_progress_from_log(self) -> float:
        log_path = self._find_current_log_path()
        if not log_path:
            return float(self._ap1_progress.get())

        try:
            content = log_path.read_text(encoding='utf-8', errors='ignore')
        except Exception:
            return float(self._ap1_progress.get())

        derived = float(self._ap1_progress.get())
        for marker, percent in AP1_PROGRESS_MARKERS:
            if marker in content:
                derived = max(derived, float(percent))
        return derived

    def _poll_ap1_progress(self) -> None:
        self._ap1_monitor_after = None
        proc = self._ap1_process
        if proc is None:
            return

        derived = self._derive_progress_from_log()
        current = float(self._ap1_progress.get())
        if proc.poll() is None:
            derived = min(95.0, derived)
        if derived != current:
            self.set_progress(derived)

        return_code = proc.poll()
        if return_code is None:
            self._schedule_progress_monitor()
            return

        self.start_button.configure(state='normal')
        self.refresh_status(download_nuera=False)

        if return_code == 0:
            self.set_progress(100, text='Fertig', style='Progress.Done.Horizontal.TProgressbar')
            self.set_status('Fertig.')
        else:
            self.set_progress(100, text='Fehler', style='Progress.Error.Horizontal.TProgressbar')
            self.set_status(f'AP1-Konfiguration beendet mit Fehler (Code {return_code}).')

    def refresh_status_with_download(self) -> None:
        if self._nuera_update_running:
            return

        self._nuera_update_running = True
        self.refresh_button.state(['disabled'])
        self.set_status('Prüfe lokalen Nüra-Bestand und lade nur bei Bedarf ...')

        worker = threading.Thread(target=self._download_nuera_worker, daemon=True)
        worker.start()

    def _download_nuera_worker(self) -> None:
        ok, detail = download_and_extract_latest_nuera(self.paths['nuera_root'])
        if ok:
            app_nuera = self.app_dir / 'data' / '3. Nuera-Dateien'
            if self.paths['nuera_root'] != app_nuera:
                try:
                    if app_nuera.exists():
                        shutil.rmtree(app_nuera)
                    sync_directory(self.paths['nuera_root'], app_nuera)
                except OSError as exc:
                    ok = False
                    detail = f'Nüra wurde gespeichert, aber nicht in den Laufzeitordner synchronisiert: {exc}'
        self.after(0, lambda: self._on_download_nuera_done(ok, detail))

    def _on_download_nuera_done(self, ok: bool, detail: str) -> None:
        try:
            if ok:
                self.set_note(f'Nüra-Update: {detail} wurde heruntergeladen/aktualisiert.')
            else:
                self.set_note(f'Nüra-Update: {detail}')
            self.refresh_status(download_nuera=False)
        finally:
            self._nuera_update_running = False
            self.refresh_button.state(['!disabled'])

    def _set_status_row(self, key: str, value: str, ok: bool = True) -> None:
        self.status_rows[key].set(value)
        self.status_value_labels[key].configure(
            style='StatusValue.TLabel' if ok else 'StatusError.TLabel'
        )

    def refresh_status(self, download_nuera: bool = False) -> None:
        self.set_status('Status wird aktualisiert ...')

        if download_nuera:
            self.refresh_status_with_download()
            return

        app_dir = self.app_dir
        self._set_status_row('app_dir', str(app_dir), app_dir.exists())
        ap1_tn_ok = self.paths['ap1_tn'].exists()
        normal_dotm_ok = self.paths['normal_dotm'].exists()
        mappe_xltx_ok = self.paths['mappe_xltx'].exists()
        self._set_status_row('ap1_tn', 'OK' if ap1_tn_ok else 'Fehlt', ap1_tn_ok)
        self._set_status_row('normal_dotm', 'OK' if normal_dotm_ok else 'Fehlt', normal_dotm_ok)
        self._set_status_row('mappe_xltx', 'OK' if mappe_xltx_ok else 'Fehlt', mappe_xltx_ok)

        latest_nuera = find_latest_folder(self.paths['nuera_root'], 'nuera')
        self._set_status_row('nuera', latest_nuera.name if latest_nuera else 'Nicht gefunden', latest_nuera is not None)

        log_files = [p for p in self.paths['logs_root'].iterdir()] if self.paths['logs_root'].exists() else []
        latest_log = find_latest_log(self.paths['logs_root'])
        if latest_log:
            self._set_status_row('logs', f'{len(log_files)} Dateien · {latest_log.name}')
        else:
            self._set_status_row('logs', '0 Dateien', False)

        proxy_state, proxy_server = read_proxy_state()
        self._set_status_row('proxy', f'{proxy_state} · {proxy_server}')
        if not download_nuera:
            self.set_note(f'Letzte Aktualisierung: {datetime.now().strftime("%d.%m.%Y %H:%M:%S")}')
        self.set_status('Bereit.')


def main() -> int:
    runtime_dir = resolve_runtime_dir()
    try:
        app_dir = prepare_app_dir(runtime_dir)
        gui = AP1ConfiguratorGUI(app_dir, runtime_dir)
        gui.mainloop()
        return 0
    except Exception as exc:
        show_error(str(exc))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
