#!/usr/bin/env python3
"""
Hebrew Subtitle Generator — Installer
Checks and installs all required dependencies, then creates a desktop shortcut.
Run via INSTALL.bat  or directly:  python installer.py
"""

import importlib
import importlib.util
import os
import shutil
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path
from tkinter import ttk
import tkinter as tk

APP_DIR = Path(__file__).parent.resolve()

# ── Package list (pip_name, import_name, min_version) ────────────────────────
PIP_PACKAGES = [
    ('pysubs2',         'pysubs2',        '1.7.0'),
    ('deep-translator', 'deep_translator','1.11.4'),
    ('opencv-python',   'cv2',            '4.9.0'),
    ('pytesseract',     'pytesseract',    '0.3.10'),
    ('chardet',         'chardet',        '5.2.0'),
    ('ffsubsync',       'ffsubsync',      '0.4.26'),
    ('Pillow',          'PIL',            '10.0.0'),
    ('tkinterdnd2',     'tkinterdnd2',    '0.3.0'),
]

# ── winget package IDs ────────────────────────────────────────────────────────
WINGET_FFMPEG    = 'Gyan.FFmpeg'
WINGET_TESSERACT = 'UB-Mannheim.TesseractOCR'
WINGET_OLLAMA    = 'Ollama.Ollama'

# Known install paths (winget may not update PATH immediately)
TESSERACT_KNOWN = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
]
FFMPEG_KNOWN = [
    r'C:\Program Files\ffmpeg\bin\ffmpeg.exe',
    r'C:\ProgramData\chocolatey\bin\ffmpeg.exe',
    r'C:\tools\ffmpeg\bin\ffmpeg.exe',
]

# Download URLs shown when winget fails
URL_FFMPEG    = 'https://ffmpeg.org/download.html#build-windows'
URL_TESSERACT = 'https://github.com/UB-Mannheim/tesseract/wiki'
URL_OLLAMA    = 'https://ollama.com/download'

_NO_WIN = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0

# ── Step status constants ─────────────────────────────────────────────────────
PENDING  = 'pending'
CHECKING = 'checking'
RUNNING  = 'running'
DONE     = 'done'
SKIPPED  = 'skipped'
MANUAL   = 'manual'
ERROR    = 'error'

STATUS_COLOR = {
    PENDING:  '#888888',
    CHECKING: '#0078d4',
    RUNNING:  '#0078d4',
    DONE:     '#107c10',
    SKIPPED:  '#107c10',
    MANUAL:   '#ca5010',
    ERROR:    '#d13438',
}
STATUS_TEXT = {
    PENDING:  'Waiting',
    CHECKING: 'Checking…',
    RUNNING:  'Installing…',
    DONE:     'Installed  ✓',
    SKIPPED:  'Already installed  ✓',
    MANUAL:   'Manual install needed',
    ERROR:    'Failed',
}


# ── System helpers ────────────────────────────────────────────────────────────

def refresh_path():
    """Re-read PATH from Windows registry into the current process."""
    if sys.platform != 'win32':
        return
    try:
        import winreg
        paths = []
        for hive, sub in [
            (winreg.HKEY_LOCAL_MACHINE,
             r'SYSTEM\CurrentControlSet\Control\Session Manager\Environment'),
            (winreg.HKEY_CURRENT_USER, r'Environment'),
        ]:
            try:
                with winreg.OpenKey(hive, sub) as k:
                    val, _ = winreg.QueryValueEx(k, 'Path')
                    paths.extend(p.strip() for p in val.split(';') if p.strip())
            except OSError:
                pass
        if paths:
            seen  = set(paths)
            extra = [p for p in os.environ.get('PATH', '').split(';')
                     if p and p not in seen]
            os.environ['PATH'] = ';'.join(paths + extra)
    except Exception:
        pass


def find_binary(name: str, known: list = None) -> str:
    """Return full path to binary, searching PATH then known locations."""
    found = shutil.which(name)
    if found:
        return found
    for p in (known or []):
        if Path(p).exists():
            os.environ['PATH'] = str(Path(p).parent) + ';' + os.environ.get('PATH', '')
            return p
    return ''


def winget_available() -> bool:
    try:
        r = subprocess.run(['winget', '--version'], capture_output=True,
                           timeout=5, creationflags=_NO_WIN)
        return r.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def winget_install(pkg_id: str, log_cb) -> bool:
    """Run winget install silently. Returns True on success."""
    try:
        r = subprocess.run(
            ['winget', 'install', pkg_id,
             '--accept-package-agreements', '--accept-source-agreements',
             '--silent', '--force'],
            capture_output=True, text=True, timeout=360, creationflags=_NO_WIN,
        )
        for line in (r.stdout + r.stderr).splitlines():
            s = line.strip()
            if s and len(s) > 2:
                log_cb(f'  {s}')
        refresh_path()
        return r.returncode == 0
    except subprocess.TimeoutExpired:
        log_cb('  Timed out (>6 min) — try installing manually')
        return False
    except Exception as e:
        log_cb(f'  winget error: {e}')
        return False


# ── GUI Installer ─────────────────────────────────────────────────────────────

class InstallerApp(tk.Tk):

    REQUIRED_STEPS = [
        dict(id='packages',  label='Python packages',
             detail='pysubs2, deep-translator, opencv, pytesseract, ffsubsync…'),
        dict(id='ffmpeg',    label='ffmpeg + ffprobe',
             detail='Video subtitle extraction and probing',
             url=URL_FFMPEG),
        dict(id='tesseract', label='Tesseract OCR',
             detail='Hard-coded subtitle recognition',
             url=URL_TESSERACT),
        dict(id='shortcut',  label='Desktop shortcut',
             detail='Launcher on your Windows Desktop'),
    ]

    def __init__(self):
        super().__init__()
        self.title('Hebrew Subtitle Generator — Setup')
        self.geometry('640x620')
        self.resizable(False, False)
        self._widgets = {}          # step_id → {dot, status_lbl, dl_btn, opt_btn}
        self._installing = False
        self._build_ui()
        # Auto-start installation 800 ms after the window appears
        self.after(800, self._start_install)

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Blue header ──
        hdr = tk.Frame(self, bg='#0078d4', height=56)
        hdr.pack(fill='x')
        hdr.pack_propagate(False)
        tk.Label(hdr, text=' Hebrew Subtitle Generator',
                 bg='#0078d4', fg='white',
                 font=('Segoe UI', 14, 'bold')).pack(side='left', padx=14, pady=12)
        tk.Label(hdr, text='Setup',
                 bg='#0078d4', fg='#c8e6ff',
                 font=('Segoe UI', 11)).pack(side='left', pady=12)

        # ── Content ──
        body = ttk.Frame(self, padding=(18, 14, 18, 8))
        body.pack(fill='both', expand=True)

        # Required section
        self._section_label(body, 'REQUIRED COMPONENTS')
        for step in self.REQUIRED_STEPS:
            self._step_row(body, **step)

        ttk.Separator(body, orient='horizontal').pack(fill='x', pady=10)

        # Optional section
        self._section_label(body, 'OPTIONAL — AI Translation (gender-aware Hebrew)')
        self._step_row(body, id='ollama', label='Ollama  (free local AI)',
                       detail='Better Hebrew translation — gender-aware, context-aware',
                       url=URL_OLLAMA, opt_label='Install Ollama')

        ttk.Separator(body, orient='horizontal').pack(fill='x', pady=(10, 6))

        # Progress
        self.progress_lbl = ttk.Label(body, text='', foreground='#555')
        self.progress_lbl.pack(anchor='w')
        self.progress = ttk.Progressbar(body, maximum=4, value=0, length=600)
        self.progress.pack(fill='x', pady=(2, 8))

        # Log
        log_frame = ttk.LabelFrame(body, text='Log', padding=(6, 4))
        log_frame.pack(fill='both', expand=True)

        self.log_box = tk.Text(
            log_frame, height=7, state=tk.DISABLED, wrap=tk.WORD,
            font=('Consolas', 8), bg='#1e1e1e', fg='#d4d4d4', relief='flat',
        )
        log_sb = ttk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=log_sb.set)
        self.log_box.pack(side='left', fill='both', expand=True)
        log_sb.pack(side='right', fill='y')

        # Buttons
        btn = ttk.Frame(self, padding=(18, 6, 18, 12))
        btn.pack(fill='x')
        self.install_btn_lbl = ttk.Label(btn, text='Starting installation…',
                                         foreground='#0078d4',
                                         font=('Segoe UI', 9))
        self.install_btn_lbl.pack(side='left')
        ttk.Label(btn, text='').pack(side='left', expand=True)
        self.close_btn = ttk.Button(btn, text='Close',
                                    command=self.destroy, width=10,
                                    state=tk.DISABLED)
        self.close_btn.pack(side='right')

    def _section_label(self, parent, text):
        ttk.Label(parent, text=text,
                  font=('Segoe UI', 8, 'bold'),
                  foreground='#444').pack(anchor='w', pady=(0, 4))

    def _step_row(self, parent, id, label, detail, url='', opt_label=''):
        row = ttk.Frame(parent)
        row.pack(fill='x', pady=2)

        # Coloured status dot
        dot = tk.Label(row, text='●', font=('Segoe UI', 13),
                       foreground=STATUS_COLOR[PENDING], width=2)
        dot.pack(side='left', padx=(0, 8))

        # Name + description
        txt = ttk.Frame(row)
        txt.pack(side='left', fill='x', expand=True)
        ttk.Label(txt, text=label,
                  font=('Segoe UI', 10, 'bold')).pack(anchor='w')
        ttk.Label(txt, text=detail,
                  font=('Segoe UI', 8), foreground='#888').pack(anchor='w')

        # Right side: status text + buttons
        right = ttk.Frame(row)
        right.pack(side='right', padx=(6, 0))

        status_lbl = ttk.Label(right, text=STATUS_TEXT[PENDING],
                                foreground=STATUS_COLOR[PENDING],
                                font=('Segoe UI', 9), width=24, anchor='e')
        status_lbl.pack(side='left')

        dl_btn = None
        if url:
            dl_btn = ttk.Button(right, text='Download',
                                command=lambda u=url: webbrowser.open(u), width=9)
            # shown only when status == MANUAL

        opt_btn = None
        if opt_label:
            opt_btn = ttk.Button(right, text=opt_label,
                                 command=lambda i=id: self._optional_clicked(i),
                                 width=15)
            opt_btn.pack(side='left', padx=(6, 0))

        self._widgets[id] = dict(dot=dot, status_lbl=status_lbl,
                                 dl_btn=dl_btn, opt_btn=opt_btn)

    # ── State updates (all thread-safe) ──────────────────────────────────────

    def set_step(self, step_id: str, status: str, text: str = ''):
        self.after(0, self._set_step_ui, step_id, status, text)

    def _set_step_ui(self, step_id: str, status: str, text: str):
        w = self._widgets.get(step_id)
        if not w:
            return
        color = STATUS_COLOR.get(status, '#888')
        label = text or STATUS_TEXT.get(status, status)
        w['dot'].configure(foreground=color)
        w['status_lbl'].configure(text=label, foreground=color)
        if w['dl_btn']:
            if status == MANUAL:
                w['dl_btn'].pack(side='left', padx=(6, 0))
            else:
                w['dl_btn'].pack_forget()

    def log(self, msg: str):
        self.after(0, self._write_log, msg)

    def _write_log(self, msg: str):
        self.log_box.configure(state=tk.NORMAL)
        self.log_box.insert(tk.END, msg + '\n')
        self.log_box.see(tk.END)
        self.log_box.configure(state=tk.DISABLED)

    def set_progress(self, value: int, text: str = ''):
        self.after(0, lambda: self.progress.configure(value=value))
        if text:
            self.after(0, lambda: self.progress_lbl.configure(text=text))

    # ── Install orchestration ─────────────────────────────────────────────────

    def _start_install(self):
        if self._installing:
            return
        self._installing = True
        self.close_btn.configure(state=tk.DISABLED)
        self.after(0, lambda: self.install_btn_lbl.configure(text='Installing — please wait…'))
        threading.Thread(target=self._run_all, daemon=True).start()

    def _optional_clicked(self, step_id: str):
        w = self._widgets.get(step_id)
        if w and w['opt_btn']:
            w['opt_btn'].configure(state=tk.DISABLED)
        threading.Thread(target=self._run_ollama, daemon=True).start()

    def _run_all(self):
        steps = [
            ('packages',  self._do_packages),
            ('ffmpeg',    self._do_ffmpeg),
            ('tesseract', self._do_tesseract),
            ('shortcut',  self._do_shortcut),
        ]
        passed = 0
        for i, (sid, fn) in enumerate(steps):
            ok = fn()
            if ok:
                passed += 1
            self.set_progress(i + 1, f'Step {i + 1} / {len(steps)}')

        self.log('─' * 52)
        if passed == len(steps):
            self.log('All done!  Launch the app from your Desktop.')
            self.set_progress(4, 'Installation complete!')
            self.after(0, lambda: self.install_btn_lbl.configure(
                text='Installation complete!', foreground='#107c10'))
        else:
            self.log(f'{passed}/{len(steps)} steps succeeded — see orange items above.')
            self.set_progress(4, 'Finished (some steps need manual action — see orange items)')
            self.after(0, lambda: self.install_btn_lbl.configure(
                text='Finished — some steps need manual action (see orange items)',
                foreground='#ca5010'))

        self.after(0, lambda: self.close_btn.configure(state=tk.NORMAL))

    # ── Individual install steps ──────────────────────────────────────────────

    def _do_packages(self) -> bool:
        self.set_step('packages', CHECKING)
        self.log('Checking Python packages…')

        missing_specs = []
        for pip_name, import_name, min_ver in PIP_PACKAGES:
            try:
                importlib.import_module(import_name)
            except ImportError:
                missing_specs.append(f'{pip_name}>={min_ver}')

        if not missing_specs:
            self.log('  All packages already installed.')
            self.set_step('packages', SKIPPED)
            return True

        self.log(f'  Missing: {", ".join(s.split(">=")[0] for s in missing_specs)}')
        self.set_step('packages', RUNNING)

        r = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', '--upgrade'] + missing_specs,
            capture_output=True, text=True, timeout=600, creationflags=_NO_WIN,
        )
        for line in r.stdout.splitlines():
            s = line.strip()
            if 'Successfully installed' in s or 'already satisfied' in s:
                self.log(f'  {s}')
        if r.returncode != 0:
            self.log(f'  pip failed:\n{r.stderr[-500:]}')
            self.set_step('packages', ERROR)
            return False

        self.set_step('packages', DONE)
        return True

    def _do_ffmpeg(self) -> bool:
        self.set_step('ffmpeg', CHECKING)
        self.log('Checking ffmpeg…')

        if find_binary('ffmpeg', FFMPEG_KNOWN):
            self.log(f'  Found: {shutil.which("ffmpeg") or "known path"}')
            self.set_step('ffmpeg', SKIPPED)
            return True

        self.set_step('ffmpeg', RUNNING)

        if winget_available():
            self.log(f'  Running: winget install {WINGET_FFMPEG}')
            if winget_install(WINGET_FFMPEG, self.log):
                if find_binary('ffmpeg', FFMPEG_KNOWN):
                    self.set_step('ffmpeg', DONE)
                    return True
            self.log('  winget did not find ffmpeg in PATH after install.')
            self.log('  It may need a system restart, or install manually.')

        self.log(f'  Download: {URL_FFMPEG}')
        self.log('  Extract and add the bin\\ folder to your system PATH.')
        self.set_step('ffmpeg', MANUAL)
        return False

    def _do_tesseract(self) -> bool:
        self.set_step('tesseract', CHECKING)
        self.log('Checking Tesseract OCR…')

        if find_binary('tesseract', TESSERACT_KNOWN):
            self.log(f'  Found: {find_binary("tesseract", TESSERACT_KNOWN)}')
            self.set_step('tesseract', SKIPPED)
            return True

        self.set_step('tesseract', RUNNING)

        if winget_available():
            self.log(f'  Running: winget install {WINGET_TESSERACT}')
            if winget_install(WINGET_TESSERACT, self.log):
                if find_binary('tesseract', TESSERACT_KNOWN):
                    self.set_step('tesseract', DONE)
                    return True
            self.log('  winget did not confirm Tesseract after install.')

        self.log(f'  Download: {URL_TESSERACT}')
        self.log('  Install and make sure tesseract.exe is in your PATH.')
        self.set_step('tesseract', MANUAL)
        return False

    def _do_shortcut(self) -> bool:
        self.set_step('shortcut', RUNNING)
        self.log('Creating desktop shortcut…')
        try:
            sc_file = APP_DIR / 'create_shortcut.py'
            spec = importlib.util.spec_from_file_location('create_shortcut', sc_file)
            mod  = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)

            icon = APP_DIR / 'icon.ico'
            if not icon.exists():
                mod.create_icon(str(icon))

            ok = mod.create_shortcut(
                sys.executable,
                str(APP_DIR / 'subtitle_app.py'),
                str(icon),
                'Hebrew Subtitle Generator',
            )
            if ok:
                self.set_step('shortcut', DONE)
                self.log('  Shortcut added to Desktop.')
                return True
        except Exception as e:
            self.log(f'  Error: {e}')

        self.set_step('shortcut', ERROR)
        return False

    def _run_ollama(self):
        self.set_step('ollama', CHECKING)
        self.log('Checking Ollama…')

        if shutil.which('ollama'):
            self.log('  Ollama already installed.')
            self.set_step('ollama', SKIPPED)
            return

        self.set_step('ollama', RUNNING)

        if winget_available():
            self.log(f'  Running: winget install {WINGET_OLLAMA}')
            if winget_install(WINGET_OLLAMA, self.log):
                if shutil.which('ollama'):
                    self.log('  Ollama installed!')
                    self.log('  Next: open a terminal and run:')
                    self.log('    ollama pull qwen2.5:7b')
                    self.set_step('ollama', DONE)
                    return

        self.log(f'  Download Ollama from: {URL_OLLAMA}')
        self.log('  After installing, run:  ollama pull qwen2.5:7b')
        self.set_step('ollama', MANUAL)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    InstallerApp().mainloop()


if __name__ == '__main__':
    main()
