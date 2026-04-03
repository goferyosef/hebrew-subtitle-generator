"""
Subtitle Translator -> Hebrew
Uses Google Gemini API (free tier)
No extra packages needed - only Python built-in libraries.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import re
import json
import urllib.request
import urllib.error
import time

# ── Config ────────────────────────────────────────────────────────────────────

API_KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".gemini_key")
MODEL        = "gemini-2.5-flash"
BATCH_SIZE   = 50
RETRY_WAIT   = 15
MAX_RETRIES  = 4


# ── SRT helpers ───────────────────────────────────────────────────────────────

def parse_srt(content):
    content = content.replace("\r\n", "\n").replace("\r", "\n")
    entries = []
    for block in re.split(r"\n{2,}", content.strip()):
        lines = block.strip().split("\n")
        if len(lines) >= 2 and "-->" in lines[1]:
            entries.append({
                "idx":    lines[0].strip(),
                "timing": lines[1].strip(),
                "text":   "\n".join(lines[2:]) if len(lines) > 2 else "",
            })
    return entries


def build_srt(entries):
    return "\n".join(
        f"{e['idx']}\n{e['timing']}\n{e['text']}\n"
        for e in entries
    )


# ── Gemini API ────────────────────────────────────────────────────────────────

def call_gemini(api_key, prompt):
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{MODEL}:generateContent?key={api_key}"
    )
    payload = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"maxOutputTokens": 8192},
    }).encode("utf-8")

    for attempt in range(MAX_RETRIES):
        req = urllib.request.Request(
            url, data=payload,
            headers={"Content-Type": "application/json"},
        )
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read())
            return data["candidates"][0]["content"]["parts"][0]["text"]
        except urllib.error.HTTPError as e:
            if e.code == 429:
                wait = RETRY_WAIT * (attempt + 1)
                time.sleep(wait)
                continue
            body = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"HTTP {e.code}: {body}") from e

    raise RuntimeError("Rate limit hit after several retries. Please wait a minute and try again.")


def translate_batch(api_key, texts):
    numbered = "\n".join(f"[{i+1}] {t}" for i, t in enumerate(texts))
    prompt = (
        "Translate the following subtitle lines to Hebrew.\n"
        "Rules:\n"
        "- Keep translations natural and suitable for TV subtitles.\n"
        "- Preserve any line breaks inside a subtitle as-is.\n"
        "- Return ONLY the translations, each prefixed with [N] where N is the number.\n"
        "- Do NOT add any explanation or extra text.\n\n"
        + numbered
    )
    result = call_gemini(api_key, prompt)
    out = {}
    for line in result.split("\n"):
        m = re.match(r"\[(\d+)\]\s*(.*)", line.strip())
        if m:
            out[int(m.group(1))] = m.group(2)
    return out


# ── Translation orchestrator ──────────────────────────────────────────────────

def translate_entries(api_key, entries, progress_cb, cancel_flag):
    texts_with_pos = [
        (i, e["text"]) for i, e in enumerate(entries) if e["text"].strip()
    ]
    total = len(texts_with_pos)
    done  = 0

    for start in range(0, total, BATCH_SIZE):
        if cancel_flag[0]:
            break
        batch       = texts_with_pos[start : start + BATCH_SIZE]
        batch_texts = [t for _, t in batch]
        batch_pos   = [p for p, _ in batch]

        translations = translate_batch(api_key, batch_texts)

        for local_idx, pos in enumerate(batch_pos):
            key = local_idx + 1
            if key in translations:
                entries[pos]["text"] = translations[key]

        done += len(batch)
        progress_cb(done, total)

    return entries


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Subtitle Translator - Hebrew")
        self.resizable(False, False)
        self.configure(bg="#0f0f0f")

        self._cancel_flag = [False]
        self._api_key     = tk.StringVar(value=self._load_key())
        self._file_path   = tk.StringVar(value="")
        self._status      = tk.StringVar(value="Ready.")

        self._build_ui()
        self._center()

    # ── UI layout ────────────────────────────────────────────────────────────

    def _build_ui(self):
        PAD  = 20
        BG   = "#0f0f0f"
        CARD = "#1a1a1a"
        ACC  = "#4fc3f7"
        TXT  = "#f0ede6"
        DIM  = "#777"

        # ── header ──
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=PAD, pady=(PAD, 0))
        tk.Label(hdr, text="◈", font=("Georgia", 26),
                 fg=ACC, bg=BG).pack(side="left")
        tk.Label(hdr, text="  Subtitle Translator",
                 font=("Georgia", 19, "bold"),
                 fg=TXT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  \u2192 \u05e2\u05d1\u05e8\u05d9\u05ea",
                 font=("Georgia", 15), fg=ACC, bg=BG).pack(side="left")

        tk.Frame(self, height=1, bg=ACC).pack(fill="x", padx=PAD, pady=12)

        # ── API key ──
        card_key = tk.Frame(self, bg=CARD)
        card_key.pack(fill="x", padx=PAD, pady=(0, 10))
        tk.Label(card_key, text="Google Gemini API Key",
                 font=("Courier New", 9, "bold"),
                 fg=DIM, bg=CARD).pack(anchor="w", padx=12, pady=(10, 2))

        key_row = tk.Frame(card_key, bg=CARD)
        key_row.pack(fill="x", padx=12, pady=(0, 10))
        self._key_entry = tk.Entry(
            key_row, textvariable=self._api_key, show="\u2022", width=46,
            font=("Courier New", 11), bg="#252525", fg=TXT,
            insertbackground=ACC, relief="flat", bd=6,
        )
        self._key_entry.pack(side="left", fill="x", expand=True)
        tk.Button(
            key_row, text="Save", command=self._save_key,
            font=("Courier New", 9, "bold"), bg=ACC, fg="#0f0f0f",
            activebackground="#81d4fa", relief="flat", bd=0,
            padx=10, cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        # ── file chooser ──
        card_file = tk.Frame(self, bg=CARD)
        card_file.pack(fill="x", padx=PAD, pady=(0, 10))
        tk.Label(card_file, text="Subtitle File  (.srt or .txt)",
                 font=("Courier New", 9, "bold"),
                 fg=DIM, bg=CARD).pack(anchor="w", padx=12, pady=(10, 2))

        file_row = tk.Frame(card_file, bg=CARD)
        file_row.pack(fill="x", padx=12, pady=(0, 10))
        tk.Label(
            file_row, textvariable=self._file_path,
            font=("Courier New", 10), fg=TXT, bg="#252525",
            anchor="w", width=40, relief="flat", padx=8, pady=6,
        ).pack(side="left", fill="x", expand=True)
        tk.Button(
            file_row, text="Browse\u2026", command=self._browse,
            font=("Courier New", 9, "bold"), bg="#2a2a2a", fg=TXT,
            activebackground="#333", relief="flat", bd=0,
            padx=12, cursor="hand2",
        ).pack(side="left", padx=(8, 0))

        # ── progress bar ──
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Blue.Horizontal.TProgressbar",
            troughcolor="#1a1a1a", background=ACC,
            bordercolor="#0f0f0f", lightcolor=ACC, darkcolor=ACC,
        )
        self._progress = ttk.Progressbar(
            self, style="Blue.Horizontal.TProgressbar",
            orient="horizontal", mode="determinate", length=460,
        )
        self._progress.pack(fill="x", padx=PAD, pady=(0, 4))

        # ── status label ──
        tk.Label(self, textvariable=self._status,
                 font=("Courier New", 9), fg=DIM, bg=BG,
                 anchor="w").pack(fill="x", padx=PAD)

        # ── action buttons ──
        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(pady=PAD)

        self._translate_btn = tk.Button(
            btn_row, text="\u27e9  Translate to Hebrew",
            command=self._start_translation,
            font=("Georgia", 13, "bold"), bg=ACC, fg="#0f0f0f",
            activebackground="#81d4fa", relief="flat", bd=0,
            padx=28, pady=12, cursor="hand2",
        )
        self._translate_btn.pack(side="left", padx=(0, 10))

        self._cancel_btn = tk.Button(
            btn_row, text="Cancel", command=self._cancel,
            font=("Courier New", 9), state="disabled",
            bg="#2a2a2a", fg=DIM, activebackground="#333",
            relief="flat", bd=0, padx=16, pady=12, cursor="hand2",
        )
        self._cancel_btn.pack(side="left")

        # ── footer ──
        tk.Label(
            self,
            text="Powered by Google Gemini  \u00b7  Saved alongside the original file",
            font=("Courier New", 8), fg="#333", bg=BG,
        ).pack(pady=(0, PAD))

    # ── helpers ──────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w  = self.winfo_width()
        h  = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _load_key(self):
        if os.path.exists(API_KEY_FILE):
            with open(API_KEY_FILE) as f:
                return f.read().strip()
        return ""

    def _save_key(self):
        with open(API_KEY_FILE, "w") as f:
            f.write(self._api_key.get().strip())
        self._status.set("API key saved.")

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Open subtitle file",
            filetypes=[
                ("Subtitle files", "*.srt *.txt"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._file_path.set(path)
            self._status.set(f"Loaded: {os.path.basename(path)}")
            self._progress["value"] = 0

    # ── translation flow ─────────────────────────────────────────────────────

    def _start_translation(self):
        api_key   = self._api_key.get().strip()
        file_path = self._file_path.get().strip()

        if not api_key:
            messagebox.showerror(
                "Missing API Key",
                "Please enter your Gemini API key and click Save."
            )
            return
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("No File", "Please choose a subtitle file first.")
            return

        self._cancel_flag = [False]
        self._translate_btn.config(state="disabled")
        self._cancel_btn.config(state="normal")
        self._progress["value"] = 0
        self._status.set("Reading file...")

        threading.Thread(
            target=self._run, args=(api_key, file_path), daemon=True
        ).start()

    def _run(self, api_key, file_path):
        try:
            with open(file_path, "r", encoding="utf-8-sig") as f:
                content = f.read()

            entries = parse_srt(content)
            if not entries:
                self._gui(self._status.set, "No subtitle entries found.")
                return

            n = sum(1 for e in entries if e["text"].strip())
            self._gui(self._status.set, f"Translating {n} subtitles...")

            def progress_cb(done, total):
                pct = int(done / total * 100)
                self._gui(self._progress.__setitem__, "value", pct)
                self._gui(
                    self._status.set,
                    f"Translating...  {done}/{total}  ({pct}%)"
                )

            entries = translate_entries(
                api_key, entries, progress_cb, self._cancel_flag
            )

            if self._cancel_flag[0]:
                self._gui(self._status.set, "Cancelled.")
                return

            base, _ = os.path.splitext(file_path)
            out_path = base + "_hebrew.srt"
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(build_srt(entries))

            self._gui(self._progress.__setitem__, "value", 100)
            self._gui(
                self._status.set,
                f"Done!  Saved: {os.path.basename(out_path)}"
            )
            self._gui(
                messagebox.showinfo,
                "Translation Complete",
                f"File saved to:\n{out_path}"
            )

        except Exception as exc:
            self._gui(messagebox.showerror, "Error", str(exc))
            self._gui(self._status.set, f"Error: {exc}")
        finally:
            self._gui(self._translate_btn.config, state="normal")
            self._gui(self._cancel_btn.config, state="disabled")

    def _cancel(self):
        self._cancel_flag[0] = True
        self._status.set("Cancelling...")

    @staticmethod
    def _gui(fn, *args, **kwargs):
        try:
            fn(*args, **kwargs)
        except Exception:
            pass


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
