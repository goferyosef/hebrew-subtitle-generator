#!/usr/bin/env python3
"""
Subtitle Generator & Hebrew Translator
  - Subtitle text files  (SRT/ASS/VTT/…)  → translate to Hebrew
  - Video with soft subs                   → extract → translate to Hebrew
  - Video with hard-coded subs             → OCR → translate to Hebrew
  - Auto-sync unsynchronised subtitles     → align to movie audio via ffsubsync

Drag-and-drop a file onto the window, or use the buttons.
Output: ORIGINAL_NAME_HEB.srt saved next to the source file.

Translation backends (auto-selected):
  1. Ollama (local AI, free) — gender-aware, context-aware Hebrew
  2. Google Translate        — free fallback via deep-translator
"""

import difflib
import importlib
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.request
import urllib.error
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

# ─── Tesseract auto-detect (finds it even if not in PATH) ────────────────────
_TESSERACT_KNOWN = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
]
def _configure_tesseract():
    if shutil.which('tesseract'):
        return
    for path in _TESSERACT_KNOWN:
        if Path(path).exists():
            os.environ['PATH'] = str(Path(path).parent) + ';' + os.environ.get('PATH', '')
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = path
            except ImportError:
                pass
            break
_configure_tesseract()


# ─── ffmpeg auto-detect (finds it even if not in PATH) ───────────────────────
def _configure_ffmpeg():
    if shutil.which('ffmpeg'):
        return
    local_app = os.environ.get('LOCALAPPDATA', '')
    candidates = [
        os.path.join(local_app, 'ffmpegio', 'ffmpeg-downloader', 'ffmpeg', 'bin'),
        r'C:\ffmpeg\bin',
        r'C:\Program Files\ffmpeg\bin',
        r'C:\Program Files (x86)\ffmpeg\bin',
    ]
    for bin_dir in candidates:
        if Path(bin_dir, 'ffmpeg.exe').exists():
            os.environ['PATH'] = bin_dir + ';' + os.environ.get('PATH', '')
            break
_configure_ffmpeg()


# ─── Optional drag-and-drop support ──────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _TK_BASE = TkinterDnD.Tk
    _HAS_DND = True
except ImportError:
    _TK_BASE = tk.Tk
    _HAS_DND = False


# ─── Constants ────────────────────────────────────────────────────────────────

SUBTITLE_EXTS = {'.srt', '.ass', '.ssa', '.vtt', '.sub', '.smi'}
VIDEO_EXTS    = {'.mkv', '.mp4', '.avi', '.mov', '.m4v', '.ts', '.wmv', '.flv', '.webm'}

TEXT_CODECS  = frozenset({
    'subrip', 'ass', 'ssa', 'webvtt', 'mov_text',
    'microdvd', 'sami', 'realtext', 'subviewer', 'text', 'dvbsub'
})
IMAGE_CODECS = frozenset({'hdmv_pgs_subtitle', 'dvd_subtitle', 'dvdsub', 'pgssub'})

HEBREW_LANG_TAGS  = {'heb', 'he', 'iw'}
ENGLISH_LANG_TAGS = {'eng', 'en'}

# Google Translate
GOOGLE_BATCH_SEP  = '\n\n@@SEP@@\n\n'
GOOGLE_BATCH_MAX  = 4500
GOOGLE_RATE_DELAY = 0.4

# Ollama
OLLAMA_URL        = "http://localhost:11434"
OLLAMA_BATCH_SIZE = 15
OLLAMA_CONTEXT    = 6
OLLAMA_TIMEOUT    = 45

# Hebrew RTL marker — prepended to each translated line so players display it correctly
RTL_MARK = '\u200f'

# Gemini
GEMINI_API_KEY_FILE = Path(__file__).parent / '.gemini_key'
GEMINI_MODEL        = 'gemini-2.5-flash'
GEMINI_BATCH_SIZE   = 50
GEMINI_RETRY_WAIT   = 15
GEMINI_MAX_RETRIES  = 4
GEMINI_BR           = '<<BR>>'

PREFERRED_MODELS = [
    'qwen2.5:7b', 'qwen2.5:3b', 'qwen2.5:1.5b', 'qwen2.5:latest',
    'qwen2:7b',   'qwen2:latest',
    'llama3.1:8b', 'llama3.2:3b', 'llama3.2:1b', 'llama3:latest',
    'mistral:7b',  'mistral:latest',
    'gemma2:9b',   'gemma2:2b',
    'phi3:medium', 'phi3:mini',
]

HEBREW_SYSTEM_PROMPT = """\
You are a professional Hebrew subtitle translator for films and TV shows.

TRANSLATION RULES — follow these exactly:
- Translate to natural, conversational Israeli Hebrew (עברית מדוברת)
- Hebrew is a gendered language — use the CORRECT gender form for ALL verbs, adjectives, and pronouns:
    Addressing a MALE:    אתה, שלך (m), עשית (m), הלכת (m)
    Addressing a FEMALE:  את,  שלך (f), עשית (f), הלכת (f)
    Speaking ABOUT a male:   הוא, שלו, עשה, הלך, טוב, גדול
    Speaking ABOUT a female: היא, שלה, עשתה, הלכה, טובה, גדולה
- Infer gender of who is SPEAKING and who is SPOKEN TO from the context provided
- Maintain consistent gender for each character — do NOT switch mid-scene
- Keep translations concise — subtitles must be readable in 1–2 seconds
- Preserve register: slang stays slangy, formal stays formal, urgency stays urgent
- Swear words → use natural Hebrew equivalents, do not soften them
- Do NOT transliterate English words that have common Hebrew equivalents
- Return ONLY a valid JSON array of translated strings — no explanation, no markdown

{gender_block}"""


# ─── Utility helpers ──────────────────────────────────────────────────────────

_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0


def run_cmd(*args, timeout=60):
    return subprocess.run(
        list(args), capture_output=True, text=True,
        timeout=timeout, creationflags=_NO_WINDOW,
    )


def detect_encoding(path: str) -> str:
    try:
        import chardet
        with open(path, 'rb') as f:
            raw = f.read(65536)
        r = chardet.detect(raw)
        enc = r.get('encoding') or 'utf-8'
        return enc if r.get('confidence', 0) >= 0.6 else 'utf-8'
    except ImportError:
        return 'utf-8'


def strip_sub_tags(text: str) -> str:
    text = re.sub(r'\{[^}]*\}', '', text)
    text = re.sub(r'<[^>]+>', '', text)
    return text.replace('\\N', ' ').replace('\\n', ' ').strip()


# ─── Stream selection ─────────────────────────────────────────────────────────

def probe_subtitle_streams(video_path: str) -> list:
    try:
        r = run_cmd('ffprobe', '-v', 'quiet', '-print_format', 'json',
                    '-show_streams', '-select_streams', 's', video_path, timeout=30)
        return json.loads(r.stdout or '{}').get('streams', [])
    except Exception:
        return []


def select_best_stream(streams: list) -> tuple:
    def lang(s):  return s.get('tags', {}).get('language', '').lower()
    def codec(s): return s.get('codec_name', '').lower()
    text_streams  = [(i, s) for i, s in enumerate(streams) if codec(s) in TEXT_CODECS]
    image_streams = [(i, s) for i, s in enumerate(streams) if codec(s) in IMAGE_CODECS]
    for i, s in text_streams:
        lg = lang(s)
        if lg and lg not in HEBREW_LANG_TAGS and lg not in ENGLISH_LANG_TAGS:
            return i, False
    for i, s in text_streams:
        if lang(s) in ENGLISH_LANG_TAGS:
            return i, False
    if text_streams:  return text_streams[0][0],  False
    if image_streams: return image_streams[0][0], True
    return None, False


# ─── Soft subtitle extraction ─────────────────────────────────────────────────

def extract_soft_subtitles(video_path: str, stream_index: int, out_path: str):
    for codec in ('srt', 'subrip'):
        run_cmd('ffmpeg', '-y', '-i', video_path,
                '-map', f'0:s:{stream_index}', '-c:s', codec, out_path, timeout=120)
        if Path(out_path).exists() and Path(out_path).stat().st_size > 0:
            return
    raise RuntimeError("ffmpeg failed to extract subtitle stream.")


# ─── Auto-sync (ffsubsync) ────────────────────────────────────────────────────

def sync_subtitles(video_path: str, srt_path: str, output_path: str, log_cb) -> bool:
    """
    Align subtitle timing to the movie's audio using ffsubsync.
    Returns True on success.  Tries CLI first, then Python API.
    """
    # Try CLI (installed as 'ffs' by pip install ffsubsync)
    for cli in ('ffs', 'ffsubsync'):
        try:
            log_cb(f"  Running {cli}…", 'dim')
            r = run_cmd(cli, video_path, '-i', srt_path, '-o', output_path, timeout=600)
            if r.returncode == 0 and Path(output_path).exists() and Path(output_path).stat().st_size > 0:
                combined = r.stdout + r.stderr
                m = re.search(r'offset.*?([+-]?\d+\.?\d*)\s*s', combined, re.IGNORECASE)
                if m:
                    log_cb(f"  Detected offset: {m.group(1)} seconds")
                return True
            if r.returncode != 0:
                log_cb(f"  {cli} error: {(r.stderr or r.stdout)[-300:]}", 'warning')
        except FileNotFoundError:
            continue
        except subprocess.TimeoutExpired:
            log_cb("  ffsubsync timed out (>10 min). Try a shorter video clip.", 'error')
            return False

    # Try Python API directly
    try:
        from ffsubsync.ffsubsync import make_parser, run as ffs_run   # type: ignore
        parser = make_parser()
        args   = parser.parse_args([video_path, '-i', srt_path, '-o', output_path])
        ffs_run(args)
        return Path(output_path).exists() and Path(output_path).stat().st_size > 0
    except ImportError:
        log_cb("ffsubsync not installed.  Run:  pip install ffsubsync", 'error')
        return False
    except Exception as e:
        log_cb(f"ffsubsync API error: {e}", 'error')
        return False


# ─── OCR pipeline ─────────────────────────────────────────────────────────────

@dataclass
class OcrLine:
    text: str
    first_ms: int
    last_ms: int


def preprocess_for_ocr(region):
    import cv2
    import numpy as np
    gray   = cv2.cvtColor(region, cv2.COLOR_BGR2GRAY)
    scaled = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    _, t1  = cv2.threshold(scaled, 190, 255, cv2.THRESH_BINARY)
    _, t2  = cv2.threshold(scaled, 100, 255, cv2.THRESH_BINARY_INV)
    t3     = cv2.adaptiveThreshold(scaled, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY, 15, 3)
    hsv    = cv2.cvtColor(cv2.resize(region, None, fx=2, fy=2), cv2.COLOR_BGR2HSV)
    yellow = cv2.inRange(hsv, np.array([15, 80, 80]), np.array([40, 255, 255]))
    return [t1, t2, t3, yellow]


def ocr_frame(frame) -> str:
    import pytesseract
    h, w   = frame.shape[:2]
    region = frame[int(h * 0.78):int(h * 0.97), int(w * 0.03):int(w * 0.97)]
    if region.size == 0:
        return ''
    best, cfg = '', r'--oem 1 --psm 6 -c tessedit_char_blacklist=|~^_'
    for img in preprocess_for_ocr(region):
        try:
            raw     = pytesseract.image_to_string(img, config=cfg, lang='eng')
            cleaned = ' '.join(raw.split())
            if len(cleaned) < 3:
                continue
            sym = sum(1 for c in cleaned if not (c.isalpha() or c in " ,.!?'-")) / max(len(cleaned), 1)
            if sym > 0.4:
                continue
            if len(cleaned) > len(best):
                best = cleaned
        except Exception:
            continue
    return best


def deduplicate_ocr_lines(raw: list) -> list:
    if not raw:
        return []
    lines, cur_text, cur_start, cur_last = [], raw[0][1], raw[0][0], raw[0][0]
    for ms, text in raw[1:]:
        similar = difflib.SequenceMatcher(None, cur_text, text).ratio() > 0.80
        if similar and ms - cur_last <= 2000:
            cur_last = ms
        else:
            if cur_last - cur_start >= 400 and len(cur_text.strip()) >= 3:
                lines.append(OcrLine(cur_text, cur_start, cur_last + 800))
            cur_text, cur_start, cur_last = text, ms, ms
    if cur_last - cur_start >= 400 and len(cur_text.strip()) >= 3:
        lines.append(OcrLine(cur_text, cur_start, cur_last + 800))
    return lines


def build_srt_from_ocr(lines: list):
    import pysubs2
    subs = pysubs2.SSAFile()
    subs.events = []
    for line in lines:
        subs.events.append(pysubs2.SSAEvent(
            start=line.first_ms,
            end=min(line.last_ms, line.first_ms + 8000),
            text=line.text,
        ))
    return subs


# ─── Ollama AI Translation ────────────────────────────────────────────────────

def check_ollama() -> tuple:
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=3) as resp:
            data   = json.loads(resp.read())
            models = [m['name'] for m in data.get('models', [])]
            return True, models
    except Exception:
        return False, []


def pick_best_model(models: list) -> str:
    available  = set(models)
    by_prefix  = {}
    for m in models:
        by_prefix.setdefault(m.split(':')[0], m)
    for pref in PREFERRED_MODELS:
        if pref in available:
            return pref
        pfx = pref.split(':')[0]
        if pfx in by_prefix:
            return by_prefix[pfx]
    return models[0] if models else ''


def ollama_chat(model: str, system: str, user: str, timeout: int = OLLAMA_TIMEOUT) -> str:
    payload = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        "stream": False,
    }).encode()
    req = urllib.request.Request(
        f"{OLLAMA_URL}/api/chat", data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["message"]["content"]


def _parse_gender_response(response: str, log_cb) -> str:
    """Parse a JSON gender-detection response and return the gender block string."""
    m = re.search(r'\{[\s\S]*\}', response)
    if not m:
        return ""
    data  = json.loads(m.group())
    chars = [c for c in data.get('characters', [])
             if c.get('name') and c.get('gender') != 'unknown']
    if not chars:
        return ""
    summary   = ', '.join(f"{c['name']} ({c['gender']})" for c in chars)
    log_cb(f"  Characters: {summary}", 'dim')
    char_list = '\n'.join(f"  - {c['name']}: {c['gender']}" for c in chars)
    return (
        "CHARACTER GENDERS (use correct Hebrew gender forms when these characters "
        "speak or are spoken to):\n" + char_list
    )


def detect_character_genders(sample_texts: list, model: str, log_cb,
                              gemini_key: str = None) -> str:
    if not sample_texts:
        return ""
    numbered = "\n".join(f"{i+1}. {t}" for i, t in enumerate(sample_texts[:60]))
    prompt = (
        "Read these subtitle lines and identify character names and their genders.\n"
        "Return ONLY JSON: {\"characters\": [{\"name\": \"...\", \"gender\": \"male|female|unknown\"}, ...]}\n"
        "If no character names are detectable, return {\"characters\": []}\n\n"
        f"Lines:\n{numbered}"
    )

    # Try Gemini first (fast, reliable)
    if gemini_key:
        try:
            response = _call_gemini(gemini_key, prompt)
            result   = _parse_gender_response(response, log_cb)
            if result:
                return result
        except Exception as e:
            log_cb(f"  Gemini gender detection failed: {e} — trying Ollama", 'dim')

    # Fall back to Ollama
    try:
        response = ollama_chat(
            model,
            "You are a script analyst. Identify character names and genders from dialogue. Return only valid JSON.",
            prompt, timeout=40,
        )
        return _parse_gender_response(response, log_cb)
    except Exception as e:
        log_cb(f"  Gender detection skipped: {e}", 'dim')
    return ""


def _parse_ollama_json(response: str, n: int) -> list:
    m = re.search(r'\[[\s\S]*\]', response)
    if m:
        parsed = json.loads(m.group())
        if isinstance(parsed, list) and len(parsed) >= n:
            return [str(t).strip() for t in parsed[:n]]
    items = re.findall(r'^\s*\d+\.\s*["\']?(.+?)["\']?\s*$', response, re.MULTILINE)
    if len(items) >= n:
        return [t.strip() for t in items[:n]]
    raise ValueError(f"Could not parse {n} items from Ollama response")


def _ollama_translate_batch(texts: list, model: str, system: str, context_window: list) -> list:
    ctx = ""
    if context_window:
        ctx = "RECENT CONTEXT (already translated — reference only, do NOT retranslate):\n"
        ctx += "\n".join(f"  {t}" for t in context_window) + "\n\n"
    numbered = "\n".join(f'{i+1}. "{t}"' for i, t in enumerate(texts))
    user_msg  = (
        f"{ctx}Translate these {len(texts)} subtitle lines to Hebrew.\n"
        f"Return ONLY a JSON array of exactly {len(texts)} Hebrew strings:\n\n{numbered}"
    )
    return _parse_ollama_json(ollama_chat(model, system, user_msg), len(texts))


def _ollama_full_translate(raw_texts: list, model: str, log_cb,
                           cancel_check=None, gemini_key: str = None,
                           progress_cb=None, failed_counter=None) -> list:
    clean_texts = [strip_sub_tags(t) for t in raw_texts]
    log_cb("  Detecting character genders…", 'dim')
    gender_block = detect_character_genders([t for t in clean_texts if t][:60], model, log_cb,
                                            gemini_key=gemini_key)
    system       = HEBREW_SYSTEM_PROMPT.format(gender_block=gender_block)

    results        = list(raw_texts)
    context_window = []
    total          = len(raw_texts)

    for batch_start in range(0, total, OLLAMA_BATCH_SIZE):
        if cancel_check and cancel_check():
            log_cb("  Cancelled.", 'warning')
            return results

        batch_raw   = raw_texts[batch_start : batch_start + OLLAMA_BATCH_SIZE]
        batch_clean = clean_texts[batch_start : batch_start + OLLAMA_BATCH_SIZE]
        non_empty   = [(j, t) for j, t in enumerate(batch_clean) if t.strip()]

        if not non_empty:
            continue

        try:
            translated = _ollama_translate_batch([t for _, t in non_empty], model, system, context_window)
            for (j, _), heb in zip(non_empty, translated):
                results[batch_start + j] = RTL_MARK + heb
                context_window.append(heb)
            context_window = context_window[-OLLAMA_CONTEXT:]
        except Exception as e:
            batch_num = batch_start // OLLAMA_BATCH_SIZE + 1
            if failed_counter is not None:
                failed_counter[0] += 1
            if gemini_key:
                log_cb(f"  Batch {batch_num} failed ({e}) — Gemini fallback", 'warning')
                batch_texts = [raw_texts[batch_start + j] for j, _ in non_empty]
                fallback    = _gemini_full_translate(batch_texts, gemini_key, log_cb, cancel_check)
                for (j, _), heb in zip(non_empty, fallback):
                    results[batch_start + j] = heb
            else:
                log_cb(f"  Batch {batch_num} failed ({e}) — Google fallback", 'warning')
                try:
                    from deep_translator import GoogleTranslator
                    gt = GoogleTranslator(source='auto', target='iw')
                    for j, text in enumerate(batch_clean):
                        if text.strip():
                            try:
                                results[batch_start + j] = RTL_MARK + gt.translate(text)
                                time.sleep(0.2)
                            except Exception:
                                pass
                except ImportError:
                    pass

        done = min(batch_start + OLLAMA_BATCH_SIZE, total)
        if done % 60 == 0 or done == total:
            log_cb(f"  {done}/{total} lines translated", 'dim')
        if progress_cb:
            progress_cb(done, total)

    return results


# ─── Gemini AI Translation ────────────────────────────────────────────────────

def _load_gemini_key() -> str:
    if GEMINI_API_KEY_FILE.exists():
        return GEMINI_API_KEY_FILE.read_text().strip()
    return ''


def _call_gemini(api_key: str, prompt: str) -> str:
    url = (
        'https://generativelanguage.googleapis.com/v1beta/models/'
        f'{GEMINI_MODEL}:generateContent?key={api_key}'
    )
    payload = json.dumps({
        'contents': [{'parts': [{'text': prompt}]}],
        'generationConfig': {'maxOutputTokens': 8192},
    }).encode('utf-8')
    for attempt in range(GEMINI_MAX_RETRIES):
        req = urllib.request.Request(
            url, data=payload,
            headers={'Content-Type': 'application/json'},
        )
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.loads(resp.read())
            return data['candidates'][0]['content']['parts'][0]['text']
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(GEMINI_RETRY_WAIT * (attempt + 1))
                continue
            raise RuntimeError(
                f'Gemini HTTP {e.code}: {e.read().decode("utf-8", errors="replace")}'
            ) from e
    raise RuntimeError('Gemini rate limit — try again in a minute.')


def _gemini_full_translate(raw_texts: list, api_key: str, log_cb,
                           cancel_check=None, progress_cb=None,
                           failed_counter=None) -> list:
    results     = list(raw_texts)
    clean_texts = [strip_sub_tags(t) for t in raw_texts]
    total       = len(raw_texts)

    for start in range(0, total, GEMINI_BATCH_SIZE):
        if cancel_check and cancel_check():
            log_cb('  Cancelled.', 'warning')
            return results

        batch_clean   = clean_texts[start : start + GEMINI_BATCH_SIZE]
        batch_indices = list(range(start, min(start + GEMINI_BATCH_SIZE, total)))
        non_empty     = [(pos, t) for pos, t in zip(batch_indices, batch_clean) if t.strip()]
        if not non_empty:
            continue

        flat     = [t.replace('\n', GEMINI_BR) for _, t in non_empty]
        numbered = '\n'.join(f'[{i+1}] {t}' for i, t in enumerate(flat))
        prompt   = (
            'Translate the following subtitle lines to Hebrew.\n'
            'Rules:\n'
            '- Keep translations natural and suitable for TV subtitles.\n'
            f'- The token {GEMINI_BR} marks a line break inside a subtitle — preserve it.\n'
            '- Return ONLY the translations, each prefixed with [N] where N is the number.\n'
            '- Every translation on exactly ONE line. No extra text.\n\n'
            + numbered
        )

        try:
            result = _call_gemini(api_key, prompt)
            out    = {}
            for line in result.split('\n'):
                m = re.match(r'\[(\d+)\]\s*(.*)', line.strip())
                if m:
                    out[int(m.group(1))] = m.group(2).strip().replace(GEMINI_BR, '\n')
            for local_idx, (pos, _) in enumerate(non_empty):
                key = local_idx + 1
                if key in out:
                    results[pos] = RTL_MARK + out[key]
        except Exception as e:
            log_cb(f'  Gemini batch {start // GEMINI_BATCH_SIZE + 1} failed: {e}', 'error')
            if failed_counter is not None:
                failed_counter[0] += 1

        done = min(start + GEMINI_BATCH_SIZE, total)
        log_cb(f'  {done}/{total} lines translated', 'dim')
        if progress_cb:
            progress_cb(done, total)

    return results


# ─── Google Translate (fallback) ──────────────────────────────────────────────

def _google_batch_translate(texts: list, log_cb, cancel_check=None, progress_cb=None,
                            failed_counter=None) -> list:
    from deep_translator import GoogleTranslator
    translator = GoogleTranslator(source='auto', target='iw')
    results    = list(texts)
    total      = len(texts)
    batch_t, batch_i, batch_chars = [], [], 0

    def flush(b_texts, b_indices, attempt=0):
        if not b_texts:
            return
        try:
            joined     = GOOGLE_BATCH_SEP.join(b_texts)
            translated = translator.translate(joined)
            parts      = re.split(r'@@\s*SEP\s*@@', translated)
            if len(parts) == len(b_texts):
                for k, idx in enumerate(b_indices):
                    results[idx] = RTL_MARK + parts[k].strip()
                return
        except Exception as e:
            if attempt < 2:
                time.sleep(2 ** attempt)
                flush(b_texts, b_indices, attempt + 1)
                return
            log_cb("  Batch failed — retrying line-by-line…", 'dim')
            if failed_counter is not None:
                failed_counter[0] += 1
        for txt, idx in zip(b_texts, b_indices):
            for retry in range(3):
                try:
                    results[idx] = RTL_MARK + translator.translate(txt)
                    time.sleep(0.25)
                    break
                except Exception:
                    if retry == 2:
                        pass
                    time.sleep(1.0)

    for i, text in enumerate(texts):
        if cancel_check and cancel_check():
            log_cb("  Cancelled.", 'warning')
            break
        clean = strip_sub_tags(text)
        if not clean:
            continue
        if batch_chars + len(clean) + len(GOOGLE_BATCH_SEP) > GOOGLE_BATCH_MAX and batch_t:
            flush(batch_t, batch_i)
            time.sleep(GOOGLE_RATE_DELAY)
            batch_t, batch_i, batch_chars = [], [], 0
            if i % 100 == 0:
                log_cb(f"  {i}/{total} lines translated", 'dim')
            if progress_cb:
                progress_cb(i, total)
        batch_t.append(clean)
        batch_i.append(i)
        batch_chars += len(clean) + len(GOOGLE_BATCH_SEP)

    if batch_t:
        flush(batch_t, batch_i)
    if progress_cb:
        progress_cb(total, total)
    return results


# ─── Translation dispatcher ───────────────────────────────────────────────────

def translate_and_save(subs, out_path: str, log_cb,
                       ollama_model: str = None, gemini_key: str = None,
                       cancel_check=None, progress_cb=None):
    import pysubs2

    with_text = [(i, e) for i, e in enumerate(subs.events) if strip_sub_tags(e.text)]
    if not with_text:
        log_cb("No subtitle text to translate.", 'warning')
        subs.save(out_path, encoding='utf-8-sig')
        return

    raw_texts      = [e.text for _, e in with_text]
    failed_counter = [0]
    start_time     = time.time()

    if ollama_model:
        translator_name = f"Ollama ({ollama_model})"
        log_cb(f"Translating {len(raw_texts)} lines…")
        log_cb(f"  {translator_name} — AI, gender-aware")
        translated = _ollama_full_translate(raw_texts, ollama_model, log_cb, cancel_check,
                                            gemini_key=gemini_key, progress_cb=progress_cb,
                                            failed_counter=failed_counter)
    elif gemini_key:
        translator_name = f"Gemini ({GEMINI_MODEL})"
        log_cb(f"Translating {len(raw_texts)} lines…")
        log_cb(f"  {translator_name}")
        translated = _gemini_full_translate(raw_texts, gemini_key, log_cb, cancel_check,
                                            progress_cb=progress_cb,
                                            failed_counter=failed_counter)
    else:
        translator_name = "Google Translate"
        log_cb(f"Translating {len(raw_texts)} lines…")
        log_cb(f"  {translator_name} (free)")
        try:
            translated = _google_batch_translate(raw_texts, log_cb, cancel_check,
                                                 progress_cb=progress_cb,
                                                 failed_counter=failed_counter)
        except ImportError:
            log_cb("deep-translator not installed — pip install deep-translator", 'error')
            return

    t_map  = {i: t for (i, _), t in zip(with_text, translated)}
    result = pysubs2.SSAFile()
    result.events = []
    for i, event in enumerate(subs.events):
        result.events.append(pysubs2.SSAEvent(
            start=event.start, end=event.end,
            text=t_map.get(i, event.text),
        ))
    result.save(out_path, encoding='utf-8-sig')

    elapsed  = time.time() - start_time
    duration = f"{int(elapsed // 60)}m {int(elapsed % 60)}s" if elapsed >= 60 else f"{int(elapsed)}s"

    log_cb("─" * 50)
    log_cb(f"  Translated {len(raw_texts)} lines → Hebrew")
    log_cb(f"  Translator:  {translator_name}")
    log_cb(f"  Duration:    {duration}")
    if failed_counter[0]:
        log_cb(f"  Failed batches: {failed_counter[0]} (recovered via fallback)", 'warning')
    else:
        log_cb(f"  Failed batches: none", 'success')
    log_cb(f"  Output: {Path(out_path).name}", 'success')
    log_cb("─" * 50)


# ─── Main GUI Application ─────────────────────────────────────────────────────

class SubtitleApp(_TK_BASE):

    def __init__(self):
        super().__init__()
        self.title("Hebrew Subtitle Generator")
        self.geometry("780x540")
        self.resizable(True, True)
        self.ollama_model     = None
        self.available_models = []
        self.gemini_key       = _load_gemini_key()
        self._cancel_event    = threading.Event()
        self._build_widgets()
        self.after(200, self._check_dependencies)

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # ── Top bar ──
        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.grid(row=0, column=0, sticky='ew')
        top.columnconfigure(3, weight=1)

        ttk.Label(top, text="Hebrew Subtitle Generator",
                  font=('Segoe UI', 13, 'bold')).grid(
            row=0, column=0, columnspan=6, sticky='w', pady=(0, 8)
        )

        # Buttons
        self.open_btn = ttk.Button(top, text="Open File…",
                                   command=self._open_file, width=13)
        self.open_btn.grid(row=1, column=0, padx=(0, 6))

        self.sync_btn = ttk.Button(top, text="Sync Subtitles…",
                                   command=self._sync_workflow, width=16)
        self.sync_btn.grid(row=1, column=1, padx=(0, 6))

        self.cancel_btn = ttk.Button(top, text="Cancel",
                                     command=self._cancel_job,
                                     state=tk.DISABLED, width=8)
        self.cancel_btn.grid(row=1, column=2, padx=(0, 12))

        # File label (stretchy)
        self.file_var = tk.StringVar(value="No file selected")
        ttk.Label(top, textvariable=self.file_var,
                  foreground='#666').grid(row=1, column=3, sticky='w', padx=(0, 16))

        # Translator selector
        ttk.Label(top, text="Translator:").grid(row=1, column=4, sticky='e', padx=(0, 4))
        self.translator_var = tk.StringVar(value="Google Translate (free)")
        self.trans_combo = ttk.Combobox(
            top, textvariable=self.translator_var,
            values=["Google Translate (free)"], state="readonly", width=32,
        )
        self.trans_combo.grid(row=1, column=5, sticky='w')

        # DnD hint label
        if _HAS_DND:
            ttk.Label(top, text="(or drag & drop a file below)",
                      foreground='#888', font=('Segoe UI', 8)).grid(
                row=2, column=0, columnspan=6, sticky='w', pady=(4, 0)
            )

        # ── Log area ──
        mid = ttk.LabelFrame(self, text="Log", padding=(8, 4))
        mid.grid(row=1, column=0, sticky='nsew', padx=12, pady=4)
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            mid, state=tk.DISABLED, wrap=tk.WORD,
            font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
            insertbackground='white', relief='flat',
        )
        sb = ttk.Scrollbar(mid, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.grid(row=0, column=0, sticky='nsew')
        sb.grid(row=0, column=1, sticky='ns')

        for tag, color in [('info', '#d4d4d4'), ('success', '#6fc56f'),
                            ('warning', '#e8c060'), ('error', '#f87171'),
                            ('dim', '#777777')]:
            self.log_text.tag_configure(tag, foreground=color)

        # Drag-and-drop target
        if _HAS_DND:
            self.log_text.drop_target_register(DND_FILES)
            self.log_text.dnd_bind('<<Drop>>', self._on_drop)

        # ── Bottom bar ──
        bot = ttk.Frame(self, padding=(12, 4, 12, 8))
        bot.grid(row=2, column=0, sticky='ew')
        bot.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(bot, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky='ew', pady=(0, 4))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(bot, textvariable=self.status_var,
                  foreground='#555').grid(row=1, column=0, sticky='w')

    # ── Drag and drop ──────────────────────────────────────────────────────────

    def _on_drop(self, event):
        # tkinterdnd2 wraps paths in {} on Windows when there are spaces
        raw  = event.data.strip()
        path = re.sub(r'^\{|\}$', '', raw)
        if Path(path).exists():
            self._dispatch_file(path)
        else:
            self.log(f"Dropped path not found: {path}", 'error')

    # ── Logging ────────────────────────────────────────────────────────────────

    def log(self, message: str, level: str = 'info'):
        ts   = datetime.now().strftime('%H:%M:%S')
        line = f"[{ts}] {message}\n"
        self.after(0, self._write_log, line, level)

    def _write_log(self, line: str, level: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, line, level)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def set_status(self, text: str, running: bool = False):
        def _do():
            self.status_var.set(text)
            if running:
                self.progress.configure(mode='indeterminate')
                self.progress.start(12)
            else:
                self.progress.stop()
                self.progress.configure(mode='indeterminate')
                self.progress['value'] = 0
        self.after(0, _do)

    def _set_progress(self, done: int, total: int):
        def _do():
            self.progress.stop()
            self.progress.configure(mode='determinate', maximum=total, value=done)
            self.status_var.set(f"Translating… {done}/{total} lines ({int(done/total*100)}%)")
        self.after(0, _do)

    # ── Dependency & Ollama detection ─────────────────────────────────────────

    def _check_dependencies(self):
        self.log("Checking dependencies…", 'dim')

        for binary, label in [('ffmpeg', 'ffmpeg'), ('ffprobe', 'ffprobe'),
                               ('tesseract', 'Tesseract OCR')]:
            try:
                run_cmd(binary, '-version', timeout=5)
                self.log(f"  ✓ {label}", 'success')
            except FileNotFoundError:
                self.log(f"  ✗ {label} not found — install for video support", 'warning')

        for pkg, imp in [
            ('pysubs2',         'pysubs2'),
            ('deep-translator', 'deep_translator'),
            ('opencv-python',   'cv2'),
            ('pytesseract',     'pytesseract'),
            ('chardet',         'chardet'),
            ('ffsubsync',       'ffsubsync'),
            ('tkinterdnd2',     'tkinterdnd2'),
        ]:
            try:
                importlib.import_module(imp)
                self.log(f"  ✓ {pkg}", 'success')
            except ImportError:
                level = 'warning' if pkg in ('ffsubsync', 'tkinterdnd2') else 'error'
                self.log(f"  ✗ {pkg}  →  pip install {pkg}", level)

        # Gemini
        if self.gemini_key:
            self.log(f"  Gemini API key found ({GEMINI_MODEL})", 'success')
        else:
            self.log("  No Gemini key — add one to .gemini_key file", 'dim')

        # Ollama
        self.log("Checking Ollama…", 'dim')
        available, models = check_ollama()
        if available and models:
            self.available_models = models
            best                  = pick_best_model(models)
            self.ollama_model     = best
            self.log(f"  Ollama running — best model: {best}", 'success')
        else:
            self.ollama_model = None
            fallback = f"Gemini ({GEMINI_MODEL})" if self.gemini_key else "Google Translate"
            self.log(f"  Ollama not running → using {fallback}", 'warning')

        self._update_translator_options()
        self.log("Ready.", 'success')

    def _update_translator_options(self):
        def _do():
            options = []
            # Gemini first — fast and reliable
            if self.gemini_key:
                options.append(f"Gemini ({GEMINI_MODEL})")
            if self.available_models:
                best = self.ollama_model or ''
                options.append(f"Ollama — {best} (AI, gender-aware)")
                for m in self.available_models:
                    if m != best:
                        options.append(f"Ollama — {m}")
            options.append("Google Translate (free)")
            self.trans_combo['values'] = options
            self.trans_combo.current(0)
        self.after(0, _do)

    def _get_ollama_model(self) -> str:
        sel = self.translator_var.get()
        if sel.startswith("Ollama"):
            m = re.match(r'Ollama\s*—\s*(\S+)', sel)
            return m.group(1) if m else (self.ollama_model or '')
        return ''

    def _get_gemini_key(self) -> str:
        # Always return the key — used as primary when Gemini is selected,
        # and as fallback when Ollama is selected but fails.
        return self.gemini_key

    # ── File dispatch ──────────────────────────────────────────────────────────

    def _open_file(self):
        path = filedialog.askopenfilename(
            title="Select subtitle or video file",
            filetypes=[
                ("All supported", " ".join(f"*{e}" for e in sorted(SUBTITLE_EXTS | VIDEO_EXTS))),
                ("Subtitle files", " ".join(f"*{e}" for e in sorted(SUBTITLE_EXTS))),
                ("Video files",    " ".join(f"*{e}" for e in sorted(VIDEO_EXTS))),
                ("All files", "*.*"),
            ]
        )
        if path:
            self._dispatch_file(path)

    def _dispatch_file(self, path: str):
        self.file_var.set(Path(path).name)
        ext = Path(path).suffix.lower()
        if ext in SUBTITLE_EXTS:
            self._run_in_thread(self._process_subtitle_file, path)
        elif ext in VIDEO_EXTS:
            self._run_in_thread(self._process_video_file, path)
        else:
            messagebox.showerror("Unsupported", f"Unknown file type: {ext}")

    # ── Sync workflow ──────────────────────────────────────────────────────────

    def _sync_workflow(self):
        """
        Sync a subtitle file to a video file using ffsubsync,
        then optionally translate the result to Hebrew.
        """
        srt_path = filedialog.askopenfilename(
            title="Step 1 — Select the subtitle file to sync",
            filetypes=[
                ("Subtitle files", " ".join(f"*{e}" for e in sorted(SUBTITLE_EXTS))),
                ("All files", "*.*"),
            ]
        )
        if not srt_path:
            return

        video_path = filedialog.askopenfilename(
            title="Step 2 — Select the video to sync against",
            filetypes=[
                ("Video files", " ".join(f"*{e}" for e in sorted(VIDEO_EXTS))),
                ("All files", "*.*"),
            ]
        )
        if not video_path:
            return

        self.file_var.set(f"{Path(srt_path).name} ↔ {Path(video_path).name}")
        self._run_in_thread(self._do_sync, srt_path, video_path)

    def _do_sync(self, srt_path: str, video_path: str):
        p          = Path(srt_path)
        sync_path  = str(p.parent / (p.stem + "_SYNC.srt"))

        self.log(f"Syncing: {p.name}")
        self.log(f"  Reference video: {Path(video_path).name}")
        self.log("  Analysing audio… (may take 1–3 minutes)", 'warning')

        ok = sync_subtitles(video_path, srt_path, sync_path, self.log)

        if not ok:
            self.log("Sync failed — see log above.", 'error')
            return

        self.log(f"Sync complete → {Path(sync_path).name}", 'success')

        # Ask whether to also translate
        translate = messagebox.askyesno(
            "Translate?",
            f"Synced subtitles saved as:\n{Path(sync_path).name}\n\n"
            "Also translate to Hebrew now?"
        )
        if translate:
            import pysubs2
            subs = pysubs2.load(sync_path)
            out  = str(p.parent / (p.stem + "_SYNC_HEB.srt"))
            translate_and_save(subs, out, self.log, self._get_ollama_model(),
                               gemini_key=self._get_gemini_key(),
                               cancel_check=self._should_cancel)
            self.log(f"Done! → {Path(out).name}", 'success')

    # ── Cancel ─────────────────────────────────────────────────────────────────

    def _cancel_job(self):
        self._cancel_event.set()
        self.log("Cancelling after current step…", 'warning')
        self.after(0, lambda: self.cancel_btn.configure(state=tk.DISABLED))

    def _should_cancel(self) -> bool:
        return self._cancel_event.is_set()

    # ── Threading ──────────────────────────────────────────────────────────────

    def _run_in_thread(self, func, *args):
        self._cancel_event.clear()
        self.open_btn.configure(state=tk.DISABLED)
        self.sync_btn.configure(state=tk.DISABLED)
        self.cancel_btn.configure(state=tk.NORMAL)
        self.trans_combo.configure(state=tk.DISABLED)
        self.set_status("Processing…", running=True)
        threading.Thread(target=self._wrap, args=(func, *args), daemon=True).start()

    def _wrap(self, func, *args):
        try:
            func(*args)
        except Exception as exc:
            import traceback
            self.log(f"Unexpected error: {exc}", 'error')
            self.log(traceback.format_exc(), 'error')
        finally:
            self.after(0, self._restore_ui)
            self.set_status("Ready", running=False)

    def _restore_ui(self):
        self.open_btn.configure(state=tk.NORMAL)
        self.sync_btn.configure(state=tk.NORMAL)
        self.cancel_btn.configure(state=tk.DISABLED)
        self.trans_combo.configure(state='readonly')

    # ── Pipeline 1: subtitle text file ────────────────────────────────────────

    def _process_subtitle_file(self, path: str):
        import pysubs2
        p = Path(path)
        self.log(f"Loading: {p.name}")
        enc = detect_encoding(path)
        self.log(f"Encoding: {enc}", 'dim')
        try:
            subs = pysubs2.load(path, encoding=enc)
        except Exception:
            subs = pysubs2.load(path, encoding='utf-8', errors='replace')
        self.log(f"Loaded {len(subs.events)} entries")
        if self._should_cancel():
            return
        out = str(p.parent / (p.stem + "_HEB.srt"))
        translate_and_save(subs, out, self.log, self._get_ollama_model(),
                           gemini_key=self._get_gemini_key(),
                           cancel_check=self._should_cancel,
                           progress_cb=self._set_progress)
        if not self._should_cancel():
            self.log(f"Done! → {out}", 'success')

    # ── Pipeline 2 & 3: video file ────────────────────────────────────────────

    def _process_video_file(self, path: str):
        p = Path(path)
        self.log(f"Probing: {p.name}")
        streams = probe_subtitle_streams(path)

        if not streams:
            self.log("No subtitle streams — switching to OCR mode")
            self._process_video_ocr(path)
            return

        self.log(f"Found {len(streams)} subtitle stream(s)")
        idx, is_image = select_best_stream(streams)

        if idx is None or is_image:
            reason = "image-based (PGS/DVD)" if is_image else "none usable"
            self.log(f"Stream type: {reason} — switching to OCR mode")
            self._process_video_ocr(path)
            return

        lang  = streams[idx].get('tags', {}).get('language', 'unknown')
        codec = streams[idx].get('codec_name', 'unknown')
        self.log(f"Extracting stream {idx}  (lang={lang}, codec={codec})")

        srt_path = str(p.parent / (p.stem + ".srt"))
        extract_soft_subtitles(path, idx, srt_path)
        self.log(f"Extracted → {p.stem}.srt", 'success')

        if self._should_cancel():
            return

        import pysubs2
        try:
            subs = pysubs2.load(srt_path)
        except Exception as e:
            self.log(f"Could not parse extracted SRT: {e}", 'error')
            return

        self.log(f"Loaded {len(subs.events)} entries")
        out = str(p.parent / (p.stem + "_HEB.srt"))
        translate_and_save(subs, out, self.log, self._get_ollama_model(),
                           gemini_key=self._get_gemini_key(),
                           cancel_check=self._should_cancel,
                           progress_cb=self._set_progress)
        if not self._should_cancel():
            Path(srt_path).unlink(missing_ok=True)
            self.log(f"Done! → {Path(out).name}", 'success')

    # ── Pipeline 3: hard-coded OCR ────────────────────────────────────────────

    def _process_video_ocr(self, path: str):
        try:
            import cv2
            import pytesseract
        except ImportError:
            self.log("OCR requires opencv-python + pytesseract", 'error')
            self.log("  pip install opencv-python pytesseract", 'error')
            self.log("  Tesseract: https://github.com/UB-Mannheim/tesseract/wiki", 'error')
            return

        p   = Path(path)
        cap = cv2.VideoCapture(path)
        if not cap.isOpened():
            self.log(f"Cannot open video: {path}", 'error')
            return

        fps         = cap.get(cv2.CAP_PROP_FPS) or 25.0
        total_f     = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        duration_ms = int((total_f / fps) * 1000)
        dur_str     = f"{duration_ms // 60000}m {(duration_ms % 60000) // 1000}s"
        self.log(f"Video: {dur_str}  ({total_f:,} frames @ {fps:.1f} fps)")
        self.log("Scanning at 1 fps — may take several minutes…", 'warning')

        raw_ocr, ms, report_at = [], 0, 60_000
        while ms < duration_ms:
            if self._should_cancel():
                cap.release()
                self.log("OCR cancelled.", 'warning')
                return
            cap.set(cv2.CAP_PROP_POS_MSEC, ms)
            ret, frame = cap.read()
            if not ret:
                break
            text = ocr_frame(frame)
            if text:
                raw_ocr.append((ms, text))
            if ms >= report_at:
                self.log(f"  OCR: {ms / duration_ms * 100:.0f}%  ({ms // 60000}m scanned)", 'dim')
                report_at += 60_000
            ms += 1000

        cap.release()
        self.log(f"Scan complete — {len(raw_ocr)} frames with text")

        if not raw_ocr:
            self.log("No subtitle text detected via OCR.", 'warning')
            return

        lines = deduplicate_ocr_lines(raw_ocr)
        self.log(f"Deduplicated to {len(lines)} entries")

        subs     = build_srt_from_ocr(lines)
        srt_path = str(p.parent / (p.stem + ".srt"))
        subs.save(srt_path, encoding='utf-8')
        self.log(f"Saved OCR subtitles → {p.stem}.srt", 'success')

        if self._should_cancel():
            return

        out = str(p.parent / (p.stem + "_HEB.srt"))
        translate_and_save(subs, out, self.log, self._get_ollama_model(),
                           gemini_key=self._get_gemini_key(),
                           cancel_check=self._should_cancel,
                           progress_cb=self._set_progress)
        if not self._should_cancel():
            Path(srt_path).unlink(missing_ok=True)
            self.log(f"Done! → {Path(out).name}", 'success')


# ─── Entry point ──────────────────────────────────────────────────────────────

def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    SubtitleApp().mainloop()


if __name__ == '__main__':
    main()
