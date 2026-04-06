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
  1. Groq  (free cloud AI)  — gender-aware, context-aware Hebrew (Llama 3.3 70B)
  2. Google Translate        — free fallback via deep-translator
"""

import difflib
import hashlib
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

# ─── ffmpeg / ffprobe auto-detect ────────────────────────────────────────────
# Prefer imageio_ffmpeg for ffmpeg (full build with all encoders/muxers).
# Use static_ffmpeg for ffprobe.
def _configure_ffmpeg():
    # Always prefer imageio_ffmpeg (full build) for ffmpeg — static_ffmpeg's
    # win32 build is missing the SRT encoder. Prepend its dir so it wins PATH.
    try:
        import imageio_ffmpeg
        src = Path(imageio_ffmpeg.get_ffmpeg_exe())
        dst = src.parent / 'ffmpeg.exe'
        if not dst.exists():
            shutil.copy2(src, dst)
        os.environ['PATH'] = str(src.parent) + ';' + os.environ.get('PATH', '')
    except ImportError:
        pass
    # static_ffmpeg provides ffprobe
    if not shutil.which('ffprobe'):
        try:
            import static_ffmpeg
            static_ffmpeg.add_paths()
        except ImportError:
            pass
_configure_ffmpeg()


# ─── Tesseract auto-detect (finds it even if not in PATH) ────────────────────
_TESSERACT_KNOWN = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    str(Path.home() / 'AppData' / 'Local' / 'Tesseract-OCR' / 'tesseract.exe'),
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

# Groq (free cloud AI — https://console.groq.com)
GROQ_API_URL   = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL     = "llama-3.3-70b-versatile"
GROQ_BATCH_SIZE = 15
GROQ_CONTEXT    = 20      # lines of prior context kept for gender consistency
GROQ_TIMEOUT    = 60
GROQ_CONFIG_PATH = Path.home() / ".hebrew_subtitle_config.json"

# Hebrew RTL marker — prepended to each translated line so players display it correctly
RTL_MARK = '\u200f'

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
        list(args), capture_output=True, text=True, encoding='utf-8', errors='replace',
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
    # Use the absolute stream index from ffprobe data, not the enumeration index
    def idx(i, s): return s.get('index', i)
    text_streams  = [(idx(i, s), s) for i, s in enumerate(streams) if codec(s) in TEXT_CODECS]
    image_streams = [(idx(i, s), s) for i, s in enumerate(streams) if codec(s) in IMAGE_CODECS]
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
    last_err = ""
    for codec in ('srt', 'subrip', 'copy'):
        r = run_cmd('ffmpeg', '-y', '-i', video_path,
                    '-map', f'0:{stream_index}', '-c:s', codec, out_path, timeout=120)
        if Path(out_path).exists() and Path(out_path).stat().st_size > 0:
            return
        last_err = (r.stderr or r.stdout)[-500:]
    raise RuntimeError(f"ffmpeg failed to extract subtitle stream.\n{last_err}")


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


def _tess_lang() -> str:
    """Return best available Tesseract language combo (prefers common subtitle languages)."""
    try:
        import pytesseract
        available = set(pytesseract.get_languages(config=''))
        wanted = ['eng', 'heb', 'ara', 'rus', 'spa', 'fra', 'deu', 'ita', 'por',
                  'chi_sim', 'jpn', 'kor']
        langs = [l for l in wanted if l in available]
        return '+'.join(langs) if langs else 'eng'
    except Exception:
        return 'eng'


def preprocess_for_ocr(region):
    """Takes an already-upscaled BGR region; returns list of threshold variants."""
    import cv2
    import numpy as np
    gray     = cv2.cvtColor(region, cv2.COLOR_BGR2GRAY)
    denoised = cv2.fastNlMeansDenoising(gray, h=10)
    kernel   = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], dtype=np.float32)
    sharp    = cv2.filter2D(denoised, -1, kernel)
    _, t1    = cv2.threshold(sharp, 190, 255, cv2.THRESH_BINARY)
    _, t2    = cv2.threshold(sharp, 100, 255, cv2.THRESH_BINARY_INV)
    t3       = cv2.adaptiveThreshold(sharp, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                      cv2.THRESH_BINARY, 15, 3)
    hsv      = cv2.cvtColor(region, cv2.COLOR_BGR2HSV)
    yellow   = cv2.inRange(hsv, np.array([15, 80, 80]), np.array([40, 255, 255]))
    return [t1, t2, t3, yellow]


def _subtitle_region(frame):
    """Crop bottom 20% (subtitle zone), resize to max 960px wide, upscale 2x."""
    import cv2
    h, w   = frame.shape[:2]
    region = frame[int(h * 0.80):, :]
    if region.size == 0:
        return None
    rh, rw = region.shape[:2]
    if rw > 960:
        scale  = 960 / rw
        region = cv2.resize(region, (960, max(1, int(rh * scale))),
                             interpolation=cv2.INTER_AREA)
    return cv2.resize(region, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)


def _region_thumb_hash(frame) -> str | None:
    """64×16 thumbnail hash of subtitle zone — fast change detector."""
    import cv2
    h, w   = frame.shape[:2]
    region = frame[int(h * 0.80):, :]
    if region.size == 0:
        return None
    small = cv2.resize(cv2.cvtColor(region, cv2.COLOR_BGR2GRAY), (64, 16),
                        interpolation=cv2.INTER_AREA)
    return hashlib.md5(small.tobytes()).hexdigest()


def ocr_frame(frame, tess_lang: str = 'eng') -> str:
    import pytesseract
    region = _subtitle_region(frame)
    if region is None:
        return ''
    best, cfg = '', r'--oem 1 --psm 6 -c tessedit_char_blacklist=|~^_'
    for img in preprocess_for_ocr(region):
        try:
            raw     = pytesseract.image_to_string(img, config=cfg, lang=tess_lang)
            cleaned = ' '.join(raw.split())
            if len(cleaned) < 3:
                continue
            sym = sum(1 for c in cleaned
                      if not (c.isalpha() or c in " ,.!?'-–—")) / max(len(cleaned), 1)
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


# ─── Groq AI Translation ──────────────────────────────────────────────────────

def load_groq_key() -> str:
    try:
        if GROQ_CONFIG_PATH.exists():
            return json.loads(GROQ_CONFIG_PATH.read_text(encoding='utf-8')).get('groq_api_key', '')
    except Exception:
        pass
    return ''

def save_groq_key(key: str):
    try:
        data = {}
        if GROQ_CONFIG_PATH.exists():
            data = json.loads(GROQ_CONFIG_PATH.read_text(encoding='utf-8'))
        data['groq_api_key'] = key.strip()
        GROQ_CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding='utf-8')
    except Exception:
        pass

def check_groq(key: str) -> tuple:
    """Returns (ok: bool, error_msg: str)."""
    if not key:
        return False, "No key provided."
    try:
        payload = json.dumps({
            "model": GROQ_MODEL,
            "messages": [{"role": "user", "content": "hi"}],
            "max_tokens": 5,
        }).encode()
        req = urllib.request.Request(
            GROQ_API_URL, data=payload,
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {key}",
                     "User-Agent": "Mozilla/5.0"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status == 200, ""
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode(errors='replace')[:300]
        except Exception:
            pass
        return False, f"HTTP {e.code} {e.reason}: {body}"
    except urllib.error.URLError as e:
        return False, f"Network error: {e.reason}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def groq_chat(system: str, user: str, key: str, timeout: int = GROQ_TIMEOUT) -> str:
    payload = json.dumps({
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user",   "content": user},
        ],
        "temperature": 0.3,
        "max_tokens": 4096,
    }).encode()
    req = urllib.request.Request(
        GROQ_API_URL, data=payload,
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {key}",
                 "User-Agent": "Mozilla/5.0"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read())["choices"][0]["message"]["content"]


def detect_character_genders(sample_texts: list, key: str, log_cb) -> str:
    if not sample_texts:
        return ""
    # Sample evenly across the whole file (up to 80 lines) for better coverage
    step    = max(1, len(sample_texts) // 80)
    sampled = [sample_texts[i] for i in range(0, len(sample_texts), step)][:80]
    numbered = "\n".join(f"{i+1}. {t}" for i, t in enumerate(sampled))
    user_msg = (
        "Read these subtitle lines and identify character names and their genders.\n"
        "Return ONLY JSON: {\"characters\": [{\"name\": \"...\", \"gender\": \"male|female|unknown\"}, ...]}\n"
        "If no character names are detectable, return {\"characters\": []}\n\n"
        f"Lines:\n{numbered}"
    )
    try:
        response = groq_chat(
            "You are a script analyst. Identify character names and genders from dialogue. Return only valid JSON.",
            user_msg, key, timeout=30,
        )
        m = re.search(r'\{[\s\S]*\}', response)
        if m:
            data  = json.loads(m.group())
            chars = [c for c in data.get('characters', [])
                     if c.get('name') and c.get('gender') != 'unknown']
            if chars:
                summary  = ', '.join(f"{c['name']} ({c['gender']})" for c in chars)
                log_cb(f"  Characters: {summary}", 'dim')
                char_list = '\n'.join(f"  - {c['name']}: {c['gender']}" for c in chars)
                return (
                    "CHARACTER GENDERS (use correct Hebrew gender forms when these characters "
                    "speak or are spoken to):\n" + char_list
                )
    except Exception as e:
        log_cb(f"  Gender detection skipped: {e}", 'dim')
    return ""


def _parse_llm_json(response: str, n: int) -> list:
    m = re.search(r'\[[\s\S]*\]', response)
    if m:
        parsed = json.loads(m.group())
        if isinstance(parsed, list) and len(parsed) >= n:
            return [str(t).strip() for t in parsed[:n]]
    items = re.findall(r'^\s*\d+\.\s*["\']?(.+?)["\']?\s*$', response, re.MULTILINE)
    if len(items) >= n:
        return [t.strip() for t in items[:n]]
    raise ValueError(f"Could not parse {n} items from response")


def _groq_translate_batch(texts: list, key: str, system: str, context_window: list) -> list:
    ctx = ""
    if context_window:
        ctx = "RECENT CONTEXT (already translated — reference only, do NOT retranslate):\n"
        ctx += "\n".join(f"  {t}" for t in context_window) + "\n\n"
    numbered = "\n".join(f'{i+1}. "{t}"' for i, t in enumerate(texts))
    user_msg  = (
        f"{ctx}Translate these {len(texts)} subtitle lines to Hebrew.\n"
        f"Return ONLY a JSON array of exactly {len(texts)} Hebrew strings:\n\n{numbered}"
    )
    return _parse_llm_json(groq_chat(system, user_msg, key), len(texts))


def _groq_full_translate(raw_texts: list, key: str, log_cb, cancel_check=None) -> list:
    clean_texts = [strip_sub_tags(t) for t in raw_texts]
    log_cb("  Detecting character genders…", 'dim')
    gender_block = detect_character_genders([t for t in clean_texts if t], key, log_cb)
    system       = HEBREW_SYSTEM_PROMPT.format(gender_block=gender_block)

    results        = list(raw_texts)
    context_window = []
    total          = len(raw_texts)

    for batch_start in range(0, total, GROQ_BATCH_SIZE):
        if cancel_check and cancel_check():
            log_cb("  Cancelled.", 'warning')
            return results

        batch_raw   = raw_texts[batch_start : batch_start + GROQ_BATCH_SIZE]
        batch_clean = clean_texts[batch_start : batch_start + GROQ_BATCH_SIZE]
        non_empty   = [(j, t) for j, t in enumerate(batch_clean) if t.strip()]

        if not non_empty:
            continue

        # Retry with backoff on 429, fall back to Google only after exhausting retries
        for attempt in range(4):
            try:
                translated = _groq_translate_batch([t for _, t in non_empty], key, system, context_window)
                for (j, _), heb in zip(non_empty, translated):
                    results[batch_start + j] = RTL_MARK + heb
                    context_window.append(heb)
                context_window = context_window[-GROQ_CONTEXT:]
                break
            except urllib.error.HTTPError as e:
                if e.code == 429 and attempt < 3:
                    wait = 15 * (attempt + 1)
                    log_cb(f"  Rate limited — waiting {wait}s…", 'dim')
                    time.sleep(wait)
                else:
                    log_cb(f"  Batch {batch_start // GROQ_BATCH_SIZE + 1} failed ({e}) — Google fallback", 'warning')
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
                    break
            except Exception as e:
                log_cb(f"  Batch {batch_start // GROQ_BATCH_SIZE + 1} failed ({e}) — Google fallback", 'warning')
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
                break

        done = min(batch_start + GROQ_BATCH_SIZE, total)
        if done % 60 == 0 or done == total:
            log_cb(f"  {done}/{total} lines translated", 'dim')

    return results


# ─── Google Translate (fallback) ──────────────────────────────────────────────

def _google_batch_translate(texts: list, log_cb, cancel_check=None) -> list:
    from deep_translator import GoogleTranslator
    translator = GoogleTranslator(source='auto', target='iw')
    results    = list(texts)
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
                log_cb(f"  {i}/{len(texts)} lines translated", 'dim')
        batch_t.append(clean)
        batch_i.append(i)
        batch_chars += len(clean) + len(GOOGLE_BATCH_SEP)

    if batch_t:
        flush(batch_t, batch_i)
    return results


# ─── Translation dispatcher ───────────────────────────────────────────────────

def translate_and_save(subs, out_path: str, log_cb,
                       groq_key: str = None, cancel_check=None):
    import pysubs2

    with_text = [(i, e) for i, e in enumerate(subs.events) if strip_sub_tags(e.text)]
    if not with_text:
        log_cb("No subtitle text to translate.", 'warning')
        subs.save(out_path, encoding='utf-8-sig')
        return

    raw_texts = [e.text for _, e in with_text]
    log_cb(f"Translating {len(raw_texts)} lines…")

    if groq_key:
        log_cb(f"  Groq ({GROQ_MODEL}) — AI, gender-aware")
        translated = _groq_full_translate(raw_texts, groq_key, log_cb, cancel_check)
    else:
        log_cb("  Google Translate (free)")
        try:
            translated = _google_batch_translate(raw_texts, log_cb, cancel_check)
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
    log_cb(f"Saved: {Path(out_path).name}", 'success')


# ─── Main GUI Application ─────────────────────────────────────────────────────

class SubtitleApp(_TK_BASE):

    def __init__(self):
        super().__init__()
        self.title("Hebrew Subtitle Generator")
        self.geometry("780x540")
        self.resizable(True, True)
        self.groq_key      = load_groq_key()
        self._cancel_event = threading.Event()
        self._build_widgets()
        self.after(200, self._check_dependencies)

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_widgets(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # ── Top bar ──
        top = ttk.Frame(self, padding=(12, 10, 12, 4))
        top.grid(row=0, column=0, sticky='ew')
        top.columnconfigure(4, weight=1)

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

        self.groq_btn = ttk.Button(top, text="🔑 Groq Key",
                                   command=self._set_groq_key, width=12)
        self.groq_btn.grid(row=1, column=3, padx=(0, 6))

        # File label (stretchy)
        self.file_var = tk.StringVar(value="No file selected")
        ttk.Label(top, textvariable=self.file_var,
                  foreground='#666').grid(row=1, column=4, sticky='w', padx=(0, 16))

        # Translator selector
        ttk.Label(top, text="Translator:").grid(row=1, column=5, sticky='e', padx=(0, 4))
        self.translator_var = tk.StringVar(value="Google Translate (free)")
        self.trans_combo = ttk.Combobox(
            top, textvariable=self.translator_var,
            values=["Google Translate (free)"], state="readonly", width=32,
        )
        self.trans_combo.grid(row=1, column=6, sticky='w')

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
                self.progress.start(12)
            else:
                self.progress.stop()
                self.progress['value'] = 0
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

        # Groq
        self.log("Checking Groq API key…", 'dim')
        if self.groq_key:
            ok, err = check_groq(self.groq_key)
            if ok:
                self.log(f"  ✓ Groq ready ({GROQ_MODEL})", 'success')
            else:
                self.log(f"  ✗ Groq key invalid or connection failed: {err}", 'warning')
                self.groq_key = ''
        else:
            self.log("  No Groq key set → using Google Translate", 'warning')
            self.log("  For better Hebrew: click '🔑 Groq Key' and enter your free key", 'dim')
            self.log("  Get a free key at: https://console.groq.com", 'dim')

        self._update_translator_options()
        self.log("Ready.", 'success')

    def _update_translator_options(self):
        def _do():
            options = []
            if self.groq_key:
                options.append(f"Groq — {GROQ_MODEL} (AI, gender-aware)")
            options.append("Google Translate (free)")
            self.trans_combo['values'] = options
            self.trans_combo.current(0)
        self.after(0, _do)

    def _use_groq(self) -> str:
        """Return Groq key if Groq is selected, else empty string."""
        if self.translator_var.get().startswith("Groq"):
            return self.groq_key
        return ''

    def _set_groq_key(self):
        """Dialog to enter/update the Groq API key."""
        dlg = tk.Toplevel(self)
        dlg.title("Groq API Key")
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text="Enter your free Groq API key:").grid(
            row=0, column=0, columnspan=2, padx=16, pady=(14, 4), sticky='w')
        ttk.Label(dlg, text="Get one free at https://console.groq.com",
                  foreground='#555').grid(
            row=1, column=0, columnspan=2, padx=16, pady=(0, 8), sticky='w')

        entry = ttk.Entry(dlg, width=52, show='')
        entry.grid(row=2, column=0, columnspan=2, padx=16, pady=(0, 12))
        if self.groq_key:
            entry.insert(0, self.groq_key)

        def _save():
            key = entry.get().strip()
            if not key:
                messagebox.showwarning("Empty key", "Please enter a key.", parent=dlg)
                return
            self.log("Verifying Groq key…", 'dim')
            ok, err = check_groq(key)
            if ok:
                self.groq_key = key
                save_groq_key(key)
                self.log(f"✓ Groq key saved ({GROQ_MODEL} ready)", 'success')
                self._update_translator_options()
                dlg.destroy()
            else:
                self.log(f"  ✗ Groq verification failed: {err}", 'warning')
                messagebox.showerror("Invalid key",
                    f"Could not connect to Groq with this key.\n\nError: {err}", parent=dlg)

        bf = ttk.Frame(dlg)
        bf.grid(row=3, column=0, columnspan=2, pady=(0, 12))
        ttk.Button(bf, text="Save & Verify", command=_save).pack(side='left', padx=6)
        ttk.Button(bf, text="Cancel", command=dlg.destroy).pack(side='left', padx=6)
        entry.focus_set()
        dlg.bind("<Return>", lambda _: _save())

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
            translate_and_save(subs, out, self.log, self._use_groq(),
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
        translate_and_save(subs, out, self.log, self._use_groq(),
                           cancel_check=self._should_cancel)
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

        stream_info = next((s for s in streams if s.get('index') == idx), streams[0])
        lang  = stream_info.get('tags', {}).get('language', 'unknown')
        codec = stream_info.get('codec_name', 'unknown')
        self.log(f"Extracting stream {idx}  (lang={lang}, codec={codec})")

        srt_path = str(p.parent / (p.stem + ".srt"))
        try:
            extract_soft_subtitles(path, idx, srt_path)
        except RuntimeError as e:
            self.log(str(e), 'error')
            self.log("Falling back to OCR mode…", 'warning')
            self._process_video_ocr(path)
            return
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
        translate_and_save(subs, out, self.log, self._use_groq(),
                           cancel_check=self._should_cancel)
        if not self._should_cancel():
            self.log(f"Done! → {out}", 'success')

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

        tess_lang = _tess_lang()
        self.log(f"Video: {dur_str}  ({total_f:,} frames @ {fps:.1f} fps)")
        self.log(f"OCR language(s): {tess_lang}", 'dim')
        self.log("Scanning at ~3 fps with change detection — faster than before…", 'warning')

        step_ms    = 333          # ~3 fps
        raw_ocr    = []
        ms         = 0
        report_at  = 30_000      # progress report every 30 s
        prev_hash  = None
        ocr_calls  = 0
        skipped    = 0

        while ms < duration_ms:
            if self._should_cancel():
                cap.release()
                self.log("OCR cancelled.", 'warning')
                return

            cap.set(cv2.CAP_PROP_POS_MSEC, ms)
            ret, frame = cap.read()
            if not ret:
                break

            # Skip OCR when subtitle zone hasn't changed (saves most of the time)
            curr_hash = _region_thumb_hash(frame)
            if curr_hash is not None and curr_hash == prev_hash:
                skipped += 1
                ms += step_ms
                continue
            prev_hash = curr_hash

            text = ocr_frame(frame, tess_lang)
            ocr_calls += 1
            if text:
                raw_ocr.append((ms, text))

            if ms >= report_at:
                pct = ms / duration_ms * 100
                self.log(f"  OCR: {pct:.0f}%  ({ms // 60000}m)  "
                         f"OCR calls: {ocr_calls}  skipped: {skipped}", 'dim')
                report_at += 30_000

            ms += step_ms

        cap.release()
        self.log(f"Scan complete — {ocr_calls} OCR calls, {skipped} frames skipped, "
                 f"{len(raw_ocr)} with text")

        if not raw_ocr:
            self.log("No subtitle text detected via OCR.", 'warning')
            return

        lines = deduplicate_ocr_lines(raw_ocr)
        self.log(f"Deduplicated to {len(lines)} subtitle entries")

        # Save raw OCR SRT (original language) before translation
        subs     = build_srt_from_ocr(lines)
        raw_path = str(p.parent / (p.stem + "_RAW.srt"))
        subs.save(raw_path, encoding='utf-8')
        self.log(f"Saved raw OCR subtitles → {p.stem}_RAW.srt", 'success')

        if self._should_cancel():
            return

        out = str(p.parent / (p.stem + "_HEB.srt"))
        translate_and_save(subs, out, self.log, self._use_groq(),
                           cancel_check=self._should_cancel)
        if not self._should_cancel():
            self.log(f"Done! → {out}", 'success')


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
