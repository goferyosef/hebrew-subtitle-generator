"""
Video Editor - Cut & Splice Tool
=================================
Requirements (install once):
    pip install opencv-python Pillow

Also needs ffmpeg on your system:
    Windows: download from https://ffmpeg.org/download.html
             extract, then add the /bin folder to your PATH
    Or just place ffmpeg.exe in the same folder as this script.

HOW IT WORKS
------------
1. Open a video file.
2. Drag the slider to scrub through the video - the preview updates live.
3. Click "Add Marker" to drop a cut-point at the current position.
   Markers appear as yellow lines on the timeline.
4. Click "Remove Last" to undo the most recent marker.
5. Click "Save Edited Video" when done.
   The app keeps every OTHER segment starting from the SECOND one:
       [discard] marker1 [KEEP] marker2 [discard] marker3 [KEEP] ...
   So: place markers around the parts you want to REMOVE,
   and the kept sections are saved into a new file next to the original.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import subprocess
import shutil
import tempfile
import time

try:
    import cv2
    from PIL import Image, ImageTk
    DEPS_OK = True
except ImportError:
    DEPS_OK = False


# ── constants ─────────────────────────────────────────────────────────────────

PREVIEW_W  = 640
PREVIEW_H  = 360
TL_HEIGHT  = 48          # timeline canvas height
THUMB_STEP = 0.05        # slider step as fraction of total frames


# ── ffmpeg helper ─────────────────────────────────────────────────────────────

def find_ffmpeg():
    """Return path to ffmpeg executable or None."""
    # check same folder as script first
    local = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ffmpeg.exe")
    if os.path.exists(local):
        return local
    found = shutil.which("ffmpeg")
    return found


def seconds_to_ts(s):
    h  = int(s // 3600)
    m  = int((s % 3600) // 60)
    ss = s % 60
    return f"{h:02d}:{m:02d}:{ss:06.3f}"


def cut_segment(ffmpeg, src, t_start, t_end, out_path):
    """Extract [t_start, t_end] from src into out_path (lossless stream copy)."""
    duration = t_end - t_start
    cmd = [
        ffmpeg, "-y",
        "-ss", seconds_to_ts(t_start),
        "-i", src,
        "-t",  seconds_to_ts(duration),
        "-c", "copy",
        "-avoid_negative_ts", "make_zero",
        out_path,
    ]
    subprocess.run(cmd, check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def concat_segments(ffmpeg, segment_paths, out_path):
    """Concatenate segment files into out_path using ffmpeg concat demuxer."""
    with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8") as f:
        list_path = f.name
        for p in segment_paths:
            f.write(f"file '{p.replace(chr(39), chr(39)+chr(92)+chr(39)+chr(39))}'\n")
    try:
        cmd = [
            ffmpeg, "-y",
            "-f", "concat", "-safe", "0",
            "-i", list_path,
            "-c", "copy",
            out_path,
        ]
        subprocess.run(cmd, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    finally:
        os.unlink(list_path)


# ── main application ──────────────────────────────────────────────────────────

class VideoEditor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Video Editor  ·  Cut & Splice")
        self.resizable(False, False)
        self.configure(bg="#111")

        # state
        self._cap        = None       # cv2.VideoCapture
        self._total_frames = 0
        self._fps        = 30.0
        self._duration   = 0.0        # seconds
        self._current_frame = 0
        self._markers    = []         # list of frame numbers, kept sorted
        self._video_path = ""
        self._dragging   = False

        self._build_ui()
        self._center()

        if not DEPS_OK:
            messagebox.showerror(
                "Missing packages",
                "Please install required packages:\n\n"
                "    pip install opencv-python Pillow\n\n"
                "Then restart the app."
            )

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        BG   = "#111"
        CARD = "#1c1c1c"
        ACC  = "#f0a500"     # amber
        TXT  = "#eeebe4"
        DIM  = "#666"

        # ── top bar ──
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=16, pady=(14, 6))

        tk.Label(top, text="✂  Video Editor",
                 font=("Georgia", 17, "bold"), fg=TXT, bg=BG).pack(side="left")

        self._open_btn = tk.Button(
            top, text="Open Video",
            command=self._open_video,
            font=("Courier New", 10, "bold"),
            bg=ACC, fg="#111", activebackground="#d4911a",
            relief="flat", bd=0, padx=16, pady=6, cursor="hand2",
        )
        self._open_btn.pack(side="right")

        tk.Frame(self, height=1, bg="#333").pack(fill="x", padx=16, pady=4)

        # ── preview canvas ──
        preview_frame = tk.Frame(self, bg="#000",
                                 width=PREVIEW_W, height=PREVIEW_H)
        preview_frame.pack_propagate(False)
        preview_frame.pack(padx=16, pady=(6, 0))

        self._canvas = tk.Canvas(
            preview_frame, width=PREVIEW_W, height=PREVIEW_H,
            bg="#000", highlightthickness=0,
        )
        self._canvas.pack()
        self._draw_placeholder()

        # ── time label ──
        time_row = tk.Frame(self, bg=BG)
        time_row.pack(fill="x", padx=16, pady=(4, 0))
        self._time_var = tk.StringVar(value="0:00:00.000  /  0:00:00.000")
        tk.Label(time_row, textvariable=self._time_var,
                 font=("Courier New", 10), fg=ACC, bg=BG).pack(side="left")

        self._markers_var = tk.StringVar(value="Markers: none")
        tk.Label(time_row, textvariable=self._markers_var,
                 font=("Courier New", 9), fg=DIM, bg=BG).pack(side="right")

        # ── timeline canvas ──
        tl_frame = tk.Frame(self, bg=CARD, bd=0)
        tl_frame.pack(fill="x", padx=16, pady=(6, 0))

        self._tl = tk.Canvas(
            tl_frame, height=TL_HEIGHT, bg="#222",
            highlightthickness=1, highlightbackground="#333",
            cursor="sb_h_double_arrow",
        )
        self._tl.pack(fill="x")
        self._tl.bind("<ButtonPress-1>",   self._tl_press)
        self._tl.bind("<B1-Motion>",       self._tl_drag)
        self._tl.bind("<ButtonRelease-1>", self._tl_release)
        self._tl.bind("<Configure>",       self._tl_redraw)

        # ── slider ──
        slider_frame = tk.Frame(self, bg=BG)
        slider_frame.pack(fill="x", padx=16, pady=(2, 0))

        self._slider_var = tk.DoubleVar(value=0)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Amber.Horizontal.TScale",
                        background=BG, troughcolor="#2a2a2a",
                        sliderlength=18, sliderrelief="flat")
        self._slider = ttk.Scale(
            slider_frame, from_=0, to=1000,
            orient="horizontal", variable=self._slider_var,
            style="Amber.Horizontal.TScale",
            command=self._slider_moved,
        )
        self._slider.pack(fill="x")

        # ── control buttons ──
        ctrl = tk.Frame(self, bg=BG)
        ctrl.pack(pady=10)

        btn_cfg = dict(font=("Courier New", 10, "bold"),
                       relief="flat", bd=0, padx=18, pady=8, cursor="hand2")

        self._mark_btn = tk.Button(
            ctrl, text="◆  Add Marker",
            command=self._add_marker,
            bg="#2a2a2a", fg=ACC, activebackground="#333",
            state="disabled", **btn_cfg,
        )
        self._mark_btn.pack(side="left", padx=(0, 8))

        self._undo_btn = tk.Button(
            ctrl, text="↩  Remove Last",
            command=self._remove_last_marker,
            bg="#2a2a2a", fg=TXT, activebackground="#333",
            state="disabled", **btn_cfg,
        )
        self._undo_btn.pack(side="left", padx=(0, 8))

        self._clear_btn = tk.Button(
            ctrl, text="✕  Clear All",
            command=self._clear_markers,
            bg="#2a2a2a", fg="#e06060", activebackground="#333",
            state="disabled", **btn_cfg,
        )
        self._clear_btn.pack(side="left", padx=(0, 20))

        self._save_btn = tk.Button(
            ctrl, text="▶  Save Edited Video",
            command=self._save_video,
            bg=ACC, fg="#111", activebackground="#d4911a",
            state="disabled", **btn_cfg,
        )
        self._save_btn.pack(side="left")

        # ── status bar ──
        self._status_var = tk.StringVar(value="Open a video file to begin.")
        tk.Label(self, textvariable=self._status_var,
                 font=("Courier New", 9), fg=DIM, bg=BG,
                 anchor="w").pack(fill="x", padx=16, pady=(0, 10))

        # ── progress bar (hidden until exporting) ──
        self._progress = ttk.Progressbar(
            self, orient="horizontal", mode="indeterminate", length=400,
        )

    # ── helpers ───────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w  = self.winfo_width()
        h  = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _draw_placeholder(self):
        self._canvas.delete("all")
        self._canvas.create_rectangle(
            0, 0, PREVIEW_W, PREVIEW_H, fill="#000", outline=""
        )
        self._canvas.create_text(
            PREVIEW_W // 2, PREVIEW_H // 2,
            text="No video loaded",
            font=("Georgia", 16), fill="#444",
        )

    def _fmt_time(self, seconds):
        h  = int(seconds // 3600)
        m  = int((seconds % 3600) // 60)
        s  = seconds % 60
        return f"{h}:{m:02d}:{s:06.3f}"

    # ── open video ────────────────────────────────────────────────────────────

    def _open_video(self):
        path = filedialog.askopenfilename(
            title="Open video file",
            filetypes=[
                ("Video files", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv *.webm"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        if self._cap:
            self._cap.release()

        cap = cv2.VideoCapture(path)
        if not cap.isOpened():
            messagebox.showerror("Error", "Could not open video file.")
            return

        self._cap           = cap
        self._video_path    = path
        self._fps           = cap.get(cv2.CAP_PROP_FPS) or 30.0
        self._total_frames  = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        self._duration      = self._total_frames / self._fps
        self._markers       = []
        self._current_frame = 0

        self._slider_var.set(0)
        self._slider.config(to=self._total_frames - 1)

        self._show_frame(0)
        self._update_time_label()
        self._update_markers_label()
        self._tl_redraw()

        name = os.path.basename(path)
        self._status_var.set(
            f"Loaded: {name}  |  "
            f"{self._fmt_time(self._duration)}  |  "
            f"{self._fps:.2f} fps  |  "
            f"{self._total_frames} frames"
        )
        self._mark_btn.config(state="normal")
        self._save_btn.config(state="normal")

    # ── frame display ─────────────────────────────────────────────────────────

    def _show_frame(self, frame_no):
        if not self._cap:
            return
        frame_no = max(0, min(frame_no, self._total_frames - 1))
        self._cap.set(cv2.CAP_PROP_POS_FRAMES, frame_no)
        ret, frame = self._cap.read()
        if not ret:
            return
        self._current_frame = frame_no

        # resize to preview size maintaining aspect ratio
        fh, fw = frame.shape[:2]
        scale  = min(PREVIEW_W / fw, PREVIEW_H / fh)
        nw, nh = int(fw * scale), int(fh * scale)
        resized = cv2.resize(frame, (nw, nh), interpolation=cv2.INTER_AREA)
        rgb     = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)

        img    = Image.fromarray(rgb)
        photo  = ImageTk.PhotoImage(img)

        # center on canvas
        x = (PREVIEW_W - nw) // 2
        y = (PREVIEW_H - nh) // 2
        self._canvas.delete("all")
        self._canvas.create_rectangle(
            0, 0, PREVIEW_W, PREVIEW_H, fill="#000", outline=""
        )
        self._canvas.create_image(x, y, anchor="nw", image=photo)
        self._canvas.image = photo   # prevent GC

    def _update_time_label(self):
        cur = self._current_frame / self._fps
        self._time_var.set(
            f"{self._fmt_time(cur)}  /  {self._fmt_time(self._duration)}"
        )

    # ── slider ────────────────────────────────────────────────────────────────

    def _slider_moved(self, val):
        if not self._cap:
            return
        frame_no = int(float(val))
        self._show_frame(frame_no)
        self._update_time_label()
        self._tl_redraw()

    # ── timeline ──────────────────────────────────────────────────────────────

    def _tl_press(self, event):
        self._dragging = True
        self._tl_seek(event.x)

    def _tl_drag(self, event):
        if self._dragging:
            self._tl_seek(event.x)

    def _tl_release(self, event):
        self._dragging = False

    def _tl_seek(self, x):
        if not self._cap:
            return
        w = self._tl.winfo_width()
        if w <= 0:
            return
        frac     = max(0.0, min(1.0, x / w))
        frame_no = int(frac * (self._total_frames - 1))
        self._slider_var.set(frame_no)
        self._show_frame(frame_no)
        self._update_time_label()
        self._tl_redraw()

    def _tl_redraw(self, event=None):
        tl = self._tl
        w  = tl.winfo_width()
        h  = TL_HEIGHT
        tl.delete("all")

        # background
        tl.create_rectangle(0, 0, w, h, fill="#222", outline="")

        if not self._cap or self._total_frames == 0:
            return

        # shade the KEPT segments (every other, starting from index 1)
        boundaries = [0] + sorted(self._markers) + [self._total_frames]
        for i in range(1, len(boundaries) - 1, 2):
            x1 = int(boundaries[i]     / self._total_frames * w)
            x2 = int(boundaries[i + 1] / self._total_frames * w)
            tl.create_rectangle(x1, 0, x2, h, fill="#2a4a2a", outline="")

        # tick marks every 10 %
        for pct in range(0, 101, 10):
            x = int(pct / 100 * w)
            tl.create_line(x, h - 10, x, h, fill="#444")

        # marker lines
        for m in self._markers:
            x = int(m / self._total_frames * w)
            tl.create_line(x, 0, x, h, fill="#f0a500", width=2)
            tl.create_polygon(x - 5, 0, x + 5, 0, x, 8,
                              fill="#f0a500", outline="")

        # playhead
        x = int(self._current_frame / self._total_frames * w)
        tl.create_line(x, 0, x, h, fill="#ffffff", width=2)
        tl.create_polygon(x - 5, 0, x + 5, 0, x, 8,
                          fill="#ffffff", outline="")

    # ── markers ───────────────────────────────────────────────────────────────

    def _add_marker(self):
        f = self._current_frame
        if f in self._markers:
            self._status_var.set("Marker already exists at this position.")
            return
        self._markers.append(f)
        self._markers.sort()
        self._update_markers_label()
        self._tl_redraw()
        self._undo_btn.config(state="normal")
        self._clear_btn.config(state="normal")
        t = self._fmt_time(f / self._fps)
        self._status_var.set(f"Marker added at {t}  |  {len(self._markers)} marker(s) total")

    def _remove_last_marker(self):
        if not self._markers:
            return
        removed = self._markers.pop()
        t = self._fmt_time(removed / self._fps)
        self._update_markers_label()
        self._tl_redraw()
        if not self._markers:
            self._undo_btn.config(state="disabled")
            self._clear_btn.config(state="disabled")
        self._status_var.set(f"Removed marker at {t}")

    def _clear_markers(self):
        self._markers.clear()
        self._update_markers_label()
        self._tl_redraw()
        self._undo_btn.config(state="disabled")
        self._clear_btn.config(state="disabled")
        self._status_var.set("All markers cleared.")

    def _update_markers_label(self):
        n = len(self._markers)
        if n == 0:
            self._markers_var.set("Markers: none")
        else:
            times = "  |  ".join(self._fmt_time(m / self._fps) for m in self._markers)
            self._markers_var.set(f"Markers ({n}):  {times}")

    # ── save / export ─────────────────────────────────────────────────────────

    def _save_video(self):
        if not self._cap:
            messagebox.showwarning("No video", "Please open a video first.")
            return

        ffmpeg = find_ffmpeg()
        if not ffmpeg:
            messagebox.showerror(
                "ffmpeg not found",
                "ffmpeg is required to export video.\n\n"
                "Download from https://ffmpeg.org/download.html\n"
                "and place ffmpeg.exe in the same folder as this script,\n"
                "or add it to your system PATH."
            )
            return

        if len(self._markers) < 2:
            messagebox.showinfo(
                "Not enough markers",
                "Please add at least 2 markers to define the sections.\n\n"
                "The app keeps segments 2, 4, 6 … and discards 1, 3, 5 …\n"
                "So place markers around sections you want to REMOVE."
            )
            return

        # build list of kept segments (every other, starting at index 1)
        boundaries = sorted([0] + self._markers + [self._total_frames])
        kept = []
        for i in range(1, len(boundaries) - 1, 2):
            t_start = boundaries[i]     / self._fps
            t_end   = boundaries[i + 1] / self._fps
            if t_end - t_start > 0.05:   # skip tiny rounding gaps
                kept.append((t_start, t_end))

        if not kept:
            messagebox.showwarning("Nothing to keep",
                                   "No segments selected to keep. "
                                   "Adjust your markers and try again.")
            return

        # choose output path
        base, ext = os.path.splitext(self._video_path)
        default_out = base + "_edited" + ext
        out_path = filedialog.asksaveasfilename(
            title="Save edited video as…",
            initialfile=os.path.basename(default_out),
            initialdir=os.path.dirname(default_out),
            defaultextension=ext,
            filetypes=[("Video files", f"*{ext}"), ("All files", "*.*")],
        )
        if not out_path:
            return

        # disable UI and start export thread
        self._set_ui_state("disabled")
        self._progress.pack(pady=(0, 6))
        self._progress.start(12)
        self._status_var.set("Exporting… please wait.")

        threading.Thread(
            target=self._run_export,
            args=(ffmpeg, self._video_path, kept, out_path),
            daemon=True,
        ).start()

    def _run_export(self, ffmpeg, src, kept_segments, out_path):
        tmpdir = tempfile.mkdtemp()
        try:
            seg_files = []
            for i, (t_start, t_end) in enumerate(kept_segments):
                seg_out = os.path.join(tmpdir, f"seg_{i:04d}.mp4")
                self._gui(self._status_var.set,
                          f"Cutting segment {i+1}/{len(kept_segments)}…")
                cut_segment(ffmpeg, src, t_start, t_end, seg_out)
                seg_files.append(seg_out)

            self._gui(self._status_var.set, "Joining segments…")
            if len(seg_files) == 1:
                shutil.copy(seg_files[0], out_path)
            else:
                concat_segments(ffmpeg, seg_files, out_path)

            self._gui(self._on_export_done, out_path)

        except subprocess.CalledProcessError as e:
            self._gui(messagebox.showerror, "Export failed",
                      f"ffmpeg returned an error.\nCheck that the video is not corrupted.\n\n{e}")
            self._gui(self._status_var.set, "Export failed.")
        except Exception as e:
            self._gui(messagebox.showerror, "Export failed", str(e))
            self._gui(self._status_var.set, f"Error: {e}")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
            self._gui(self._set_ui_state, "normal")
            self._gui(self._progress.stop)
            self._gui(self._progress.pack_forget)

    def _on_export_done(self, out_path):
        self._status_var.set(f"Done!  Saved: {os.path.basename(out_path)}")
        messagebox.showinfo(
            "Export Complete",
            f"Video saved to:\n{out_path}"
        )

    # ── utilities ─────────────────────────────────────────────────────────────

    def _set_ui_state(self, state):
        for w in (self._open_btn, self._mark_btn,
                  self._undo_btn, self._clear_btn, self._save_btn,
                  self._slider):
            try:
                w.config(state=state)
            except Exception:
                pass

    @staticmethod
    def _gui(fn, *args, **kwargs):
        try:
            fn(*args, **kwargs)
        except Exception:
            pass


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = VideoEditor()
    app.mainloop()
