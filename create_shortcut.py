#!/usr/bin/env python3
"""
Creates the icon file and desktop shortcut for Hebrew Subtitle Generator.
Run once:  python create_shortcut.py
"""

import subprocess
import sys
from pathlib import Path


APP_DIR    = Path(__file__).parent.resolve()
APP_SCRIPT = APP_DIR / "subtitle_app.py"
ICON_PATH  = APP_DIR / "icon.ico"
SHORTCUT_NAME = "Hebrew Subtitle Generator"


# ─── Icon generation ──────────────────────────────────────────────────────────

def draw_icon_frame(size: int):
    """Draw one size variant of the app icon."""
    from PIL import Image, ImageDraw

    img  = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # ── Background: dark navy rounded rectangle ──
    radius = max(3, size // 7)
    draw.rounded_rectangle([0, 0, size - 1, size - 1],
                            radius=radius, fill=(14, 18, 48))

    # ── Gold top bar (film strip) ──
    bar_h = max(2, size // 14)
    draw.rectangle([0, 0, size - 1, bar_h], fill=(195, 155, 0))

    # ── Film perforations along gold bar ──
    hole = max(1, size // 24)
    hy   = (bar_h - hole) // 2
    step = max(hole + 2, size // 9)
    for x in range(step, size - step + 1, step):
        draw.rectangle([x, hy, x + hole, hy + hole], fill=(14, 18, 48))

    # ── Gold bottom bar (film strip) ──
    draw.rectangle([0, size - bar_h - 1, size - 1, size - 1], fill=(195, 155, 0))

    # ── Subtitle area: dim overlay at bottom ──
    sub_top = int(size * 0.56)
    draw.rectangle([1, sub_top, size - 2, size - bar_h - 2],
                   fill=(30, 40, 90, 180))

    # ── Subtitle text lines (white bars) ──
    line_h = max(1, size // 18)
    gap    = max(2, size // 14)
    margin = max(3, size // 10)
    y1     = sub_top + gap

    # First line — full width
    draw.rectangle([margin, y1, size - margin, y1 + line_h],
                   fill=(255, 255, 255, 230))

    if size >= 32:
        # Second line — shorter, centred
        y2    = y1 + line_h + gap
        short = int((size - 2 * margin) * 0.68)
        x2    = (size - short) // 2
        draw.rectangle([x2, y2, x2 + short, y2 + line_h],
                       fill=(255, 255, 255, 190))

    # ── Play-button triangle in centre (gold) ──
    if size >= 24:
        cx    = size // 2
        cy    = int(size * 0.30)
        tr    = max(4, size // 9)
        pts   = [
            (cx - tr // 2, cy - tr),
            (cx - tr // 2, cy + tr),
            (cx + tr,      cy),
        ]
        draw.polygon(pts, fill=(230, 185, 0))
        # thin dark outline so it reads well on light backgrounds
        draw.polygon(pts, outline=(14, 18, 48) if size >= 48 else None)

    return img


def create_icon(path: str) -> bool:
    """Generate a multi-resolution .ico file. Returns True on success."""
    try:
        from PIL import Image
        import io

        sizes  = [16, 32, 48, 256]
        frames = [draw_icon_frame(s).convert('RGBA') for s in sizes]

        # Pillow ICO plugin: save largest frame and let it embed all sizes
        # by passing each frame as a separate PNG into the ICO container.
        # Most reliable way: write ICO manually using raw PNG payloads.
        ico_data = _build_ico(frames, sizes)
        with open(path, 'wb') as f:
            f.write(ico_data)

        print(f"  Icon created: {path}")
        return True
    except Exception as e:
        print(f"  Icon generation failed: {e}")
        return False


def _build_ico(frames, sizes) -> bytes:
    """
    Build a valid .ico binary with one PNG-compressed entry per size.
    ICO format: 6-byte header + N*16-byte directory entries + N*PNG blobs.
    """
    import io, struct
    n     = len(frames)
    blobs = []
    for img in frames:
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        blobs.append(buf.getvalue())

    header     = struct.pack('<HHH', 0, 1, n)          # reserved, type=1 (ico), count
    dir_offset = 6 + n * 16
    entries    = b''
    offset     = dir_offset
    for i, (size, blob) in enumerate(zip(sizes, blobs)):
        w = h = size if size < 256 else 0              # 0 means 256 in ICO spec
        entries += struct.pack('<BBBBHHII',
                               w, h,                   # width, height
                               0, 0,                   # color count, reserved
                               1, 32,                  # planes, bit count
                               len(blob), offset)      # data size, offset
        offset += len(blob)

    return header + entries + b''.join(blobs)


# ─── Shortcut creation ────────────────────────────────────────────────────────

def create_shortcut(python_exe: str, script: str, icon: str, name: str) -> bool:
    desktop = Path.home() / "Desktop"
    lnk     = str(desktop / f"{name}.lnk")

    # Method 1 — pywin32 (most reliable)
    try:
        import win32com.client                                    # type: ignore
        shell = win32com.client.Dispatch("WScript.Shell")
        sc    = shell.CreateShortCut(lnk)
        sc.Targetpath        = python_exe
        sc.Arguments         = f'"{script}"'
        sc.WorkingDirectory  = str(Path(script).parent)
        sc.Description       = "Hebrew Subtitle Generator & Translator"
        sc.IconLocation      = icon
        sc.save()
        print(f"  Shortcut created (pywin32): {lnk}")
        return True
    except ImportError:
        pass
    except Exception as e:
        print(f"  pywin32 error: {e}")

    # Method 2 — PowerShell WScript.Shell
    ps = (
        f'$s = (New-Object -ComObject WScript.Shell).CreateShortcut("{lnk}");'
        f'$s.TargetPath = "{python_exe}";'
        f'$s.Arguments = \'"{script}"\';'
        f'$s.WorkingDirectory = "{Path(script).parent}";'
        f'$s.Description = "Hebrew Subtitle Generator and Translator";'
        f'$s.IconLocation = "{icon}";'
        f'$s.Save()'
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-Command", ps],
        capture_output=True, text=True,
    )
    if r.returncode == 0 and Path(lnk).exists():
        print(f"  Shortcut created (PowerShell): {lnk}")
        return True

    print(f"  PowerShell error: {r.stderr.strip()}")
    return False


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("Hebrew Subtitle Generator — Desktop Setup")
    print("=" * 45)

    python_exe = sys.executable
    print(f"  Python:  {python_exe}")
    print(f"  Script:  {APP_SCRIPT}")

    if not APP_SCRIPT.exists():
        print(f"ERROR: subtitle_app.py not found at {APP_SCRIPT}")
        sys.exit(1)

    # 1. Create icon
    print("\nGenerating icon…")
    icon_ok = create_icon(str(ICON_PATH))
    icon_arg = str(ICON_PATH) if icon_ok else f"{python_exe},0"

    # 2. Create desktop shortcut
    print("\nCreating desktop shortcut…")
    ok = create_shortcut(python_exe, str(APP_SCRIPT), icon_arg, SHORTCUT_NAME)

    print()
    if ok:
        print(f'Done!  "{SHORTCUT_NAME}" shortcut is on your Desktop.')
    else:
        print("Shortcut creation failed.")
        print("You can create it manually:")
        print(f"  Target:    {python_exe}")
        print(f"  Arguments: \"{APP_SCRIPT}\"")
        print(f"  Start in:  {APP_DIR}")
        print(f"  Icon:      {ICON_PATH}")


if __name__ == "__main__":
    main()
