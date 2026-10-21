"""A local worship presentation app for Sahabat Ibadah."""

import csv
import base64
import ctypes
from ctypes import wintypes
import hashlib
import http.server
import json
import os
import re
import secrets
import sqlite3
import subprocess
import shutil
import tempfile
import threading
import time
import zipfile
# Missing camera indexes are expected while checking available devices. Set the
# OpenCV threshold before its first import so those probes do not spam the console.
os.environ.setdefault("OPENCV_LOG_LEVEL", "ERROR")
from html import escape as html_escape
from urllib.parse import parse_qs, urlencode, urlparse
import tkinter as tk
from tkinter import filedialog, colorchooser, font as tkfont
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk


def enable_per_monitor_dpi_awareness():
    """Use physical monitor pixels instead of Windows' scaled virtual coordinates."""
    if os.name != "nt":
        return
    try:
        set_context = ctypes.windll.user32.SetProcessDpiAwarenessContext
        set_context.argtypes = [ctypes.c_void_p]
        set_context.restype = wintypes.BOOL
        if set_context(ctypes.c_void_p(-4)):  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except (AttributeError, OSError):
        pass


# This must run before Tk creates any windows; changing awareness later is ignored by Windows.
enable_per_monitor_dpi_awareness()


APP_DIR = Path(__file__).resolve().parent

# Arahkan folder data aplikasi ke %LOCALAPPDATA%/SahabatIbadah agar bebas izin tulis
DATA_DIR = Path(os.environ.get("LOCALAPPDATA", APP_DIR)) / "SahabatIbadahData"
DATA_DIR.mkdir(parents=True, exist_ok=True)

# Simpan database di folder DATA_DIR
DB_PATH = DATA_DIR / "sahabat_ibadah.db"

# Jika file database lama ada di folder aplikasi/Program Files, pindahkan/salin secara otomatis ke DATA_DIR
OLD_DB_PATH = APP_DIR / "sahabat_ibadah.db"
if OLD_DB_PATH.exists() and not DB_PATH.exists():
    try:
        shutil.copy2(OLD_DB_PATH, DB_PATH)
    except OSError:
        pass
    
POWERPOINT_CACHE = Path(os.environ.get("LOCALAPPDATA", APP_DIR)) / "SahabatIbadahCache" / "powerpoint"
AYT_BOOK_NAMES = (
    "Kejadian", "Keluaran", "Imamat", "Bilangan", "Ulangan", "Yosua", "Hakim-hakim", "Rut",
    "1 Samuel", "2 Samuel", "1 Raja-raja", "2 Raja-raja", "1 Tawarikh", "2 Tawarikh",
    "Ezra", "Nehemia", "Ester", "Ayub", "Mazmur", "Amsal", "Pengkhotbah", "Kidung Agung",
    "Yesaya", "Yeremia", "Ratapan", "Yehezkiel", "Daniel", "Hosea", "Yoel", "Amos",
    "Obaja", "Yunus", "Mikha", "Nahum", "Habakuk", "Zefanya", "Hagai", "Zakharia", "Maleakhi",
    "Matius", "Markus", "Lukas", "Yohanes", "Kisah Para Rasul", "Roma", "1 Korintus", "2 Korintus",
    "Galatia", "Efesus", "Filipi", "Kolose", "1 Tesalonika", "2 Tesalonika", "1 Timotius",
    "2 Timotius", "Titus", "Filemon", "Ibrani", "Yakobus", "1 Petrus", "2 Petrus", "1 Yohanes",
    "2 Yohanes", "3 Yohanes", "Yudas", "Wahyu",
)
ENGLISH_BIBLE_BOOK_NAMES = (
    "Genesis", "Exodus", "Leviticus", "Numbers", "Deuteronomy", "Joshua", "Judges", "Ruth",
    "1 Samuel", "2 Samuel", "1 Kings", "2 Kings", "1 Chronicles", "2 Chronicles", "Ezra",
    "Nehemiah", "Esther", "Job", "Psalms", "Proverbs", "Ecclesiastes", "Song of Solomon",
    "Isaiah", "Jeremiah", "Lamentations", "Ezekiel", "Daniel", "Hosea", "Joel", "Amos",
    "Obadiah", "Jonah", "Micah", "Nahum", "Habakkuk", "Zephaniah", "Haggai", "Zechariah",
    "Malachi", "Matthew", "Mark", "Luke", "John", "Acts", "Romans", "1 Corinthians",
    "2 Corinthians", "Galatians", "Ephesians", "Philippians", "Colossians", "1 Thessalonians",
    "2 Thessalonians", "1 Timothy", "2 Timothy", "Titus", "Philemon", "Hebrews", "James",
    "1 Peter", "2 Peter", "1 John", "2 John", "3 John", "Jude", "Revelation",
)
ENGLISH_BOOK_ALIASES = {
    "psalm": "Psalms", "song of songs": "Song of Solomon", "songs of solomon": "Song of Solomon",
    "canticles": "Song of Solomon", "revelations": "Revelation", "acts of the apostles": "Acts",
}
INDONESIAN_TO_ENGLISH_BOOK = {
    local.casefold(): english for local, english in zip(AYT_BOOK_NAMES, ENGLISH_BIBLE_BOOK_NAMES)
}
ENGLISH_TO_INDONESIAN_BOOK = {
    english.casefold(): local for local, english in zip(AYT_BOOK_NAMES, ENGLISH_BIBLE_BOOK_NAMES)
}


def normalize_bible_book(name, version="AYT"):
    cleaned = re.sub(r"\s+", " ", str(name or "").strip())
    key = cleaned.casefold().rstrip(".")
    key = ENGLISH_BOOK_ALIASES.get(key, key)
    english_version = str(version or "").upper() in ("KJV", "AMP", "NIV")
    if cleaned.isdigit() and 1 <= int(cleaned) <= len(AYT_BOOK_NAMES):
        names = ENGLISH_BIBLE_BOOK_NAMES if english_version else AYT_BOOK_NAMES
        return names[int(cleaned) - 1]
    if english_version:
        return INDONESIAN_TO_ENGLISH_BOOK.get(key.casefold(), key.title())
    return ENGLISH_TO_INDONESIAN_BOOK.get(key.casefold(), cleaned)


class PowerPointRenderError(RuntimeError):
    pass


def available_displays():
    if os.name != "nt":
        return [(0, 0, 1280, 720)]
    class MonitorInfo(ctypes.Structure):
        _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                    ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]
    monitors = []
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC,
                                       ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)

    def collect(_monitor, _dc, _rect, _data):
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(MonitorInfo)
        if ctypes.windll.user32.GetMonitorInfoW(_monitor, ctypes.byref(info)):
            area = info.rcMonitor
            is_primary = bool(info.dwFlags & 1)  # MONITORINFOF_PRIMARY
            monitors.append((not is_primary, area.left, area.top,
                             area.right - area.left, area.bottom - area.top))
        return True

    callback = callback_type(collect)
    ctypes.windll.user32.EnumDisplayMonitors(None, None, callback, 0)
    ordered = sorted(monitors, key=lambda item: (item[0], item[1], item[2]))
    return [(left, top, width, height) for _not_primary, left, top, width, height in ordered] or [(0, 0, 1280, 720)]


def place_window_on_display(window, display, fullscreen=True):
    """Place a Tk window on the requested Windows monitor before showing it."""
    left, top, width, height = display
    if fullscreen:
        window.overrideredirect(True)
        window.attributes("-fullscreen", False)
        window.geometry(f"{width}x{height}{left:+d}{top:+d}")
        window.update_idletasks()
        if os.name == "nt":
            # Tk's -fullscreen can select the primary screen regardless of geometry.
            # Position the native HWND explicitly so HDMI/extended screens work.
            hwnd = wintypes.HWND(window.winfo_id())
            ctypes.windll.user32.SetWindowPos(
                hwnd, wintypes.HWND(-1), left, top, width, height,
                0x0040 | 0x0020 | 0x0200,  # SHOWWINDOW, FRAMECHANGED, NOOWNERZORDER
            )
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.SetFocus(hwnd)
        window.lift()
    else:
        window.attributes("-fullscreen", False)
        window.overrideredirect(False)
        window.geometry(f"{int(width * 0.7)}x{int(height * 0.7)}"
                        f"{left + int(width * 0.15):+d}{top + int(height * 0.15):+d}")


class VideoCaptureSource:
    """Reads a camera or video file in a background thread for smooth Tk updates."""
    def __init__(self, source, loop=True):
        try:
            import cv2
        except ImportError as error:
            raise RuntimeError("OpenCV belum terpasang di virtual environment.") from error
        self.cv2 = cv2
        self.is_file = isinstance(source, str)
        self.loop = bool(loop)
        self.paused = False
        self.capture = cv2.VideoCapture(source) if self.is_file else cv2.VideoCapture(source, cv2.CAP_DSHOW)
        if not self.capture.isOpened() and not self.is_file:
            self.capture.release()
            self.capture = cv2.VideoCapture(source)
        if not self.capture.isOpened():
            self.capture.release()
            raise RuntimeError("Sumber video tidak dapat dibuka. Periksa izin kamera atau pilih perangkat lain.")
        if not self.is_file:
            self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        self.lock = threading.Lock()
        self._reference_lock = threading.Lock()
        self._references = 1
        self._released = False
        self.stop_event = threading.Event()
        self.frame = None
        self.thread = threading.Thread(target=self._read_loop, daemon=True)
        self.thread.start()

    def _read_loop(self):
        while not self.stop_event.is_set():
            if self.paused:
                time.sleep(0.03)
                continue
            ok, frame = self.capture.read()
            if ok:
                with self.lock:
                    self.frame = frame
            elif self.is_file:
                if self.loop:
                    self.capture.set(self.cv2.CAP_PROP_POS_FRAMES, 0)
                else:
                    self.paused = True
            else:
                time.sleep(0.03)

    def latest_frame(self):
        with self.lock:
            return self.frame.copy() if self.frame is not None else None

    def acquire(self):
        """Retain a running camera capture for another output consumer."""
        with self._reference_lock:
            if self._released:
                raise RuntimeError("Sumber kamera sudah dihentikan.")
            self._references += 1
        return self

    def set_loop(self, enabled):
        self.loop = bool(enabled)

    def set_paused(self, paused):
        self.paused = bool(paused)

    def close(self):
        with self._reference_lock:
            if self._released:
                return
            self._references -= 1
            if self._references > 0:
                return
            self._released = True
        self.stop_event.set()
        self.capture.release()
        self.thread.join(timeout=1)


def frame_to_photo(frame, size, master):
    """Scale and crop a BGR frame to fill a Tk image area."""
    if frame is None:
        return None
    import cv2
    height, width = frame.shape[:2]
    target_width, target_height = max(1, size[0]), max(1, size[1])
    scale = max(target_width / width, target_height / height)
    resized_width = max(target_width, int(width * scale))
    resized_height = max(target_height, int(height * scale))
    image = cv2.resize(frame, (resized_width, resized_height),
                       interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
    x = (resized_width - target_width) // 2
    y = (resized_height - target_height) // 2
    image = image[y:y + target_height, x:x + target_width]
    ok, encoded = cv2.imencode(".png", image)
    if not ok:
        return None
    data = base64.b64encode(encoded.tobytes()).decode("ascii")
    return tk.PhotoImage(master=master, data=data, format="png")


_STROKE_OFFSETS = ((-1, 0), (1, 0), (0, -1), (0, 1),
                   (-1, -1), (-1, 1), (1, -1), (1, 1))
DEFAULT_STROKE_COLOR = "#000000"
DEFAULT_STROKE_WIDTH = 2


def draw_text_with_stroke(canvas, x, y, text, font, fill, stroke_color=None,
                          stroke_width=0, wrap_width=None, justify="center",
                          anchor="center", tags=(), state="normal"):
    """Draw canvas text with an optional outline (stroke).

    Tk has no native text stroke, so the outline is a ring of offset copies
    painted in the stroke colour with the normal text drawn on top.
    """
    options = dict(text=text, font=font, width=wrap_width or 0, justify=justify,
                   anchor=anchor, tags=tags, state=state)
    items = []
    try:
        thickness = max(0, int(stroke_width))
    except (TypeError, ValueError):
        thickness = 0
    if stroke_color and thickness:
        for step in range(thickness, 0, -1):
            for dx, dy in _STROKE_OFFSETS:
                items.append(canvas.create_text(x + dx * step, y + dy * step,
                                                fill=stroke_color, **options))
    items.append(canvas.create_text(x, y, fill=fill, **options))
    return items


def measure_wrapped_text(canvas, text, font, wrap_width):
    """Return the pixel height of text wrapped at wrap_width, without showing it."""
    probe = canvas.create_text(-20000, -20000, text=text, font=font, width=wrap_width,
                               justify="center", anchor="nw")
    try:
        bounds = canvas.bbox(probe)
    finally:
        canvas.delete(probe)
    return (bounds[3] - bounds[1]) if bounds else 0


OVERLAY_POSITIONS = ("Atas", "Tengah", "Bawah")
OVERLAY_ALIGNS = ("Kiri", "Tengah", "Kanan")
DEFAULT_OVERLAY_LAYOUT = {"position": "Bawah", "align": "Tengah", "scale": 100,
                          "left": 6, "right": 6, "top": 6, "bottom": 6,
                          "offset_x": 0, "offset_y": 0}
_OVERLAY_ANCHORS = {
    "Atas": {"Kiri": "nw", "Tengah": "n", "Kanan": "ne"},
    "Tengah": {"Kiri": "w", "Tengah": "center", "Kanan": "e"},
    "Bawah": {"Kiri": "sw", "Tengah": "s", "Kanan": "se"},
}
_OVERLAY_JUSTIFY = {"Kiri": "left", "Tengah": "center", "Kanan": "right"}


def _overlay_margin(size, percent):
    try:
        value = max(0, min(45, float(percent)))
    except (TypeError, ValueError):
        value = 6
    return max(10, int(size * value / 100))


def overlay_text_layout(canvas, text, width, height, settings):
    """Place the live overlay text inside the margins at the requested size."""
    settings = settings or {}
    position = settings.get("position", "Bawah")
    if position not in OVERLAY_POSITIONS:
        position = "Bawah"
    align = settings.get("align", "Tengah")
    if align not in OVERLAY_ALIGNS:
        align = "Tengah"
    try:
        scale = max(40, min(150, int(settings.get("scale", 100))))
    except (TypeError, ValueError):
        scale = 100

    left = _overlay_margin(width, settings.get("left", 6))
    right = width - _overlay_margin(width, settings.get("right", 6))
    top = _overlay_margin(height, settings.get("top", 6))
    bottom = height - _overlay_margin(height, settings.get("bottom", 6))
    if right - left < 80:
        right = min(width - 8, left + 80)
    if bottom - top < 60:
        bottom = min(height - 8, top + 60)
    wrap_width = max(80, right - left)
    available_height = max(60, bottom - top)

    max_font = max(12, int(min(120, width // 34, height // 14) * scale / 100))
    font_size = max_font
    while font_size > 10 and measure_wrapped_text(
            canvas, text, ("Segoe UI", font_size, "bold"), wrap_width) > available_height:
        font_size -= 1

    if align == "Kiri":
        x = left
    elif align == "Kanan":
        x = right
    else:
        x = (left + right) / 2
    if position == "Atas":
        y = top
    elif position == "Bawah":
        y = bottom
    else:
        y = (top + bottom) / 2
    try:
        offset_x = max(-80.0, min(80.0, float(settings.get("offset_x", 0))))
    except (TypeError, ValueError):
        offset_x = 0.0
    try:
        offset_y = max(-80.0, min(80.0, float(settings.get("offset_y", 0))))
    except (TypeError, ValueError):
        offset_y = 0.0
    x += width * offset_x / 100
    y += height * offset_y / 100
    return {"x": x, "y": y, "anchor": _OVERLAY_ANCHORS[position][align],
            "justify": _OVERLAY_JUSTIFY[align], "wrap": wrap_width,
            "font_size": font_size, "left": left, "right": right,
            "top": top, "bottom": bottom,
            "offset_x": offset_x, "offset_y": offset_y}


class SlideCanvas(tk.Canvas):
    """Canvas that draws one slide: a background colour, a photo, or stroked text.

    Tk labels cannot outline text, so the projector and its previews draw on
    this canvas instead and repaint the last slide when the widget is resized.
    """

    def __init__(self, master, **kwargs):
        kwargs.setdefault("highlightthickness", 0)
        kwargs.setdefault("bd", 0)
        kwargs.setdefault("bg", "black")
        super().__init__(master, **kwargs)
        self._slide = None
        self._repaint_after_id = None
        self.bind("<Configure>", self._on_configure, add="+")

    def _on_configure(self, _event=None):
        if self._repaint_after_id:
            try:
                self.after_cancel(self._repaint_after_id)
            except tk.TclError:
                pass
        self._repaint_after_id = self.after(80, self._repaint)

    def _repaint(self):
        self._repaint_after_id = None
        self._render()

    def show_photo(self, photo, background="#000000"):
        self._slide = {"photo": photo, "background": background, "text": ""}
        self._render()

    def show_text(self, text, color="#ffffff", background="#000000",
                  font=("Segoe UI", 18, "bold"), wrap=None, photo=None,
                  stroke_color=None, stroke_width=0):
        self._slide = {"photo": photo, "background": background, "text": str(text or ""),
                       "color": color, "font": font, "wrap": wrap,
                       "stroke_color": stroke_color, "stroke_width": stroke_width}
        self._render()

    def clear(self, background="#000000"):
        self._slide = {"photo": None, "background": background, "text": ""}
        self._render()

    def _size(self):
        width, height = self.winfo_width(), self.winfo_height()
        if width < 10 or height < 10:
            width = max(width, self.winfo_reqwidth(), 320)
            height = max(height, self.winfo_reqheight(), 180)
        return width, height

    def _render(self):
        state = self._slide
        if not state:
            return
        self.delete("all")
        self.configure(bg=state["background"])
        width, height = self._size()
        if state.get("photo") is not None:
            self.create_image(width // 2, height // 2, image=state["photo"], tags="slide")
        text = state.get("text") or ""
        if text:
            wrap_width = state.get("wrap") or max(60, width - 32)
            draw_text_with_stroke(self, width // 2, height // 2, text, state["font"],
                                  state["color"], stroke_color=state["stroke_color"],
                                  stroke_width=state["stroke_width"],
                                  wrap_width=wrap_width, justify="center",
                                  anchor="center", tags="slide")


class LiveProjector(tk.Toplevel):
    def __init__(self, parent, source, display, overlay_text, overlay_layout, step_lyrics,
                 still_image_path=None, image_mode="camera", second_source=None,
                 camera_mode="primary", overlay_text_color=None, overlay_stroke=None,
                 overlay_background=None):
        super().__init__(parent)
        try:
            self.capture = VideoCaptureSource(source)
        except Exception:
            self.destroy()
            raise
        self.second_capture = None
        if second_source is not None:
            try:
                self.second_capture = VideoCaptureSource(second_source)
            except Exception:
                self.capture.close()
                self.destroy()
                raise
        self.overlay_text = overlay_text
        self.overlay_text_color = overlay_text_color or (lambda: "#ffffff")
        self.overlay_stroke = overlay_stroke or (lambda: (None, 0))
        self.overlay_background = overlay_background or (lambda: True)
        self.overlay_layout = overlay_layout or (lambda: dict(DEFAULT_OVERLAY_LAYOUT))
        self.step_lyrics = step_lyrics
        self._overlay_signature = None
        self.display = display
        self.image_mode = image_mode
        self.camera_mode = camera_mode
        self.still_image = None
        self.still_cache = {}
        self.camera_image_items = {}
        self.still_image_item = None
        self.camera_photos = {}
        self.still_photo = None
        self.title("Sahabat Ibadah — Kamera Live")
        self.configure(bg="black")
        self.canvas = tk.Canvas(self, bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        left, top, monitor_width, monitor_height = display
        self.fullscreen = len(available_displays()) > 1
        if self.fullscreen:
            width, height, left, top = monitor_width, monitor_height, left, top
        else:
            width, height = int(monitor_width * 0.7), int(monitor_height * 0.7)
            left += (monitor_width - width) // 2
            top += (monitor_height - height) // 2
        self.geometry(f"{width}x{height}+{left}+{top}")
        self.bind("<Escape>", self.close_projector)
        self.bind("<F11>", self.toggle_fullscreen)
        self.bind("<Right>", lambda _event: self.step_lyrics(1))
        self.bind("<space>", lambda _event: self.step_lyrics(1))
        self.bind("<Left>", lambda _event: self.step_lyrics(-1))
        self.protocol("WM_DELETE_WINDOW", self.close_projector)
        self._key_bindings = [
            ("<Right>", self.bind_all("<Right>", lambda _event: self.step_lyrics(1), add="+")),
            ("<space>", self.bind_all("<space>", lambda _event: self.step_lyrics(1), add="+")),
            ("<Left>", self.bind_all("<Left>", lambda _event: self.step_lyrics(-1), add="+")),
            ("<Escape>", self.bind_all("<Escape>", self.close_projector, add="+")),
        ]
        self._closed = False
        if still_image_path:
            try:
                self.set_still_image(still_image_path)
            except Exception:
                self.capture.close()
                self.destroy()
                raise
        self.after_idle(self._apply_display_mode)
        self.after(30, self.update_frame)

    def set_still_image(self, path):
        try:
            image = self.capture.cv2.imread(str(path), self.capture.cv2.IMREAD_COLOR)
        except Exception as error:
            raise RuntimeError("File gambar tidak dapat dibaca oleh OpenCV.") from error
        if image is None:
            raise RuntimeError("File gambar tidak dapat dibuka. Gunakan PNG, JPG, BMP, atau WebP.")
        self.still_image = image
        self.still_cache.clear()

    def set_image_mode(self, mode):
        if mode not in {"camera", "split", "image"}:
            return
        if mode != "camera" and self.still_image is None:
            raise RuntimeError("Pilih gambar terlebih dahulu sebelum menampilkannya.")
        self.image_mode = mode

    def set_camera_mode(self, mode):
        if mode not in {"primary", "secondary", "split", "hidden"}:
            return
        if mode in {"secondary", "split"} and self.second_capture is None:
            raise RuntimeError("Pilih kamera kedua terlebih dahulu.")
        self.camera_mode = mode

    def _draw_camera(self, frame, x, y, region_width, region_height, slot="primary", fit=False):
        if frame is None or region_width < 2 or region_height < 2:
            return
        cv2 = self.capture.cv2
        frame_height, frame_width = frame.shape[:2]
        if fit:
            # Preserve the full camera image in split layouts; use a soft, enlarged
            # copy behind it to fill the panel without cropping or black bars.
            scale = min(region_width / frame_width, region_height / frame_height)
            fitted_width = max(1, min(region_width, int(frame_width * scale)))
            fitted_height = max(1, min(region_height, int(frame_height * scale)))
            fitted = cv2.resize(frame, (fitted_width, fitted_height),
                                interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
            fill_scale = max(region_width / frame_width, region_height / frame_height)
            fill = cv2.resize(frame, (max(region_width, int(frame_width * fill_scale)),
                                        max(region_height, int(frame_height * fill_scale))),
                              interpolation=cv2.INTER_AREA if fill_scale < 1 else cv2.INTER_LINEAR)
            x0 = (fill.shape[1] - region_width) // 2
            y0 = (fill.shape[0] - region_height) // 2
            image = cv2.GaussianBlur(fill[y0:y0 + region_height, x0:x0 + region_width], (0, 0), 22)
            image = cv2.addWeighted(image, 0.42, image, 0, 0)
            foreground_x = (region_width - fitted_width) // 2
            foreground_y = (region_height - fitted_height) // 2
            image[foreground_y:foreground_y + fitted_height,
                  foreground_x:foreground_x + fitted_width] = fitted
        else:
            # Full-screen single-camera mode fills the monitor by center-cropping.
            scale = max(region_width / frame_width, region_height / frame_height)
            resized_width = max(region_width, int(frame_width * scale))
            resized_height = max(region_height, int(frame_height * scale))
            image = cv2.resize(frame, (resized_width, resized_height),
                               interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
            crop_x = (resized_width - region_width) // 2
            crop_y = (resized_height - region_height) // 2
            image = image[crop_y:crop_y + region_height, crop_x:crop_x + region_width]
        ok, encoded = cv2.imencode(".png", image)
        if not ok:
            return
        data = base64.b64encode(encoded.tobytes()).decode("ascii")
        photo = tk.PhotoImage(master=self, data=data, format="png")
        self.camera_photos[slot] = photo
        center = (x + region_width // 2, y + region_height // 2)
        item = self.camera_image_items.get(slot)
        if item is None:
            self.camera_image_items[slot] = self.canvas.create_image(*center, image=photo)
        else:
            self.canvas.itemconfigure(item, image=photo, state="normal")
            self.canvas.coords(item, *center)

    def _draw_still_image(self, x, y, region_width, region_height):
        if self.still_image is None or region_width < 2 or region_height < 2:
            return
        key = (region_width, region_height)
        photo = self.still_cache.get(key)
        if photo is None:
            cv2 = self.capture.cv2
            image_height, image_width = self.still_image.shape[:2]
            scale = min(region_width / image_width, region_height / image_height)
            image = cv2.resize(self.still_image,
                               (max(1, int(image_width * scale)), max(1, int(image_height * scale))),
                               interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
            ok, encoded = cv2.imencode(".png", image)
            if not ok:
                return
            data = base64.b64encode(encoded.tobytes()).decode("ascii")
            photo = tk.PhotoImage(master=self, data=data, format="png")
            self.still_cache[key] = photo
        self.still_photo = photo
        center = (x + region_width // 2, y + region_height // 2)
        if self.still_image_item is None:
            self.still_image_item = self.canvas.create_image(*center, image=self.still_photo)
        else:
            self.canvas.itemconfigure(self.still_image_item, image=self.still_photo, state="normal")
            self.canvas.coords(self.still_image_item, *center)

    def _hide_canvas_item(self, item):
        if item is not None:
            self.canvas.itemconfigure(item, state="hidden")

    def _apply_display_mode(self):
        if self.fullscreen:
            place_window_on_display(self, self.display, fullscreen=True)
        else:
            place_window_on_display(self, self.display, fullscreen=False)
        self.focus_force()

    def toggle_fullscreen(self, _event=None):
        self.fullscreen = not self.fullscreen
        place_window_on_display(self, self.display, fullscreen=self.fullscreen)
        return "break"

    def update_frame(self):
        if self._closed or not self.winfo_exists():
            return
        frame = self.capture.latest_frame()
        width, height = max(1, self.canvas.winfo_width()), max(1, self.canvas.winfo_height())
        self.canvas.delete("live-divider")
        if self.image_mode == "image" and self.still_image is not None:
            for camera_item in self.camera_image_items.values():
                self._hide_canvas_item(camera_item)
            self._draw_still_image(0, 0, width, height)
        elif self.image_mode == "split" and self.still_image is not None:
            if self.camera_mode == "hidden":
                for camera_item in self.camera_image_items.values():
                    self._hide_canvas_item(camera_item)
            else:
                camera_slot = "secondary" if self.camera_mode == "secondary" else "primary"
                selected_capture = self.second_capture if camera_slot == "secondary" else self.capture
                for slot, existing_item in self.camera_image_items.items():
                    if slot != camera_slot:
                        self._hide_canvas_item(existing_item)
                self._draw_camera(selected_capture.latest_frame(), 0, 0, width // 2, height,
                                  camera_slot, fit=True)
            self._draw_still_image(width // 2, 0, width - width // 2, height)
            self.canvas.create_line(width // 2, 0, width // 2, height, fill="#d8dde3", width=3,
                                    tags="live-divider")
        elif self.camera_mode == "hidden":
            for camera_item in self.camera_image_items.values():
                self._hide_canvas_item(camera_item)
            self._hide_canvas_item(self.still_image_item)
        elif self.camera_mode == "secondary" and self.second_capture is not None:
            self._hide_canvas_item(self.camera_image_items.get("primary"))
            self._hide_canvas_item(self.still_image_item)
            self._draw_camera(self.second_capture.latest_frame(), 0, 0, width, height, "secondary")
        elif self.camera_mode == "split" and self.second_capture is not None:
            half_width = width // 2
            self._draw_camera(frame, 0, 0, half_width, height, "primary", fit=True)
            self._draw_camera(self.second_capture.latest_frame(), half_width, 0,
                              width - half_width, height, "secondary", fit=True)
            self.canvas.create_line(half_width, 0, half_width, height, fill="#d8dde3", width=3,
                                    tags="live-divider")
            self._hide_canvas_item(self.still_image_item)
        else:
            self._hide_canvas_item(self.still_image_item)
            self._hide_canvas_item(self.camera_image_items.get("secondary"))
            self._draw_camera(frame, 0, 0, width, height, "primary")
        text = self.overlay_text()
        try:
            stroke_color, stroke_width = self.overlay_stroke()
            stroke_width = max(0, int(stroke_width))
        except (TypeError, ValueError):
            stroke_color, stroke_width = None, 0
        if not stroke_color:
            stroke_width = 0
        layout_settings = dict(self.overlay_layout() or {})
        show_background = bool(self.overlay_background())
        signature = (text, self.overlay_text_color(), stroke_color, stroke_width,
                     width, height, show_background,
                     tuple(sorted(layout_settings.items())))
        if signature == self._overlay_signature:
            # Nothing changed: keep the existing items (cheaper than rebuilding
            # dozens of outline copies on every camera frame) but keep on top.
            self.canvas.tag_raise("live-overlay")
        else:
            self.canvas.delete("live-overlay")
            self._overlay_signature = signature
            if text.strip():
                layout = overlay_text_layout(self.canvas, text, width, height, layout_settings)
                text_ids = draw_text_with_stroke(
                    self.canvas, layout["x"], layout["y"], text,
                    ("Segoe UI", layout["font_size"], "bold"), self.overlay_text_color(),
                    stroke_color=stroke_color, stroke_width=stroke_width,
                    wrap_width=layout["wrap"], justify=layout["justify"],
                    anchor=layout["anchor"], tags="live-overlay")
                text_id = text_ids[-1]
                bounds = self.canvas.bbox(text_id)
                if bounds:
                    if stroke_width:
                        bounds = (bounds[0] - stroke_width, bounds[1] - stroke_width,
                                  bounds[2] + stroke_width, bounds[3] + stroke_width)
                    padding = max(8, layout["font_size"] // 3)
                    if show_background:
                        background = self.canvas.create_rectangle(max(4, bounds[0] - padding),
                                                                 max(4, bounds[1] - padding),
                                                                 min(width - 4, bounds[2] + padding),
                                                                 min(height - 4, bounds[3] + padding),
                                                                 fill="#101820", stipple="gray50", outline="",
                                                                 tags="live-overlay")
                        self.canvas.tag_lower(background, text_ids[0])
        self.after(40, self.update_frame)

    def close_projector(self, _event=None):
        if self._closed:
            return "break"
        self._closed = True
        for sequence, binding_id in self._key_bindings:
            if binding_id:
                lines = self.tk.call("bind", "all", sequence).split("\n")
                prefix = f'if {{"[{binding_id} '
                remaining = "\n".join(line for line in lines if not line.startswith(prefix))
                self.tk.call("bind", "all", sequence, remaining)
                self.deletecommand(binding_id)
        self._key_bindings.clear()
        self.capture.close()
        if self.second_capture is not None:
            self.second_capture.close()
            self.second_capture = None
        self.destroy()
        if self.master.winfo_exists():
            self.master.lift()
            self.master.focus_force()
        return "break"


def render_powerpoint_slides(path, max_size=(1600, 900)):
    """Render PPTX slides through installed Microsoft PowerPoint for visual fidelity."""
    path = Path(path).resolve()
    if not path.is_file():
        raise PowerPointRenderError(f"File PowerPoint tidak ditemukan:\n{path}")
    stat = path.stat()
    max_width, max_height = max_size
    key = hashlib.sha256(
        f"{path}|{stat.st_size}|{stat.st_mtime_ns}|{max_width}x{max_height}".encode("utf-8")).hexdigest()[:24]
    output_dir = POWERPOINT_CACHE / key
    images = sorted(output_dir.glob("Slide*.PNG"),
                    key=lambda image: int(re.search(r"(\d+)$", image.stem).group(1))
                    if re.search(r"(\d+)$", image.stem) else 0) if output_dir.exists() else []
    if images:
        return [str(image) for image in images]

    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Tambahkan penanganan agar PowerPoint COM automation dapat berjalan lebih stabil tanpa kendala izin
    script = r"""
param([string]$InputFile, [string]$OutputDir, [int]$MaxWidth, [int]$MaxHeight)
$ErrorActionPreference = 'Stop'
$app = $null
$deck = $null
try {
    # Pastikan COM Object dibuat dengan bersih
    $app = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
    $app.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
    
    # Buka presentasi dalam mode ReadOnly (-1) dan dengan jendela tersembunyi (0)
    $deck = $app.Presentations.Open($InputFile, -1, 0, 0)
    $ratio = [Math]::Min($MaxWidth / $deck.PageSetup.SlideWidth, $MaxHeight / $deck.PageSetup.SlideHeight)
    $width = [int][Math]::Round($deck.PageSetup.SlideWidth * $ratio)
    $height = [int][Math]::Round($deck.PageSetup.SlideHeight * $ratio)
    $deck.Export($OutputDir, 'PNG', $width, $height)
}
finally {
    if ($deck -ne $null) { $deck.Close() }
    if ($app -ne $null) { 
        $app.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
    }
}
"""
    script_path = None
    try:
        # Buat script sementara di folder TEMP lokal pengguna
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", encoding="utf-8", delete=False,
                                         dir=tempfile.gettempdir()) as temporary_script:
            temporary_script.write(script)
            script_path = temporary_script.name
            
        process = subprocess.run(
            ["powershell.exe", "-NoLogo", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
             "-File", script_path, str(path), str(output_dir), str(max_width), str(max_height)],
            capture_output=True, text=True, timeout=180, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired) as error:
        raise PowerPointRenderError(f"Microsoft PowerPoint tidak dapat merender file ini.\n{error}") from error
    finally:
        if script_path and os.path.exists(script_path):
            try:
                os.remove(script_path)
            except OSError:
                pass

    images = sorted(output_dir.glob("Slide*.PNG"),
                    key=lambda image: int(re.search(r"(\d+)$", image.stem).group(1))
                    if re.search(r"(\d+)$", image.stem) else 0)
    if process.returncode != 0 or not images:
        detail = (process.stderr or process.stdout).strip()
        raise PowerPointRenderError(
            "PowerPoint tidak berhasil merender slide. Pastikan Microsoft PowerPoint terpasang "
            "dan file dapat dibuka secara manual.\n\n" + (detail[-1500:] if detail else "Tidak ada gambar slide yang dihasilkan."))
    return [str(image) for image in images]


class Store:
    def __init__(self):
        self.db = sqlite3.connect(DB_PATH)
        self.db.row_factory = sqlite3.Row
        self.db.executescript("""
            PRAGMA foreign_keys = ON;
            CREATE TABLE IF NOT EXISTS songs (
                id INTEGER PRIMARY KEY, title TEXT NOT NULL, author TEXT DEFAULT '',
                lyrics TEXT NOT NULL DEFAULT '', updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS schedules (
                id INTEGER PRIMARY KEY, title TEXT NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS schedule_items (
                id INTEGER PRIMARY KEY, schedule_id INTEGER NOT NULL REFERENCES schedules(id) ON DELETE CASCADE,
                kind TEXT NOT NULL, title TEXT NOT NULL, content TEXT NOT NULL DEFAULT '', position INTEGER NOT NULL,
                media_path TEXT DEFAULT '', operator_note TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS bible_verses (
                id INTEGER PRIMARY KEY, version TEXT NOT NULL DEFAULT 'AYT',
                book TEXT NOT NULL, abbr TEXT NOT NULL DEFAULT '',
                chapter INTEGER NOT NULL, verse INTEGER NOT NULL, text TEXT NOT NULL,
                title TEXT NOT NULL DEFAULT '', UNIQUE(version, book, chapter, verse)
            );
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY, value TEXT NOT NULL
            );
        """)
        bible_columns = {row["name"] for row in self.db.execute("PRAGMA table_info(bible_verses)")}
        if "version" not in bible_columns:
            self.db.execute("BEGIN IMMEDIATE")
            self.db.execute("ALTER TABLE bible_verses RENAME TO bible_verses_legacy")
            self.db.execute("""CREATE TABLE bible_verses (
                id INTEGER PRIMARY KEY, version TEXT NOT NULL DEFAULT 'AYT',
                book TEXT NOT NULL, abbr TEXT NOT NULL DEFAULT '', chapter INTEGER NOT NULL,
                verse INTEGER NOT NULL, text TEXT NOT NULL, title TEXT NOT NULL DEFAULT '',
                UNIQUE(version, book, chapter, verse))""")
            self.db.execute("""INSERT INTO bible_verses(id,version,book,abbr,chapter,verse,text,title)
                SELECT id,'AYT',book,abbr,chapter,verse,text,title FROM bible_verses_legacy""")
            self.db.execute("DROP TABLE bible_verses_legacy")
            self.db.commit()
        else:
            unique_keys = set()
            for index in self.db.execute("PRAGMA index_list(bible_verses)"):
                if index["unique"]:
                    columns = tuple(row["name"] for row in self.db.execute(
                        f"PRAGMA index_info('{index['name']}')"))
                    unique_keys.add(columns)
            if ("version", "book", "chapter", "verse") not in unique_keys:
                # Rebuild instead of only adding an index: some intermediate databases
                # still have UNIQUE(book, chapter, verse), which blocks importing a
                # second translation even when a version-aware index is also present.
                self.db.execute("BEGIN IMMEDIATE")
                self.db.execute("ALTER TABLE bible_verses RENAME TO bible_verses_legacy")
                self.db.execute("""CREATE TABLE bible_verses (
                    id INTEGER PRIMARY KEY, version TEXT NOT NULL DEFAULT 'AYT',
                    book TEXT NOT NULL, abbr TEXT NOT NULL DEFAULT '', chapter INTEGER NOT NULL,
                    verse INTEGER NOT NULL, text TEXT NOT NULL, title TEXT NOT NULL DEFAULT '',
                    UNIQUE(version, book, chapter, verse))""")
                self.db.execute("""INSERT INTO bible_verses(id,version,book,abbr,chapter,verse,text,title)
                    SELECT id,version,book,abbr,chapter,verse,text,title FROM bible_verses_legacy
                    WHERE id IN (SELECT MIN(id) FROM bible_verses_legacy
                                 GROUP BY version,book,chapter,verse)""")
                self.db.execute("DROP TABLE bible_verses_legacy")
                self.db.commit()
        # Older imports normalized every translation to Indonesian book names.
        # Restore the CSV language for English datasets already in the database.
        with self.db:
            for version in ("KJV", "AMP", "NIV"):
                for local_name, english_name in zip(AYT_BOOK_NAMES, ENGLISH_BIBLE_BOOK_NAMES):
                    self.db.execute(
                        "DELETE FROM bible_verses WHERE version=? AND book=? AND EXISTS "
                        "(SELECT 1 FROM bible_verses AS english WHERE english.version=? "
                        "AND english.book=? AND english.chapter=bible_verses.chapter "
                        "AND english.verse=bible_verses.verse)",
                        (version, local_name, version, english_name))
                    self.db.execute("UPDATE bible_verses SET book=? WHERE version=? AND book=?",
                                    (english_name, version, local_name))
        self.db.execute("CREATE INDEX IF NOT EXISTS idx_bible_verses_version_reference "
                        "ON bible_verses(version,book,chapter,verse)")
        columns = {row["name"] for row in self.db.execute("PRAGMA table_info(schedule_items)")}
        if "media_path" not in columns:
            self.db.execute("ALTER TABLE schedule_items ADD COLUMN media_path TEXT DEFAULT ''")
        if "operator_note" not in columns:
            self.db.execute("ALTER TABLE schedule_items ADD COLUMN operator_note TEXT NOT NULL DEFAULT ''")
        if not self.db.execute("SELECT 1 FROM schedules LIMIT 1").fetchone():
            self.db.execute("INSERT INTO schedules(title) VALUES ('Ibadah Minggu')")
        self.db.commit()

    def get_setting(self, key, default=""):
        row = self.db.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.db.execute("INSERT OR REPLACE INTO app_settings(key, value) VALUES (?, ?)",
                        (key, str(value)))
        self.db.commit()

    def import_ayt_csv(self, path):
        return self.import_bible_csv(path, "AYT")

    def import_bible_csv(self, path, version_name="AYT"):
        version = re.sub(r"\s+", " ", str(version_name or "").strip()).upper()
        if not version or len(version) > 24:
            raise ValueError("Nama versi Alkitab harus diisi dan maksimal 24 karakter.")
        with open(path, "r", encoding="utf-8-sig", newline="") as source:
            reader = csv.DictReader(source)
            headers = {str(name or "").strip().casefold(): name for name in (reader.fieldnames or [])}
            aliases = {
                "book": ("book", "kitab", "book_name", "bookname"),
                "abbr": ("abbr", "abbreviation", "singkatan", "kode"),
                "chapter": ("chapter", "pasal", "chap"),
                "verse": ("verse", "ayat", "vers"),
                "text": ("text", "isi", "firman", "scripture", "scripture_text"),
                "title": ("title", "judul", "section_title"),
            }
            field_map = {}
            for target, names in aliases.items():
                found = next((headers[name] for name in names if name in headers), None)
                if found is not None:
                    field_map[target] = found
            # sabdacode's entire_kjv.csv and entire_amp.csv use header 0,1,2,3:
            # book, chapter, verse, text (no abbreviation column).
            if not {"book", "chapter", "verse", "text"}.issubset(field_map) and all(
                    str(column) in headers for column in ("0", "1", "2", "3")):
                field_map.update({"book": headers["0"], "chapter": headers["1"],
                                  "verse": headers["2"], "text": headers["3"]})
            required = {"book", "chapter", "verse", "text"}
            if not required.issubset(field_map):
                raise ValueError("CSV harus memuat kitab/book, pasal/chapter, ayat/verse, dan teks/text. "
                                 "Kolom abbr tidak wajib.")
            records = []
            for line_number, row in enumerate(reader, start=2):
                try:
                    book = normalize_bible_book(row.get(field_map["book"]), version)
                    chapter = int((row.get(field_map["chapter"]) or "").strip())
                    verse = int((row.get(field_map["verse"]) or "").strip())
                    text = re.sub(r"<t\s*/>", "", (row.get(field_map["text"]) or "")).strip()
                    if not book or not text:
                        continue
                    if chapter < 1 or verse < 1:
                        raise ValueError("nomor pasal dan ayat harus lebih dari nol")
                    abbr = (row.get(field_map.get("abbr"), "") or "").strip()
                    title = (row.get(field_map.get("title"), "") or "").strip()
                    records.append((version, book, abbr, chapter, verse, text, title))
                except (AttributeError, TypeError, ValueError) as error:
                    raise ValueError(f"Baris {line_number}: nomor pasal/ayat atau datanya tidak valid ({error}).") from error
        if not records:
            raise ValueError("Tidak ditemukan ayat yang dapat diimpor.")
        with self.db:
            self.db.executemany(
                "INSERT INTO bible_verses(version,book,abbr,chapter,verse,text,title) VALUES (?,?,?,?,?,?,?) "
                "ON CONFLICT(version,book,chapter,verse) DO UPDATE SET "
                "abbr=excluded.abbr,text=excluded.text,title=excluded.title",
                records)
        return len(records)

    def bible_versions(self):
        rows = self.db.execute("SELECT DISTINCT version FROM bible_verses ORDER BY version").fetchall()
        versions = [row["version"] for row in rows if row["version"]]
        return versions or ["AYT"]

    def bible_version_count(self, version):
        return self.db.execute("SELECT COUNT(*) FROM bible_verses WHERE version=?", (version,)).fetchone()[0]

    def search_bible_verses(self, query, limit=250, version="AYT"):
        term = query.strip()
        scope = "version=? AND " if version else ""
        scope_params = [version] if version else []
        if not term:
            return self.db.execute(
                f"SELECT * FROM bible_verses WHERE {scope}1=1 ORDER BY book, chapter, verse LIMIT ?",
                (*scope_params, limit)).fetchall()
        verse_range = re.fullmatch(r"(.+?)\s+(\d+):(\d+)\s*-\s*(?:(\d+):)?(\d+)", term)
        if verse_range:
            book, start_chapter, start_verse, end_chapter, end_verse = verse_range.groups()
            start_chapter, start_verse = int(start_chapter), int(start_verse)
            end_chapter = int(end_chapter) if end_chapter else start_chapter
            end_verse = int(end_verse)
            if (end_chapter, end_verse) < (start_chapter, start_verse):
                return []
            book_names = (normalize_bible_book(book, version),) if version else tuple(dict.fromkeys(
                (normalize_bible_book(book, "AYT"), normalize_bible_book(book, "KJV"))))
            book_clause = ("book=? COLLATE NOCASE" if len(book_names) == 1 else
                           f"book COLLATE NOCASE IN ({','.join('?' for _ in book_names)})")
            return self.db.execute(
                f"SELECT * FROM bible_verses WHERE {scope}{book_clause} AND "
                "(chapter>? OR (chapter=? AND verse>=?)) AND "
                "(chapter<? OR (chapter=? AND verse<=?)) ORDER BY chapter,verse LIMIT ?",
                (*scope_params, *book_names, start_chapter, start_chapter, start_verse,
                 end_chapter, end_chapter, end_verse, limit)).fetchall()
        reference = re.fullmatch(r"(.+?)\s+(\d+):(\d+)", term)
        if reference:
            book, chapter, verse = reference.groups()
            book_names = (normalize_bible_book(book, version),) if version else tuple(dict.fromkeys(
                (normalize_bible_book(book, "AYT"), normalize_bible_book(book, "KJV"))))
            book_clause = ("book=? COLLATE NOCASE" if len(book_names) == 1 else
                           f"book COLLATE NOCASE IN ({','.join('?' for _ in book_names)})")
            exact = self.db.execute(
                f"SELECT * FROM bible_verses WHERE {scope}({book_clause} OR abbr=? COLLATE NOCASE) "
                "AND chapter=? AND verse=? LIMIT ?",
                (*scope_params, *book_names, book.strip(), int(chapter), int(verse), limit)).fetchall()
            if exact:
                return exact
        pattern = f"%{term}%"
        return self.db.execute(
            f"SELECT * FROM bible_verses WHERE {scope}(book LIKE ? OR abbr LIKE ? OR title LIKE ? "
            "OR CAST(chapter AS TEXT) LIKE ? OR CAST(verse AS TEXT) LIKE ? OR text LIKE ?) "
            "ORDER BY book, chapter, verse LIMIT ?",
            (*scope_params, pattern, pattern, pattern, pattern, pattern, pattern, limit)).fetchall()

    def bible_books(self, version="AYT"):
        rows = self.db.execute("SELECT DISTINCT book FROM bible_verses WHERE version=?", (version,)).fetchall()
        names = [row["book"] for row in rows if row["book"]]
        order_names = ENGLISH_BIBLE_BOOK_NAMES if version in ("KJV", "AMP", "NIV") else AYT_BOOK_NAMES
        order = {name: index for index, name in enumerate(order_names)}
        return sorted(names, key=lambda name: (order.get(name, len(order)), name.casefold()))

    def bible_abbreviations(self, version="AYT"):
        rows = self.db.execute(
            "SELECT DISTINCT book, abbr FROM bible_verses WHERE version=? AND abbr <> ''", (version,)).fetchall()
        return {row["abbr"].casefold(): row["book"] for row in rows}

    def bible_chapters(self, book, version="AYT"):
        return [row[0] for row in self.db.execute(
            "SELECT DISTINCT chapter FROM bible_verses WHERE version=? AND book=? COLLATE NOCASE ORDER BY chapter",
            (version, book))]

    def bible_verse_numbers(self, book, chapter, version="AYT"):
        return [row[0] for row in self.db.execute(
            "SELECT DISTINCT verse FROM bible_verses WHERE version=? AND book=? COLLATE NOCASE AND chapter=? "
            "ORDER BY verse", (version, book, chapter))]

    def bible_chapter_verses(self, book, chapter, verse=None, version="AYT"):
        if verse is None:
            return self.db.execute(
                "SELECT * FROM bible_verses WHERE version=? AND book=? COLLATE NOCASE AND chapter=? ORDER BY verse",
                (version, book, chapter)).fetchall()
        return self.db.execute(
            "SELECT * FROM bible_verses WHERE version=? AND book=? COLLATE NOCASE AND chapter=? AND verse=?",
            (version, book, chapter, verse)).fetchall()

    def songs(self):
        return self.db.execute("SELECT * FROM songs ORDER BY title COLLATE NOCASE").fetchall()

    def search_resources(self, query, kind="Semua", limit=150):
        term = query.strip()
        if not term:
            return [], [], []
        pattern = f"%{term}%"
        songs = []
        verses = []
        media = []
        if kind in ("Semua", "Lagu"):
            songs = self.db.execute(
                "SELECT * FROM songs WHERE title LIKE ? OR author LIKE ? OR lyrics LIKE ? "
                "ORDER BY title COLLATE NOCASE LIMIT ?", (pattern, pattern, pattern, limit)).fetchall()
        if kind in ("Semua", "Ayat"):
            verses = self.search_bible_verses(term, limit=limit, version=None)
        if kind in ("Semua", "Media", "Gambar", "Video", "PowerPoint"):
            candidates = self.db.execute(
                "SELECT MIN(id) AS id, MIN(title) AS title, media_path FROM schedule_items "
                "WHERE media_path<>'' AND (title LIKE ? OR media_path LIKE ?) "
                "GROUP BY media_path ORDER BY title COLLATE NOCASE LIMIT ?",
                (pattern, pattern, limit * 4)).fetchall()
            categories = {
                ".pptx": "PowerPoint",
                ".png": "Gambar", ".jpg": "Gambar", ".jpeg": "Gambar", ".bmp": "Gambar",
                ".gif": "Gambar", ".webp": "Gambar",
                ".mp4": "Video", ".avi": "Video", ".mov": "Video", ".mkv": "Video",
                ".wmv": "Video", ".m4v": "Video",
            }
            for row in candidates:
                category = categories.get(Path(row["media_path"]).suffix.casefold(), "Media")
                if kind in ("Semua", "Media") or category == kind:
                    media.append({"id": row["id"], "title": row["title"],
                                  "media_path": row["media_path"], "category": category})
                    if len(media) >= limit:
                        break
        return songs, verses, media

    def save_song(self, song_id, title, author, lyrics):
        if song_id:
            self.db.execute("UPDATE songs SET title=?, author=?, lyrics=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                            (title, author, lyrics, song_id))
        else:
            self.db.execute("INSERT INTO songs(title,author,lyrics) VALUES (?,?,?)", (title, author, lyrics))
        self.db.commit()

    def schedule(self):
        return self.db.execute("SELECT * FROM schedules ORDER BY id DESC LIMIT 1").fetchone()

    def schedules(self):
        return self.db.execute("SELECT * FROM schedules ORDER BY id DESC").fetchall()

    def create_schedule(self, title):
        cur = self.db.execute("INSERT INTO schedules(title) VALUES (?)", (title,))
        self.db.commit()
        return cur.lastrowid

    def duplicate_schedule(self, schedule_id):
        source = self.db.execute("SELECT title FROM schedules WHERE id=?", (schedule_id,)).fetchone()
        if not source:
            return None
        cur = self.db.execute("INSERT INTO schedules(title) VALUES (?)", (f"{source['title']} (Salinan)",))
        new_id = cur.lastrowid
        self.db.execute("INSERT INTO schedule_items(schedule_id,kind,title,content,position,media_path,operator_note) "
                        "SELECT ?,kind,title,content,position,media_path,operator_note FROM schedule_items WHERE schedule_id=?",
                        (new_id, schedule_id))
        self.db.commit()
        return new_id

    def delete_schedule_by_id(self, schedule_id):
        """Menghapus data jadwal dari tabel SQLite berdasarkan ID."""
        self.db.execute("DELETE FROM schedules WHERE id=?", (schedule_id,))
        self.db.commit()

    def update_item(self, item_id, schedule_id, title, content, operator_note=""):
        self.db.execute("UPDATE schedule_items SET title=?,content=?,operator_note=? WHERE id=? AND schedule_id=?",
                        (title, content, operator_note, item_id, schedule_id))
        self.db.commit()

    def items(self, schedule_id):
        return self.db.execute("SELECT * FROM schedule_items WHERE schedule_id=? ORDER BY position,id", (schedule_id,)).fetchall()

    def add_item(self, schedule_id, kind, title, content):
        pos = self.db.execute("SELECT COALESCE(MAX(position),0)+1 FROM schedule_items WHERE schedule_id=?", (schedule_id,)).fetchone()[0]
        cursor = self.db.execute("INSERT INTO schedule_items(schedule_id,kind,title,content,position) VALUES (?,?,?,?,?)",
                                 (schedule_id, kind, title, content, pos))
        self.db.commit()
        return cursor.lastrowid

    def add_media(self, schedule_id, path):
        kind = "PowerPoint" if Path(path).suffix.casefold() == ".pptx" else "Media"
        cursor = self.db.execute("INSERT INTO schedule_items(schedule_id,kind,title,content,position,media_path) "
                                 "VALUES (?,?,?,?,(SELECT COALESCE(MAX(position),0)+1 FROM schedule_items WHERE schedule_id=?),?)",
                                 (schedule_id, kind, Path(path).name, "", schedule_id, str(path)))
        self.db.commit()
        return cursor.lastrowid

    def move_item(self, item_id, offset, schedule_id):
        items = self.items(schedule_id)
        index = next((i for i, row in enumerate(items) if row["id"] == item_id), None)
        target = index + offset if index is not None else -1
        if target < 0 or target >= len(items):
            return
        first, second = items[index], items[target]
        self.db.execute("UPDATE schedule_items SET position=? WHERE id=?", (second["position"], first["id"]))
        self.db.execute("UPDATE schedule_items SET position=? WHERE id=?", (first["position"], second["id"]))
        self.db.commit()


def fit_text_font(widget, text, max_size, min_size, width, height, family="Segoe UI",
                  weight="bold", padding=10):
    """Choose the largest font that fits wrapped text inside a pixel area."""
    width = max(40, int(width) - padding * 2)
    height = max(30, int(height) - padding * 2)
    paragraphs = str(text).splitlines() or [""]
    for size in range(int(max_size), int(min_size) - 1, -1):
        measure = tkfont.Font(root=widget, family=family, size=size, weight=weight)
        line_count = 0
        for paragraph in paragraphs:
            words = paragraph.split()
            if not words:
                line_count += 1
                continue
            line = ""
            for word in words:
                candidate = f"{line} {word}" if line else word
                if line and measure.measure(candidate) > width:
                    line_count += 1
                    line = word
                else:
                    line = candidate
                word_width = measure.measure(line)
                if word_width > width:
                    line_count += max(0, (word_width - 1) // width)
                    line = word
            line_count += 1
        if line_count * measure.metrics("linespace") * 1.12 <= height:
            return size
    return int(min_size)


_YOUTUBE_ID_RE = re.compile(r"^[A-Za-z0-9_-]{8,}$")


def _youtube_start_seconds(value):
    """Convert '90', '1m30s', or '1h2m3s' to seconds."""
    value = str(value or "").strip().casefold()
    if not value:
        return 0
    if value.isdigit():
        return int(value)
    match = re.fullmatch(r"(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?", value)
    if not match:
        return 0
    hours, minutes, seconds = (int(part) if part else 0 for part in match.groups())
    return hours * 3600 + minutes * 60 + seconds


def youtube_embed_url(url):
    """Return an autoplay YouTube embed URL for a YouTube link, or None.

    Plain watch links still show the YouTube page (recommendations, comments,
    cookie banners); the embed player can fill the projector screen and start
    by itself, which is what a worship presentation needs.
    """
    try:
        parsed = urlparse(str(url or "").strip())
    except ValueError:
        return None
    host = (parsed.netloc or "").casefold().split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    if host not in {"youtube.com", "m.youtube.com", "music.youtube.com",
                    "youtu.be", "youtube-nocookie.com"}:
        return None
    query = parse_qs(parsed.query)
    playlist_id = (query.get("list") or [""])[0]
    video_id = ""
    path = parsed.path or ""
    if host == "youtu.be":
        video_id = path.strip("/").split("/")[0]
    elif path == "/watch":
        video_id = (query.get("v") or [""])[0]
    elif path.startswith(("/embed/", "/shorts/", "/live/", "/v/")):
        parts = path.strip("/").split("/")
        video_id = parts[1] if len(parts) > 1 else ""
    elif path not in {"/playlist", "/playlist/"}:
        return None
    params = {"autoplay": "1", "rel": "0", "modestbranding": "1", "playsinline": "1"}
    if video_id and _YOUTUBE_ID_RE.match(video_id):
        base = f"https://www.youtube.com/embed/{video_id}"
        if playlist_id:
            params["list"] = playlist_id
    elif playlist_id:
        base = "https://www.youtube.com/embed/videoseries"
        params["list"] = playlist_id
    else:
        return None
    start = (query.get("t") or query.get("start") or [""])[0]
    seconds = _youtube_start_seconds(start)
    if seconds:
        params["start"] = str(seconds)
    return f"{base}?{urlencode(params)}"


def youtube_player_html(embed_url):
    """A local page that stretches the embed player to the whole window."""
    return (
        "<!doctype html><html><head><meta charset=\"utf-8\">"
        "<title>Sahabat Ibadah</title><style>"
        "html,body{margin:0;height:100%;background:#000;overflow:hidden}"
        "iframe{position:absolute;inset:0;width:100%;height:100%;border:0}"
        "</style></head><body>"
        f"<iframe src=\"{html_escape(embed_url, quote=True)}\" "
        "allow=\"autoplay; encrypted-media; fullscreen; picture-in-picture\" "
        "allowfullscreen referrerpolicy=\"strict-origin-when-cross-origin\"></iframe>"
        "</body></html>"
    )


class _YouTubePlayerServer:
    """Serve the wrapper page from loopback so YouTube receives a referrer.

    YouTube refuses embeds opened from a file:// page ("Error 153") because the
    request carries no referrer. A page served from http://127.0.0.1 gives the
    player the origin it expects, and the server never leaves this computer.
    """

    def __init__(self):
        self._pages = {}
        self._server = None
        self._lock = threading.Lock()

    def url_for(self, page_html):
        try:
            self._start()
        except OSError:
            return None
        token = secrets.token_urlsafe(12)
        self._pages[token] = page_html
        if len(self._pages) > 40:
            for key in list(self._pages)[:-20]:
                self._pages.pop(key, None)
        port = self._server.server_address[1]
        return f"http://127.0.0.1:{port}/player/{token}"

    def _start(self):
        with self._lock:
            if self._server is not None:
                return
            pages = self._pages

            class Handler(http.server.BaseHTTPRequestHandler):
                def do_GET(self):
                    body = pages.get(self.path.rsplit("/", 1)[-1])
                    if body is None:
                        self.send_error(404)
                        return
                    payload = body.encode("utf-8")
                    self.send_response(200)
                    self.send_header("Content-Type", "text/html; charset=utf-8")
                    self.send_header("Content-Length", str(len(payload)))
                    self.send_header("Cache-Control", "no-store")
                    self.end_headers()
                    self.wfile.write(payload)

                def log_message(self, *args):
                    pass

            server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
            server.daemon_threads = True
            threading.Thread(target=server.serve_forever, daemon=True).start()
            self._server = server


YOUTUBE_PLAYER_SERVER = _YouTubePlayerServer()


def youtube_player_page_url(page_html):
    """Return a loopback URL for the wrapper page, or None if it cannot run."""
    return YOUTUBE_PLAYER_SERVER.url_for(page_html)


def _contrasting_text_color(hex_color):
    """Pick black or white text so a colour swatch label stays readable."""
    try:
        value = str(hex_color).lstrip("#")
        red, green, blue = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
    except (ValueError, IndexError):
        return "#111111"
    return "#111111" if (red * 299 + green * 587 + blue * 114) / 1000 > 150 else "#ffffff"


class TextStrokeDialog(tk.Toplevel):
    """Switch the text outline on/off, choose its colour and thickness."""

    def __init__(self, parent, title, text_color, enabled, color, width, on_change):
        super().__init__(parent)
        self.title(title)
        self.configure(bg="#f2f5f8")
        self.resizable(False, False)
        self.transient(parent)
        self.on_change = on_change
        self.text_color = text_color
        self.enabled_var = tk.BooleanVar(value=bool(enabled))
        self.color = color or DEFAULT_STROKE_COLOR
        self.width_var = tk.DoubleVar(value=max(1, min(10, int(width or DEFAULT_STROKE_WIDTH))))
        self.width_value = max(1, min(10, int(width or DEFAULT_STROKE_WIDTH)))
        self._last_state = None
        self._emit_after_id = None

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Garis tepi (stroke) teks", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(body, text="Garis tepi membuat teks tetap terbaca di atas video, gambar, "
                             "atau latar yang warnanya mirip dengan teks.",
                  style="Sub.TLabel", wraplength=430, justify="left").pack(anchor="w", pady=(2, 12))
        ttk.Checkbutton(body, text="Aktifkan garis tepi", variable=self.enabled_var,
                        command=self._changed).pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(12, 0))
        ttk.Label(row, text="Warna garis:").pack(side="left")
        self.color_button = tk.Button(row, text="", width=13, command=self._choose_color,
                                      relief="solid", bd=1, cursor="hand2")
        self.color_button.pack(side="left", padx=(8, 18))
        ttk.Label(row, text="Ketebalan:").pack(side="left")
        self.width_scale = ttk.Scale(row, from_=1, to=10, variable=self.width_var,
                                     command=lambda _value: self._changed())
        self.width_scale.pack(side="left", padx=(8, 6), fill="x", expand=True)
        self.width_label = ttk.Label(row, text="", width=5)
        self.width_label.pack(side="left")
        self.sample = tk.Canvas(body, width=440, height=130, bg="#26313b",
                                highlightthickness=1, highlightbackground="#b7c3cf")
        self.sample.pack(fill="x", pady=(14, 6))
        self.sample.bind("<Configure>", lambda _event: self._paint_sample())
        self.sample_hint = ttk.Label(body, text="", foreground="#a43b42", wraplength=430)
        self.sample_hint.pack(fill="x")
        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="Selesai", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Destroy>", self._on_destroy)
        self._sync_color_button()
        self.width_label.config(text=f"{self.width_value} px")
        self._paint_sample()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx() + 140}+{parent.winfo_rooty() + 110}")

    def _choose_color(self):
        chosen = colorchooser.askcolor(color=self.color, title="Warna garis tepi teks", parent=self)
        if chosen and chosen[1]:
            self.color = chosen[1]
            self._sync_color_button()
            self._changed()

    def _on_destroy(self, event):
        if event.widget is not self or not self._emit_after_id:
            return
        try:
            self.after_cancel(self._emit_after_id)
        except tk.TclError:
            pass
        self._emit_after_id = None

    def _sync_color_button(self):
        text_color = _contrasting_text_color(self.color)
        self.color_button.config(bg=self.color, activebackground=self.color, text=self.color.upper(),
                                 fg=text_color, activeforeground=text_color,
                                 disabledforeground=text_color)

    def _paint_sample(self):
        canvas = self.sample
        width = max(120, canvas.winfo_width())
        height = max(60, canvas.winfo_height())
        canvas.delete("all")
        canvas.create_rectangle(0, 0, width, height, fill="#26313b", outline="")
        canvas.create_rectangle(0, height * 0.58, width, height, fill="#5d6a75", outline="")
        stroke_color = self.color if self.enabled_var.get() else None
        draw_text_with_stroke(canvas, width // 2, height // 2, "Yohanes 3:16",
                              ("Segoe UI", 20, "bold"), self.text_color,
                              stroke_color=stroke_color, stroke_width=self.width_value,
                              wrap_width=max(60, width - 20))
        hint = getattr(self, "sample_hint", None)
        if hint is None:
            return
        if self.enabled_var.get() and str(self.color).casefold() == str(self.text_color).casefold():
            hint.config(text="Warna garis sama dengan warna teks. Pilih warna lain agar garis tepi terlihat.")
        else:
            hint.config(text="")

    def _changed(self, _value=None):
        self.width_value = max(1, min(10, int(round(self.width_var.get()))))
        self.width_label.config(text=f"{self.width_value} px")
        self._paint_sample()
        # The scale fires continuously while dragging; apply the change once it settles.
        if self._emit_after_id:
            try:
                self.after_cancel(self._emit_after_id)
            except tk.TclError:
                pass
        self._emit_after_id = self.after(80, self._emit)

    def _emit(self):
        self._emit_after_id = None
        state = {"enabled": bool(self.enabled_var.get()), "color": self.color,
                 "width": self.width_value}
        if state != self._last_state:
            self._last_state = state
            self.on_change(state)


class TextLayoutDialog(tk.Toplevel):
    """Position, margins and size of the live overlay text."""

    def __init__(self, parent, text_color, stroke_color, stroke_width, state, on_change):
        super().__init__(parent)
        self.title("Posisi & Margin Teks Live")
        self.configure(bg="#f2f5f8")
        self.resizable(False, False)
        self.transient(parent)
        self.on_change = on_change
        self.text_color = text_color
        self.stroke_color = stroke_color
        self.stroke_width = int(stroke_width or 0)
        self._emit_after_id = None
        self._last_state = None
        self._pending_state = None
        self.offset_x = float(state.get("offset_x", 0))
        self.offset_y = float(state.get("offset_y", 0))
        self._drag = None
        self._preview_layout = None
        self._preview_size = (490, 190)

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Posisi & margin teks live",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w")
        ttk.Label(body, text="Garis margin dan teks di pratinjau bisa digeser langsung dengan mouse. "
                             "Margin dihitung dalam persen dari lebar/tinggi layar proyeksi.",
                  style="Sub.TLabel", wraplength=470, justify="left").pack(anchor="w", pady=(2, 12))

        row = ttk.Frame(body)
        row.pack(fill="x")
        ttk.Label(row, text="Posisi:").pack(side="left")
        self.position_choice = ttk.Combobox(row, state="readonly",
                                            values=list(OVERLAY_POSITIONS), width=10)
        self.position_choice.set(state["position"])
        self.position_choice.pack(side="left", padx=(6, 18))
        self.position_choice.bind("<<ComboboxSelected>>", lambda _event: self._changed())
        ttk.Label(row, text="Perataan:").pack(side="left")
        self.align_choice = ttk.Combobox(row, state="readonly",
                                         values=list(OVERLAY_ALIGNS), width=10)
        self.align_choice.set(state["align"])
        self.align_choice.pack(side="left", padx=(6, 0))
        self.align_choice.bind("<<ComboboxSelected>>", lambda _event: self._changed())

        margin = ttk.Frame(body)
        margin.pack(fill="x", pady=(14, 0))
        ttk.Label(margin, text="Margin (%):").pack(side="left")
        self.margin_vars = {}
        for label, key in (("Kiri", "left"), ("Kanan", "right"),
                           ("Atas", "top"), ("Bawah", "bottom")):
            cell = ttk.Frame(margin)
            cell.pack(side="left", padx=(10, 0))
            ttk.Label(cell, text=label).pack(side="left")
            variable = tk.StringVar(value=str(state[key]))
            spin = tk.Spinbox(cell, from_=0, to=45, width=4, textvariable=variable,
                              command=self._changed, relief="solid", bd=1)
            spin.pack(side="left", padx=(4, 0))
            spin.bind("<KeyRelease>", lambda _event: self._changed())
            spin.bind("<FocusOut>", lambda _event: self._changed())
            self.margin_vars[key] = variable

        scale_row = ttk.Frame(body)
        scale_row.pack(fill="x", pady=(16, 0))
        ttk.Label(scale_row, text="Ukuran teks:").pack(side="left")
        self.scale_var = tk.DoubleVar(value=state["scale"])
        self.scale = ttk.Scale(scale_row, from_=40, to=150, variable=self.scale_var,
                               command=lambda _value: self._changed())
        self.scale.pack(side="left", fill="x", expand=True, padx=(8, 6))
        self.scale_label = ttk.Label(scale_row, text=f"{state['scale']}%", width=6)
        self.scale_label.pack(side="left")

        self.preview = tk.Canvas(body, width=490, height=190, bg="#26313b",
                                 highlightthickness=1, highlightbackground="#b7c3cf")
        self.preview.pack(fill="x", pady=(14, 6))
        self.preview.bind("<Configure>", lambda _event: self._paint_preview())
        self.preview.bind("<ButtonPress-1>", self._preview_press)
        self.preview.bind("<B1-Motion>", self._preview_drag)
        self.preview.bind("<ButtonRelease-1>", lambda _event: self._end_drag())
        self.preview.bind("<Motion>", self._preview_hover)

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(8, 0))
        self.offset_label = ttk.Label(actions, text="", style="Sub.TLabel")
        self.offset_label.pack(side="left")
        ttk.Button(actions, text="Reset geseran", command=self._reset_offset).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Selesai", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Destroy>", self._on_destroy)
        self.update_idletasks()
        self._paint_preview()
        self.geometry(f"+{parent.winfo_rootx() + 120}+{parent.winfo_rooty() + 70}")

    def _state(self):
        def margin(key):
            try:
                return max(0, min(45, int(float(self.margin_vars[key].get()))))
            except (TypeError, ValueError):
                return 6

        return {"position": self.position_choice.get(),
                "align": self.align_choice.get(),
                "scale": max(40, min(150, int(round(self.scale_var.get())))),
                "left": margin("left"), "right": margin("right"),
                "top": margin("top"), "bottom": margin("bottom"),
                "offset_x": round(self.offset_x, 1), "offset_y": round(self.offset_y, 1)}

    def _paint_preview(self, state=None):
        canvas = self.preview
        state = state or self._state()
        width = max(140, canvas.winfo_width() or 490)
        height = max(90, canvas.winfo_height() or 190)
        self._preview_size = (width, height)
        canvas.delete("all")
        canvas.create_rectangle(0, 0, width, height, fill="#26313b", outline="")
        canvas.create_rectangle(0, height * 0.58, width, height, fill="#5d6a75", outline="")
        sample = "Yohanes 3:16\nKarena begitu besar kasih Allah"
        layout = overlay_text_layout(canvas, sample, width, height, state)
        self._preview_layout = layout
        for x in (layout["left"], layout["right"]):
            canvas.create_line(x, 0, x, height, fill="#8fa0ad", dash=(3, 5))
        for y in (layout["top"], layout["bottom"]):
            canvas.create_line(0, y, width, y, fill="#8fa0ad", dash=(3, 5))
        draw_text_with_stroke(canvas, layout["x"], layout["y"], sample,
                              ("Segoe UI", layout["font_size"], "bold"), self.text_color,
                              stroke_color=self.stroke_color, stroke_width=self.stroke_width,
                              wrap_width=layout["wrap"], justify=layout["justify"],
                              anchor=layout["anchor"], tags=("preview-text",))
        if hasattr(self, "offset_label"):
            self.offset_label.config(
                text=f"Geseran: {self.offset_x:+.0f}% / {self.offset_y:+.0f}%")

    def _hit_test(self, x, y):
        layout = self._preview_layout
        if layout is None:
            return None, None
        bounds = self.preview.bbox("preview-text")
        if bounds and bounds[0] - 4 <= x <= bounds[2] + 4 and bounds[1] - 4 <= y <= bounds[3] + 4:
            return "text", None
        tolerance = 7
        for key in ("left", "right"):
            if abs(x - layout[key]) <= tolerance:
                return "margin", key
        for key in ("top", "bottom"):
            if abs(y - layout[key]) <= tolerance:
                return "margin", key
        return None, None

    def _set_cursor(self, kind, key):
        if kind == "margin":
            cursor = "sb_h_double_arrow" if key in ("left", "right") else "sb_v_double_arrow"
        elif kind == "text":
            cursor = "fleur"
        else:
            cursor = ""
        self.preview.config(cursor=cursor)

    def _preview_hover(self, event):
        if self._drag is None:
            self._set_cursor(*self._hit_test(event.x, event.y))

    def _preview_press(self, event):
        kind, key = self._hit_test(event.x, event.y)
        if kind is None:
            return
        self._drag = {"kind": kind, "key": key, "x": event.x, "y": event.y,
                      "offset_x": self.offset_x, "offset_y": self.offset_y}
        self._set_cursor(kind, key)

    def _preview_drag(self, event):
        drag = self._drag
        if drag is None:
            return
        width = max(1, self._preview_size[0])
        height = max(1, self._preview_size[1])
        if drag["kind"] == "margin":
            key = drag["key"]
            if key in ("left", "right"):
                percent = event.x / width * 100
                if key == "right":
                    percent = 100 - percent
            else:
                percent = event.y / height * 100
                if key == "bottom":
                    percent = 100 - percent
            self.margin_vars[key].set(str(max(0, min(45, int(round(percent))))))
        else:
            dx = (event.x - drag["x"]) / width * 100
            dy = (event.y - drag["y"]) / height * 100
            self.offset_x = round(max(-80.0, min(80.0, drag["offset_x"] + dx)), 1)
            self.offset_y = round(max(-80.0, min(80.0, drag["offset_y"] + dy)), 1)
        self._changed()

    def _end_drag(self):
        self._drag = None
        self.preview.config(cursor="")

    def _reset_offset(self):
        self.offset_x = 0.0
        self.offset_y = 0.0
        self._changed()

    def _changed(self, _value=None):
        self._pending_state = self._state()
        self.scale_label.config(text=f"{self._pending_state['scale']}%")
        self._paint_preview(self._pending_state)
        if self._emit_after_id:
            try:
                self.after_cancel(self._emit_after_id)
            except tk.TclError:
                pass
        self._emit_after_id = self.after(80, self._emit)

    def _emit(self):
        self._emit_after_id = None
        state = self._pending_state
        if state and state != self._last_state:
            self._last_state = state
            self.on_change(state)

    def _on_destroy(self, event):
        if event.widget is not self or not self._emit_after_id:
            return
        try:
            self.after_cancel(self._emit_after_id)
        except tk.TclError:
            pass
        self._emit_after_id = None


class SahabatIbadahApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.store = Store()
        self.dark_mode = self.store.get_setting("dark_mode", "0") == "1"
        try:
            self.palette_overrides = json.loads(self.store.get_setting("ui_palette", "{}"))
            if not isinstance(self.palette_overrides, dict):
                self.palette_overrides = {}
        except (TypeError, json.JSONDecodeError):
            self.palette_overrides = {}
        self.title("Sahabat Ibadah — Presentasi Ibadah")
        self.geometry("1120x720")
        self.minsize(850, 560)
        self.configure(bg="#f2f5f8")
        self.current_song_id = None
        self.current_slides = []
        self.slide_index = 0
        self.preview_slides = []
        self.preview_slide_index = 0
        self.preview_item_id = None
        self.projector_window = None
        self.presentation_background_path = None
        self.presentation_background_color = "#000000"
        self.presentation_text_color = "#ffffff"
        self.presentation_background_loop = True
        self.transition_style = "Potong"
        self.youtube_fullscreen = True
        self.presentation_stroke_enabled = False
        self.presentation_stroke_color = DEFAULT_STROKE_COLOR
        self.presentation_stroke_width = DEFAULT_STROKE_WIDTH
        self.live_text_color = "#ffffff"
        self.live_stroke_enabled = False
        self.live_stroke_color = DEFAULT_STROKE_COLOR
        self.live_stroke_width = DEFAULT_STROKE_WIDTH
        self.live_text_background = self.store.get_setting("live_text_background", "0") == "1"
        self.broadcast_capture = None
        self.broadcast_second_capture = None
        self.broadcast_thread = None
        self.broadcast_stop_event = None
        self.broadcast_mic_stream = None
        self.broadcast_mic_devices = {}
        self.broadcast_camera_stopped = False
        self.broadcast_content_mode = False
        self.broadcast_camera_enabled = False
        self.broadcast_output_mode = "primary"
        self.broadcast_primary_source = None
        self.broadcast_secondary_source = None
        self._broadcast_camera_switch_generation = 0
        self.broadcast_content_frames = []
        self.broadcast_content_index = 0
        self.broadcast_content_capture = None
        self.broadcast_content_rows = []
        self.broadcast_content_frame = None
        self._closing = False
        self._camera_scan_in_progress = False
        self.live_source_current = None
        self.live_second_source_current = None
        self._load_live_layout_settings()
        self._presentation_stroke_dialog = None
        self._live_stroke_dialog = None
        self._live_layout_dialog = None
        self._style()
        self._build()
        self._apply_widget_theme()
        self.protocol("WM_DELETE_WINDOW", self._close_main_window)
        self.bind_all("<Map>", self._on_theme_window_map, add="+")
        self.refresh_songs()
        self.refresh_schedule()

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        self.colors = ({
            # Keep the requested black canvas, with navy surfaces for clear hierarchy.
            "app": "#000000", "card": "#0f172a", "panel": "#111c2e",
            "field": "#1e293b", "field_text": "#ffffff", "text": "#ffffff",
            "muted": "#ffffff", "heading": "#ffffff", "line": "#273449",
            "tab_bar": "#080f1d", "tab": "#111c2e", "tab_text": "#dbeafe",
            "button": "#1e293b", "button_hover": "#334155", "button_text": "#ffffff",
            "nav": "#0b1324", "nav_count": "#1e293b", "nav_text": "#ffffff",
            "tree_head": "#162238", "tree_select": "#1e3a8a", "tree_select_text": "#ffffff",
            "accent": "#3b82f6", "accent_hover": "#2563eb", "danger": "#e11d48",
            "scroll": "#1e293b", "border": "#273449",
        } if self.dark_mode else {
            "app": "#f1f5f9", "card": "#ffffff", "panel": "#f8fafc",
            "field": "#ffffff", "field_text": "#0f172a", "text": "#0f172a",
            "muted": "#475569", "heading": "#0f172a", "line": "#cbd5e1",
            "tab_bar": "#e2e8f0", "tab": "#f1f5f9", "tab_text": "#334155",
            "button": "#e2e8f0", "button_hover": "#cbd5e1", "button_text": "#1e293b",
            "nav": "#f8fafc", "nav_count": "#e2e8f0", "nav_text": "#0f172a",
            "tree_head": "#e2e8f0", "tree_select": "#dbeafe", "tree_select_text": "#0f172a",
            "accent": "#2563eb", "accent_hover": "#1d4ed8", "danger": "#e11d48",
            "scroll": "#e2e8f0", "border": "#cbd5e1",
        })
        mode_key = "dark" if self.dark_mode else "light"
        overrides = self.palette_overrides.get(mode_key, {})
        if isinstance(overrides, dict):
            for role, value in overrides.items():
                if role in self.colors and re.fullmatch(r"#[0-9a-fA-F]{6}", str(value)):
                    self.colors[role] = str(value).upper()
        c = self.colors
        self.configure(bg=c["app"])
        style.configure("TFrame", background=c["app"])
        style.configure("Card.TFrame", background=c["card"])
        style.configure("TLabel", background=c["app"], foreground=c["text"], font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"), foreground=c["heading"])
        style.configure("Sub.TLabel", foreground=c["muted"])
        style.configure("TNotebook", background=c["tab_bar"], borderwidth=0, tabmargins=(4, 4, 4, 0))
        style.configure("TNotebook.Tab", background=c["tab"], foreground=c["tab_text"],
                        padding=(16, 9), font=("Segoe UI", 10))
        style.map("TNotebook.Tab", background=[("selected", c["card"]), ("active", c["button_hover"])],
                  foreground=[("selected", c["heading"])])
        style.configure("TButton", background=c["button"], foreground=c["button_text"],
                        padding=(12, 8), font=("Segoe UI", 10), borderwidth=1)
        style.map("TButton", background=[("active", c["button_hover"]), ("disabled", c["panel"])],
                  foreground=[("disabled", c["muted"])])
        style.configure("TEntry", fieldbackground=c["field"], foreground=c["field_text"],
                        padding=(7, 6))
        style.configure("TCombobox", fieldbackground=c["field"], background=c["button"],
                        foreground=c["field_text"], arrowcolor=c["button_text"], padding=(6, 5))
        style.map("TCombobox", fieldbackground=[("readonly", c["field"])],
                  foreground=[("readonly", c["field_text"])])
        style.configure("TCheckbutton", background=c["card"], foreground=c["text"])
        style.map("TCheckbutton", background=[("active", c["card"])],
                  foreground=[("disabled", c["muted"])])
        style.configure("TLabelframe", background=c["card"], bordercolor=c["border"])
        style.configure("TLabelframe.Label", background=c["card"], foreground=c["heading"])
        style.configure("TPanedwindow", background=c["app"])
        style.configure("TSeparator", background=c["line"])
        style.configure("TScrollbar", background=c["scroll"], troughcolor=c["app"],
                        arrowcolor=c["button_text"])
        style.configure("Accent.TButton", background=c["accent"], foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", c["accent_hover"])])
        style.configure("LiveNext.TButton", background=c["accent"], foreground="white",
                        padding=(18, 12), font=("Segoe UI", 11, "bold"))
        style.map("LiveNext.TButton", background=[("active", c["accent_hover"]),
                                                   ("disabled", c["panel"])],
                  foreground=[("disabled", c["muted"])])
        style.configure("LivePrevious.TButton", background=c["button"], foreground=c["button_text"],
                        padding=(16, 12), font=("Segoe UI", 11, "bold"))
        style.map("LivePrevious.TButton", background=[("active", c["button_hover"]),
                                                       ("disabled", c["panel"])],
                  foreground=[("disabled", c["muted"])])
        style.configure("LiveNav.TLabelframe", background=c["nav"], bordercolor=c["border"],
                        relief="solid", borderwidth=1)
        style.configure("LiveNav.TLabelframe.Label", background=c["nav"],
                        foreground=c["heading"], font=("Segoe UI", 10, "bold"))
        style.configure("LiveNavInner.TFrame", background=c["nav"])
        style.configure("LiveCount.TLabel", background=c["nav_count"], foreground=c["nav_text"],
                        font=("Segoe UI", 11, "bold"), padding=(12, 9))
        style.configure("Danger.TButton", background=c["danger"], foreground="white", font=("Segoe UI", 10, "bold"))
        style.map("Danger.TButton", background=[("active", "#842d33")])
        style.configure("Treeview", rowheight=32, font=("Segoe UI", 10), background=c["field"],
                        fieldbackground=c["field"], foreground=c["field_text"])
        style.map("Treeview", background=[("selected", c["tree_select"])],
                  foreground=[("selected", c["tree_select_text"])])
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"),
                        background=c["tree_head"], foreground=c["text"])
        if self.dark_mode:
            # The clam theme uses bright bevel edges by default; soften them for dark UI.
            for widget_style in ("TButton", "Accent.TButton", "Danger.TButton", "LiveNext.TButton",
                                 "LivePrevious.TButton", "TEntry", "TCombobox", "Treeview",
                                 "Treeview.Heading", "TNotebook", "TNotebook.Tab", "TLabelframe",
                                 "LiveNav.TLabelframe"):
                style.configure(widget_style, bordercolor=c["border"], lightcolor=c["border"],
                                darkcolor=c["app"], focuscolor=c["border"])

    def _apply_widget_theme(self, root=None):
        """Update classic Tk widgets that do not follow ttk styles automatically."""
        root = root or self
        c = self.colors
        surfaces = {
            "white": "card", "#ffffff": "card", "#192431": "card", "#2f2f2f": "card",
            "#f2f5f8": "app", "#111923": "app", "#07090b": "app", "#212121": "app",
            "#f4f7fa": "nav", "#202d3a": "nav", "#2b2b2b": "nav",
            "#0b0b0b": "nav", "#f7f9fb": "card", "#e4edf5": "nav_count", "#314456": "nav_count", "#393939": "nav_count",
            "#1a1a1a": "nav_count", "#e8edf2": "button", "#293847": "button", "#3a3a3a": "button",
            "#1c1c1c": "button",
            "#101010": "card", "#111111": "field", "#050505": "tab_bar", "#151515": "tab",
            "#0f172a": "card", "#111c2e": "panel", "#1e293b": "field",
            "#273449": "line", "#080f1d": "tab_bar", "#0b1324": "nav",
            "#162238": "tree_head", "#334155": "button_hover", "#3b82f6": "accent",
            "#2563eb": "accent", "#e11d48": "danger",
            "#dfe6ee": "tab_bar",
        }
        text_colors = {
            "black": "text", "#000000": "text", "#1f2937": "text", "#12243a": "heading",
            "#203246": "button_text", "#172b3a": "field_text", "#425466": "tab_text",
            "#17324d": "nav_text", "#65758b": "muted", "#64748b": "muted",
            "#94a3b8": "muted", "#e7edf4": "text", "#f2f6fa": "heading",
            "#9aaabd": "muted", "#c4cfda": "tab_text",
            "#dedede": "text", "#e2e2e2": "field_text", "#e7e7e7": "heading",
            "#a3a3a3": "muted", "#c2c2c2": "tab_text", "#e0e0e0": "button_text",
            "#c8c8c8": "text", "#c9c9c9": "field_text", "#d4d4d4": "heading",
            "#858585": "muted", "#aaaaaa": "tab_text", "#c6c6c6": "button_text",
            "#65758b": "muted", "#64748b": "muted", "#176b87": "text",
            "#b44b4b": "text", "#a43b42": "text", "#8da0b3": "text",
        }
        for widget in root.winfo_children():
            try:
                klass = widget.winfo_class()
                options = set(widget.keys())
                supports_custom_surface = klass in {
                    "Toplevel", "Label", "TLabel", "Frame", "Canvas", "Text", "Entry", "Button",
                    "Menubutton", "Checkbutton", "Radiobutton", "Listbox", "Spinbox",
                }
                if supports_custom_surface and "background" in options:
                    current = str(widget.cget("background")).casefold()
                    if current in surfaces:
                        widget.configure(background=c[surfaces[current]])
                elif supports_custom_surface and "bg" in options:
                    current = str(widget.cget("bg")).casefold()
                    if current in surfaces:
                        widget.configure(bg=c[surfaces[current]])
                for option in ("foreground", "fg"):
                    if option in options:
                        current = str(widget.cget(option)).casefold()
                        if current in text_colors and (supports_custom_surface or klass == "TCheckbutton"):
                            widget.configure(**{option: c[text_colors[current]]})
                if klass == "Text":
                    # Classic Tk Text widgets do not inherit ttk colors. Give them
                    # the same dark, softly contrasted surface as ChatGPT's composer.
                    if self.dark_mode:
                        widget.configure(background=c["field"], foreground="#ffffff",
                                         insertbackground="#ffffff", relief="flat", bd=0,
                                         highlightthickness=1, highlightbackground=c["line"],
                                         highlightcolor=c["accent"], padx=10, pady=8)
                    else:
                        widget.configure(background="#ffffff", foreground=c["field_text"],
                                         insertbackground=c["field_text"], relief="solid", bd=1,
                                         highlightthickness=0)
                    widget.configure(selectbackground=c["accent"], selectforeground="#ffffff")
                if klass == "Canvas":
                    widget.configure(highlightbackground=c["line"])
            except (tk.TclError, KeyError):
                pass
            self._apply_widget_theme(widget)

    def _on_theme_window_map(self, event=None):
        if self.dark_mode:
            window = getattr(event, "widget", self)
            if window is self or isinstance(window, tk.Toplevel):
                self.after_idle(lambda: self._apply_widget_theme(window))

    def toggle_dark_mode(self):
        self.dark_mode = not self.dark_mode
        self.store.set_setting("dark_mode", "1" if self.dark_mode else "0")
        self._style()
        self._apply_widget_theme()
        self.theme_button.config(text="☀  Mode Terang" if self.dark_mode else "☾  Mode Gelap")

    def open_color_palette(self):
        """Let operators customize and persist the application's light/dark colors."""
        dialog = tk.Toplevel(self)
        dialog.title("List Warna — Sahabat Ibadah")
        dialog.transient(self)
        dialog.configure(bg=self.colors["app"])
        dialog.geometry("680x700")
        dialog.minsize(600, 620)

        definitions = (
            ("app", "Latar utama"),
            ("card", "Panel / kartu"),
            ("field", "Kolom input"),
            ("accent", "Aksen utama"),
            ("button", "Tombol"),
            ("danger", "Blackout / peringatan"),
        )
        color_buttons = {}

        def save_palette():
            self.store.set_setting("ui_palette", json.dumps(self.palette_overrides))

        def apply_color(mode, role, button):
            current = self.palette_overrides.get(mode, {}).get(role)
            if not current:
                current = (self.colors if (mode == "dark") == self.dark_mode else
                           (self.light_defaults if mode == "light" else self.dark_defaults))[role]
            chosen = colorchooser.askcolor(color=current, title=f"Pilih warna: {role}", parent=dialog)[1]
            if not chosen:
                return
            self.palette_overrides.setdefault(mode, {})[role] = chosen.upper()
            save_palette()
            button.configure(text=chosen.upper(), bg=chosen, activebackground=chosen,
                             fg=_contrasting_text_color(chosen), activeforeground=_contrasting_text_color(chosen))
            if mode == ("dark" if self.dark_mode else "light"):
                self._style()
                self._apply_widget_theme()
                dialog.configure(bg=self.colors["app"])

        header = ttk.Frame(dialog, padding=(18, 14))
        header.pack(fill="x")
        ttk.Label(header, text="Palet Warna", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text="Pilih warna untuk tiap mode. Perubahan disimpan otomatis.",
                  style="Sub.TLabel").pack(anchor="w", pady=(4, 0))
        notebook = ttk.Notebook(dialog)
        notebook.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        ttk.Style(dialog).configure(
            "PaletteRole.TLabel", background=self.colors["card"],
            foreground=self.colors["heading"], font=("Segoe UI", 11, "bold"))
        modes = (("dark", "Mode Gelap"), ("light", "Mode Terang"))
        self.dark_defaults = {
            "app": "#000000", "card": "#0F172A", "field": "#1E293B",
            "accent": "#3B82F6", "button": "#1E293B", "danger": "#E11D48",
        }
        self.light_defaults = {
            "app": "#F1F5F9", "card": "#FFFFFF", "field": "#FFFFFF",
            "accent": "#2563EB", "button": "#E2E8F0", "danger": "#E11D48",
        }

        def reset_palette_mode(mode):
            self.palette_overrides.pop(mode, None)
            save_palette()
            defaults = self.dark_defaults if mode == "dark" else self.light_defaults
            for role, _label in definitions:
                value = defaults[role]
                swatch = color_buttons[(mode, role)]
                swatch.configure(text=value, bg=value, activebackground=value,
                                 fg=_contrasting_text_color(value),
                                 activeforeground=_contrasting_text_color(value))
            if mode == ("dark" if self.dark_mode else "light"):
                self._style()
                self._apply_widget_theme()
                dialog.configure(bg=self.colors["app"])

        for mode, title in modes:
            page = ttk.Frame(notebook, padding=12)
            notebook.add(page, text=title)
            defaults = self.dark_defaults if mode == "dark" else self.light_defaults
            for row, (role, label) in enumerate(definitions):
                value = self.palette_overrides.get(mode, {}).get(role, defaults[role])
                card = ttk.Frame(page, padding=(10, 8), style="Card.TFrame")
                card.grid(row=row, column=0, sticky="ew", pady=4)
                ttk.Label(card, text=label, style="PaletteRole.TLabel").pack(side="left", padx=(0, 12))
                swatch = tk.Button(card, text=value.upper(), width=12, relief="flat", bd=0,
                                   bg=value, fg=_contrasting_text_color(value),
                                   activebackground=value,
                                   activeforeground=_contrasting_text_color(value),
                                   command=lambda m=mode, r=role: apply_color(m, r, color_buttons[(m, r)])
                                   )
                swatch.pack(side="right")
                color_buttons[(mode, role)] = swatch
            page.columnconfigure(0, weight=1)
        footer = ttk.Frame(dialog, padding=(16, 4, 16, 14))
        footer.pack(fill="x")
        ttk.Button(footer, text="Reset Mode Gelap",
                   command=lambda: reset_palette_mode("dark")).pack(side="left")
        ttk.Button(footer, text="Reset Mode Terang",
                   command=lambda: reset_palette_mode("light")).pack(side="left", padx=(8, 0))
        ttk.Button(footer, text="Tutup", command=dialog.destroy).pack(side="right")
        dialog.grab_set()

    def _build(self):
        header = ttk.Frame(self, padding=(22, 17))
        header.pack(fill="x")
        ttk.Label(header, text="SAHABAT IBADAH", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="Presentasi ibadah sederhana · tersimpan lokal", style="Sub.TLabel").pack(side="left", padx=18, pady=(8, 0))
        ttk.Button(header, text="▶  Mulai Proyeksi", style="Accent.TButton", command=self.open_projector).pack(side="right")
        ttk.Button(header, text="Data & Backup", command=self.open_data_tools).pack(side="right", padx=8)
        ttk.Button(header, text="🎨  List Warna", command=self.open_color_palette).pack(side="right", padx=8)
        self.theme_button = ttk.Button(
            header, text="☼  Mode Terang" if self.dark_mode else "☾  Mode Gelap",
            command=self.toggle_dark_mode)
        self.theme_button.pack(side="right", padx=(0, 8))

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self.schedule_tab = ttk.Frame(self.tabs, padding=18)
        self.resources_tab = ttk.Frame(self.tabs, padding=18)
        self.songs_tab = ttk.Frame(self.tabs, padding=18)
        self.verses_tab = ttk.Frame(self.tabs, padding=18)
        self.camera_tab = ttk.Frame(self.tabs, padding=18)
        self.broadcast_tab = ttk.Frame(self.tabs, padding=18)
        self.tabs.add(self.schedule_tab, text="  Susunan Ibadah  ")
        self.tabs.add(self.resources_tab, text="  Resources  ")
        self.tabs.add(self.songs_tab, text="  Perpustakaan Lagu  ")
        self.tabs.add(self.verses_tab, text="  Ayat Alkitab  ")
        self.tabs.add(self.camera_tab, text="  Kamera & Live  ")
        self.tabs.add(self.broadcast_tab, text="  Broadcast  ")
        self.tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._build_schedule()
        self._build_resources()
        self._build_songs()
        self._build_verses()
        self._build_camera()
        self._build_broadcast()

    def _build_camera(self):
        pane = ttk.Panedwindow(self.camera_tab, orient="horizontal")
        pane.pack(fill="both", expand=True)
        sources = ttk.Frame(pane, style="Card.TFrame", padding=18)
        overlay_host = ttk.Frame(pane, style="Card.TFrame")
        pane.add(sources, weight=2); pane.add(overlay_host, weight=3)
        self.overlay_canvas = tk.Canvas(overlay_host, bg="white", highlightthickness=0, bd=0)
        overlay_scrollbar = ttk.Scrollbar(overlay_host, orient="vertical",
                                          command=self.overlay_canvas.yview)
        self.overlay_canvas.configure(yscrollcommand=overlay_scrollbar.set)
        self.overlay_canvas.pack(side="left", fill="both", expand=True)
        overlay_scrollbar.pack(side="right", fill="y")
        overlays = ttk.Frame(self.overlay_canvas, style="Card.TFrame", padding=18)
        self.overlays_inner = overlays
        overlays_window = self.overlay_canvas.create_window((0, 0), window=overlays, anchor="nw")
        overlays.bind("<Configure>", lambda _event: self.overlay_canvas.configure(
            scrollregion=self.overlay_canvas.bbox("all")))
        self.overlay_canvas.bind("<Configure>", lambda event: self.overlay_canvas.itemconfigure(
            overlays_window, width=event.width))
        self.bind_all("<MouseWheel>", self._scroll_camera_panel, add="+")
        ttk.Label(sources, text="Sumber Video", font=("Segoe UI", 16, "bold"), background="white").pack(anchor="w")
        ttk.Label(sources, text="Kamera bawaan, webcam USB, atau capture device yang terdeteksi sebagai kamera.",
                  background="white", foreground="#65758b", wraplength=360).pack(anchor="w", pady=(5, 14))
        ttk.Label(sources, text="Kamera", background="white").pack(anchor="w")
        source_row = ttk.Frame(sources, style="Card.TFrame")
        source_row.pack(fill="x", pady=(4, 8))
        self.camera_choice = ttk.Combobox(source_row, state="readonly")
        self.camera_choice.pack(side="left", fill="x", expand=True)
        ttk.Button(source_row, text="Cari", command=self.refresh_camera_devices).pack(side="left", padx=(6, 0))
        ttk.Label(sources, text="Kamera kedua (opsional · layar dibagi 50:50)", background="white").pack(anchor="w")
        self.camera2_choice = ttk.Combobox(sources, state="readonly")
        self.camera2_choice.pack(fill="x", pady=(4, 10))
        ttk.Button(sources, text="Pilih File Video…", command=self.choose_video_file).pack(anchor="w", pady=(0, 14))
        ttk.Label(sources, text="Layar Proyeksi", background="white").pack(anchor="w")
        self.display_choice = ttk.Combobox(sources, state="readonly")
        self.display_choice.pack(fill="x", pady=(4, 8))
        self.display_choice.bind("<<ComboboxSelected>>", self._display_choice_selected)
        self._display_preference = None
        self.display_options = available_displays()
        self.display_labels = [f"Layar {i + 1} · {monitor[2]}×{monitor[3]}" for i, monitor in enumerate(self.display_options)]
        self.display_choice["values"] = self.display_labels
        self.display_choice.current(1 if len(self.display_labels) > 1 else 0)
        ttk.Label(sources, text="Dengan dua layar, proyeksi dibuka layar kedua dan kontrol tetap di layar operator. "
                                "Dengan satu layar, tekan F11 untuk layar penuh dan Esc untuk kembali.",
                  background="white", foreground="#65758b", wraplength=360, justify="left").pack(anchor="w", pady=(3, 14))
        ttk.Button(sources, text="▶  Mulai Kamera / Video", style="Accent.TButton",
                   command=self.start_live_projection).pack(anchor="w")
        ttk.Button(sources, text="■  Stop Kamera / Video", command=self.stop_live_projection).pack(anchor="w", pady=(6, 0))
        ttk.Label(sources, text="Tampilan Kamera Live", background="white").pack(anchor="w", pady=(12, 4))
        camera_modes = ttk.Frame(sources, style="Card.TFrame")
        camera_modes.pack(fill="x")
        ttk.Button(camera_modes, text="Kamera 1 penuh",
                   command=lambda: self.set_live_camera_mode("primary")).grid(row=0, column=0, sticky="ew", padx=(0, 4), pady=2)
        ttk.Button(camera_modes, text="Kamera 2 penuh",
                   command=lambda: self.set_live_camera_mode("secondary")).grid(row=0, column=1, sticky="ew", padx=(4, 0), pady=2)
        ttk.Button(camera_modes, text="Dua kamera 50:50",
                   command=lambda: self.set_live_camera_mode("split")).grid(row=1, column=0, sticky="ew", padx=(0, 4), pady=2)
        ttk.Button(camera_modes, text="Sembunyikan kamera",
                   command=lambda: self.set_live_camera_mode("hidden")).grid(row=1, column=1, sticky="ew", padx=(4, 0), pady=2)
        camera_modes.columnconfigure(0, weight=1)
        camera_modes.columnconfigure(1, weight=1)
        self.camera_status = ttk.Label(sources, text="Tekan Cari untuk memeriksa kamera yang tersedia.", style="Sub.TLabel", wraplength=360)
        self.camera_status.pack(anchor="w", pady=(12, 0))

        ttk.Label(overlays, text="Teks Live", font=("Segoe UI", 16, "bold"), background="white").pack(anchor="w")
        ttk.Label(overlays, text="Ketik teks sendiri, atau muat lirik / ayat dari Susunan Ibadah. "
                                "Pilih ayat tertentu dari daftar Ayat Aktif; perubahan tampil langsung di video.",
                  background="white", foreground="#65758b", wraplength=500, justify="left").pack(anchor="w", pady=(5, 12))
        ttk.Label(overlays, text="Teks yang ditampilkan", background="white").pack(anchor="w")
        self.live_text = tk.Text(overlays, height=8, wrap="word", font=("Segoe UI", 14), relief="solid", bd=1)
        self.live_text.pack(fill="both", expand=True, pady=(4, 10))
        options = ttk.Frame(overlays, style="Card.TFrame")
        options.pack(fill="x", pady=(0, 10))
        ttk.Label(options, text="Posisi:", background="white").pack(side="left")
        self.live_overlay_position = ttk.Combobox(options, state="readonly",
                                                  values=list(OVERLAY_POSITIONS), width=12)
        self.live_overlay_position.set(self.live_text_position)
        self.live_overlay_position.pack(side="left", padx=(7, 16))
        self.live_overlay_position.bind("<<ComboboxSelected>>", self.set_live_overlay_position)
        self.live_text_color_button = ttk.Button(options, text="Warna Teks…",
                                                 command=self.choose_live_text_color)
        self.live_text_color_button.pack(side="left", padx=(0, 8))
        ttk.Button(options, text="Muat Lagu Terpilih", command=self.load_live_song).pack(side="left")
        ttk.Button(options, text="Muat Ayat dari Susunan", command=self.load_live_verse).pack(side="left", padx=(6, 0))
        layout_row = ttk.Frame(overlays, style="Card.TFrame")
        layout_row.pack(fill="x", pady=(0, 8))
        self.live_layout_button = ttk.Button(layout_row, text="Atur Posisi & Margin…",
                                             command=self.open_live_layout_dialog)
        self.live_layout_button.pack(side="left")
        self.live_layout_status = ttk.Label(layout_row, text="", style="Sub.TLabel")
        self.live_layout_status.pack(side="left", padx=(9, 0))
        stroke_row = ttk.Frame(overlays, style="Card.TFrame")
        stroke_row.pack(fill="x", pady=(0, 8))
        self.live_stroke_button = ttk.Button(stroke_row, text="Garis Tepi Teks…",
                                             command=self.open_live_stroke_dialog)
        self.live_stroke_button.pack(side="left")
        ttk.Label(stroke_row, text="Garis tepi menjaga teks tetap terbaca di atas video atau gambar.",
                  background="white", foreground="#65758b").pack(side="left", padx=(9, 0))
        background_row = ttk.Frame(overlays, style="Card.TFrame")
        background_row.pack(fill="x", pady=(0, 8))
        self.live_background_var = tk.BooleanVar(value=self.live_text_background)
        ttk.Checkbutton(background_row, text="Kotak latar teks",
                        variable=self.live_background_var,
                        command=self.set_live_text_background).pack(side="left")
        ttk.Label(background_row, text="Centang untuk menambah kotak gelap di belakang teks (opsional).",
                  background="white", foreground="#65758b").pack(side="left", padx=(9, 0))
        lyric_actions = ttk.LabelFrame(overlays, text="Navigasi lirik", style="LiveNav.TLabelframe",
                                      padding=(12, 10))
        lyric_actions.pack(fill="x", pady=(8, 0))
        lyric_nav_row = ttk.Frame(lyric_actions, style="LiveNavInner.TFrame")
        lyric_nav_row.pack(fill="x")
        self.live_song_previous = ttk.Button(lyric_nav_row, text="◀  Lirik sebelumnya",
                                             style="LivePrevious.TButton",
                                             command=lambda: self.step_live_song(-1))
        self.live_song_previous.pack(side="left")
        self.live_song_position = ttk.Label(lyric_nav_row, text="Belum ada lirik",
                                            style="LiveCount.TLabel", anchor="center")
        self.live_song_position.pack(side="left", expand=True, padx=12)
        self.live_song_next = ttk.Button(lyric_nav_row, text="Lirik berikutnya  ▶",
                                         style="LiveNext.TButton",
                                         command=lambda: self.step_live_song(1))
        self.live_song_next.pack(side="left")
        ttk.Button(lyric_actions, text="Hapus Teks", command=self.clear_live_text).pack(anchor="e", pady=(8, 0))
        verse_actions = ttk.LabelFrame(overlays, text="Navigasi ayat", style="LiveNav.TLabelframe",
                                       padding=(12, 10))
        verse_actions.pack(fill="x", pady=(8, 0))
        verse_select_row = ttk.Frame(verse_actions, style="LiveNavInner.TFrame")
        verse_select_row.pack(fill="x", pady=(0, 8))
        ttk.Label(verse_select_row, text="Pilih ayat:", background="#f4f7fa").pack(side="left")
        self.live_verse_choice = ttk.Combobox(verse_select_row, state="readonly", width=36)
        self.live_verse_choice.pack(side="left", padx=(7, 0), fill="x", expand=True)
        self.live_verse_choice.bind("<<ComboboxSelected>>", self.select_live_verse)
        verse_nav_row = ttk.Frame(verse_actions, style="LiveNavInner.TFrame")
        verse_nav_row.pack(fill="x")
        self.live_verse_previous = ttk.Button(verse_nav_row, text="◀  Ayat sebelumnya",
                                              style="LivePrevious.TButton",
                                              command=lambda: self.step_live_verse(-1))
        self.live_verse_previous.pack(side="left")
        self.live_verse_position = ttk.Label(verse_nav_row, text="Belum ada ayat",
                                             style="LiveCount.TLabel", anchor="center")
        self.live_verse_position.pack(side="left", expand=True, padx=12)
        self.live_verse_next = ttk.Button(verse_nav_row, text="Ayat berikutnya  ▶",
                                          style="LiveNext.TButton",
                                          command=lambda: self.step_live_verse(1))
        self.live_verse_next.pack(side="left")
        self.image_actions = ttk.Frame(overlays, style="Card.TFrame")
        self.image_actions.pack(fill="x", pady=(8, 0))
        ttk.Button(self.image_actions, text="Pilih Media…", command=self.choose_live_media).pack(side="left")
        ttk.Button(self.image_actions, text="Kamera + Media 50/50",
                   command=self.show_live_split).pack(side="left", padx=6)
        ttk.Button(self.image_actions, text="Media Layar Penuh",
                   command=self.show_live_image_full).pack(side="left")
        ttk.Button(self.image_actions, text="Sembunyikan Media",
                   command=self.hide_live_image).pack(side="left", padx=6)
        self.slide_actions = ttk.Frame(overlays, style="Card.TFrame")
        ttk.Button(self.slide_actions, text="◀ Slide",
                   command=lambda: self.step_live_presentation(-1)).pack(side="left")
        self.live_slide_choice = ttk.Combobox(self.slide_actions, state="readonly", width=10)
        self.live_slide_choice.pack(side="left", padx=6)
        self.live_slide_choice.bind("<<ComboboxSelected>>", self.select_live_presentation_slide)
        ttk.Button(self.slide_actions, text="Slide ▶",
                   command=lambda: self.step_live_presentation(1)).pack(side="left")
        self.live_slide_status = ttk.Label(self.slide_actions, text="", style="Sub.TLabel")
        self.live_slide_status.pack(side="left", padx=(9, 0))
        self.live_image_status = ttk.Label(overlays, text="Belum ada media live dipilih.", style="Sub.TLabel")
        self.live_image_status.pack(anchor="w", pady=(5, 0))
        self.camera_devices = {}
        self.video_input_path = None
        self.live_image_path = None
        self.live_image_mode = "camera"
        self.live_media_kind = None
        self.live_presentation_path = None
        self.live_presentation_slides = []
        self.live_presentation_index = -1
        self.live_song_slides = []
        self.live_song_index = -1
        self.live_verse_slides = []
        self.live_verse_index = -1
        self.live_mode = None
        self.live_window = None
        self.live_camera_mode = "primary"
        self._update_live_layout_status()
        self._update_live_navigation()

    def _build_broadcast(self):
        root = ttk.Frame(self.broadcast_tab)
        root.pack(fill="both", expand=True)
        output_actions = ttk.Frame(root, style="Card.TFrame")
        output_actions.pack(fill="x", pady=(0, 10))
        ttk.Label(
            output_actions,
            text="Kontrol siaran · Untuk suara, pilih mikrofon di Meet/Zoom secara terpisah",
            font=("Segoe UI", 10, "bold"), wraplength=700, justify="left").pack(
                side="left", padx=10, pady=8)
        self.broadcast_camera_stop_button = ttk.Button(
            output_actions, text="■  Hentikan Sumber · Tampilkan Hitam",
            style="Danger.TButton", command=self.stop_broadcast_camera)
        self.broadcast_camera_stop_button.pack(side="right", padx=8, pady=6)
        controls_shell = ttk.Frame(root, style="Card.TFrame", width=440)
        controls_shell.pack(side="left", fill="y", padx=(0, 12))
        controls_shell.pack_propagate(False)
        self.broadcast_controls_canvas = tk.Canvas(
            controls_shell, highlightthickness=0, borderwidth=0,
            background=self.colors["card"])
        self.broadcast_controls_canvas.pack(side="left", fill="both", expand=True)
        self.broadcast_controls_scrollbar = ttk.Scrollbar(
            controls_shell, orient="vertical", command=self.broadcast_controls_canvas.yview)
        self.broadcast_controls_scrollbar.pack(side="right", fill="y")
        self.broadcast_controls_canvas.configure(yscrollcommand=self.broadcast_controls_scrollbar.set)
        controls = ttk.Frame(self.broadcast_controls_canvas, style="Card.TFrame", padding=18)
        self.broadcast_controls_window = self.broadcast_controls_canvas.create_window(
            (0, 0), window=controls, anchor="nw")
        controls.bind("<Configure>", lambda _event: self.broadcast_controls_canvas.configure(
            scrollregion=self.broadcast_controls_canvas.bbox("all")))
        self.broadcast_controls_canvas.bind("<Configure>", lambda event: self.broadcast_controls_canvas.itemconfigure(
            self.broadcast_controls_window, width=event.width))
        preview_panel = ttk.Frame(root, style="Card.TFrame", padding=18)
        preview_panel.pack(side="left", fill="both", expand=True)

        ttk.Label(controls, text="Sumber Siaran", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(controls, text="Kirim kamera atau konten susunan ibadah ke OBS Virtual Camera.",
                  style="Sub.TLabel", wraplength=310, justify="left").pack(anchor="w", pady=(4, 14))
        ttk.Label(controls, text="Sumber output").pack(anchor="w")
        self.broadcast_source_choice = ttk.Combobox(
            controls, state="readonly", values=("Kamera", "Item Susunan", "Kamera + Item Susunan"))
        self.broadcast_source_choice.current(0)
        self.broadcast_source_choice.pack(fill="x", pady=(4, 10))
        self.broadcast_source_choice.bind("<<ComboboxSelected>>", self._broadcast_source_selected)
        ttk.Label(controls, text="Item susunan").pack(anchor="w")
        self.broadcast_item_choice = ttk.Combobox(controls, state="readonly")
        self.broadcast_item_choice.pack(fill="x", pady=(4, 5))
        self.broadcast_item_choice.bind("<<ComboboxSelected>>", self._broadcast_item_selected)
        content_nav = ttk.Frame(controls, style="Card.TFrame")
        content_nav.pack(fill="x", pady=(0, 10))
        self.broadcast_prev_button = ttk.Button(content_nav, text="← Sebelumnya",
                                                command=lambda: self._broadcast_step_content(-1))
        self.broadcast_prev_button.pack(side="left", fill="x", expand=True)
        self.broadcast_next_button = ttk.Button(content_nav, text="Berikutnya →",
                                                command=lambda: self._broadcast_step_content(1))
        self.broadcast_next_button.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.broadcast_content_status = ttk.Label(controls, text="Pilih item susunan untuk menyiapkan konten.",
                                                  style="Sub.TLabel", wraplength=310)
        self.broadcast_content_status.pack(anchor="w", pady=(0, 10))
        self._broadcast_refresh_item_choices()
        text_style = ttk.LabelFrame(controls, text="Posisi dan arah teks", padding=8)
        text_style.pack(fill="x", pady=(0, 12))
        self.broadcast_song_text_color = "#FFFFFF"
        self.broadcast_verse_text_color = "#FFFFFF"
        ttk.Label(text_style, text="Warna teks lagu").pack(anchor="w")
        self.broadcast_song_color_button = ttk.Button(
            text_style, text="Warna teks lagu · #FFFFFF", command=self._broadcast_choose_song_text_color)
        self.broadcast_song_color_button.pack(fill="x", pady=(0, 6))
        ttk.Label(text_style, text="Warna teks ayat Alkitab").pack(anchor="w")
        self.broadcast_verse_color_button = ttk.Button(
            text_style, text="Warna teks ayat · #FFFFFF", command=self._broadcast_choose_verse_text_color)
        self.broadcast_verse_color_button.pack(fill="x", pady=(0, 6))
        ttk.Label(text_style, text="Posisi mendatar").pack(anchor="w")
        self.broadcast_text_halign = ttk.Combobox(
            text_style, state="readonly", values=("Kiri", "Tengah", "Kanan"))
        self.broadcast_text_halign.set("Tengah")
        self.broadcast_text_halign.pack(fill="x", pady=(2, 6))
        self.broadcast_text_halign.bind("<<ComboboxSelected>>", self._broadcast_text_style_changed)
        ttk.Label(text_style, text="Posisi vertikal").pack(anchor="w")
        self.broadcast_text_valign = ttk.Combobox(
            text_style, state="readonly", values=("Atas", "Tengah", "Bawah"))
        self.broadcast_text_valign.set("Tengah")
        self.broadcast_text_valign.pack(fill="x", pady=(2, 6))
        self.broadcast_text_valign.bind("<<ComboboxSelected>>", self._broadcast_text_style_changed)
        ttk.Label(text_style, text="Margin kiri / kanan (%)").pack(anchor="w")
        self.broadcast_margin_x = tk.IntVar(value=6)
        self.broadcast_text_style_after = None
        tk.Scale(text_style, from_=0, to=30, orient="horizontal", variable=self.broadcast_margin_x,
                 showvalue=True, command=lambda _value: self._broadcast_schedule_text_style_changed()).pack(fill="x")
        ttk.Label(text_style, text="Margin atas / bawah (%)").pack(anchor="w")
        self.broadcast_margin_y = tk.IntVar(value=8)
        tk.Scale(text_style, from_=0, to=30, orient="horizontal", variable=self.broadcast_margin_y,
                 showvalue=True, command=lambda _value: self._broadcast_schedule_text_style_changed()).pack(fill="x")
        self.broadcast_flip_horizontal = tk.BooleanVar(value=False)
        self.broadcast_flip_enabled = False
        ttk.Checkbutton(text_style, text="Balik horizontal seluruh output",
                        variable=self.broadcast_flip_horizontal,
                        command=self._broadcast_flip_toggled).pack(anchor="w", pady=(4, 0))
        ttk.Label(controls, text="Kamera 1").pack(anchor="w")
        camera1_row = ttk.Frame(controls, style="Card.TFrame")
        camera1_row.pack(fill="x", pady=(4, 10))
        self.broadcast_camera1_choice = ttk.Combobox(camera1_row, state="readonly", width=30)
        self.broadcast_camera1_choice.pack(side="left", fill="x", expand=True)
        self.broadcast_camera1_choice.bind("<<ComboboxSelected>>", self._broadcast_camera_choice_changed)
        ttk.Button(camera1_row, text="Cari", command=self.refresh_camera_devices).pack(side="left", padx=(6, 0))
        ttk.Label(controls, text="Kamera 2").pack(anchor="w")
        self.broadcast_camera2_choice = ttk.Combobox(controls, state="readonly")
        self.broadcast_camera2_choice.pack(fill="x", pady=(4, 10))
        self.broadcast_camera2_choice.bind("<<ComboboxSelected>>", self._broadcast_update_mode_options)
        self.broadcast_camera2_choice.bind("<<ComboboxSelected>>", self._broadcast_camera_choice_changed, add="+")
        ttk.Label(controls, text="Tampilan output").pack(anchor="w")
        self.broadcast_mode_choice = ttk.Combobox(
            controls, state="readonly", values=("Kamera 1 penuh", "Kamera 2 penuh", "Dua kamera 50:50"))
        self.broadcast_mode_choice.current(0)
        self.broadcast_mode_choice.pack(fill="x", pady=(4, 14))
        self.broadcast_mode_choice.bind("<<ComboboxSelected>>", self._broadcast_mode_selected)

        self.broadcast_start_button = ttk.Button(controls, text="▶  Mulai Output Kamera / Konten",
                                                  style="Accent.TButton", command=self.start_broadcast)
        self.broadcast_start_button.pack(fill="x", pady=(0, 7))
        self.broadcast_stop_button = ttk.Button(controls, text="■  Hentikan Output",
                                                 command=self.stop_broadcast, state="disabled")
        self.broadcast_stop_button.pack(fill="x", pady=(0, 10))

        ttk.Separator(controls).pack(fill="x", pady=(0, 12))
        ttk.Label(controls, text="Mikrofon", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        self.broadcast_mic_choice = ttk.Combobox(controls, state="readonly")
        self.broadcast_mic_choice.pack(fill="x", pady=(7, 8))
        mic_actions = ttk.Frame(controls, style="Card.TFrame")
        mic_actions.pack(fill="x")
        ttk.Button(mic_actions, text="Cari Mikrofon", command=self.refresh_broadcast_microphones).pack(side="left")
        self.broadcast_mic_button = ttk.Button(mic_actions, text="Aktifkan Mikrofon",
                                               style="Accent.TButton",
                                               command=self.toggle_broadcast_microphone)
        self.broadcast_mic_button.pack(side="left", padx=(7, 0))
        self.broadcast_mic_meter = ttk.Progressbar(controls, maximum=100, mode="determinate")
        self.broadcast_mic_meter.pack(fill="x", pady=(10, 3))
        self.broadcast_mic_status = ttk.Label(controls, text="Mikrofon belum diuji.", style="Sub.TLabel",
                                              wraplength=310)
        self.broadcast_mic_status.pack(anchor="w")
        ttk.Label(preview_panel, text="Pratinjau Output", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(preview_panel, text="Tampilan yang dikirim ke OBS Virtual Camera · 1280 × 720 · 25 FPS",
                  style="Sub.TLabel").pack(anchor="w", pady=(4, 12))
        self.broadcast_preview = tk.Label(preview_panel, text="Output kamera akan muncul di sini",
                                          bg="#080808", fg="#ffffff",
                                          font=("Segoe UI", 16, "bold"), compound="center")
        self.broadcast_preview.pack(fill="both", expand=True)
        self.broadcast_status = ttk.Label(preview_panel, text="Siap · OBS Virtual Camera akan dipilih di aplikasi rapat.",
                                          style="Sub.TLabel")
        self.broadcast_status.pack(anchor="w", pady=(10, 0))
        # Scrolling over a readonly dropdown should scroll the panel, never change its value.
        for choice in (self.broadcast_source_choice, self.broadcast_item_choice,
                       self.broadcast_text_halign, self.broadcast_text_valign,
                       self.broadcast_camera1_choice, self.broadcast_camera2_choice,
                       self.broadcast_mode_choice, self.broadcast_mic_choice):
            choice.bind("<MouseWheel>", self._broadcast_scroll_over_choice)
            choice.bind("<Button-4>", self._broadcast_scroll_over_choice)
            choice.bind("<Button-5>", self._broadcast_scroll_over_choice)
        self.broadcast_preview_photo = None
        self._broadcast_set_source_controls()
        self._broadcast_refresh_preview()
        self.refresh_broadcast_microphones()

    def _broadcast_refresh_preview(self):
        if not hasattr(self, "broadcast_preview") or not self.broadcast_preview.winfo_exists():
            return
        if hasattr(self, "tabs") and self.tabs.select() != str(self.broadcast_tab):
            self.after(500, self._broadcast_refresh_preview)
            return
        content_mode = getattr(self, "broadcast_content_mode", False)
        if content_mode and self.broadcast_content_capture is not None:
            latest_content = self.broadcast_content_capture.latest_frame()
            if latest_content is not None:
                self.broadcast_content_frame = latest_content
        content_frame = self.broadcast_content_frame if content_mode else None
        capture = self.broadcast_capture
        camera_frame = capture.latest_frame() if capture is not None else None
        second = self.broadcast_second_capture
        second_frame = second.latest_frame() if second is not None else None
        mode = getattr(self, "broadcast_output_mode", "primary")
        if self.broadcast_camera_stopped and self.broadcast_thread is not None:
            self.broadcast_preview.configure(bg="#000000", image="", text="")
            self.broadcast_preview_photo = None
        elif (camera_frame is not None or content_frame is not None) and \
                (content_mode or getattr(self, "broadcast_camera_enabled", False)):
            self.broadcast_preview.configure(bg="#000000")
            # Keep UI-side preview work small; the virtual camera composes at full output size.
            composed = self._broadcast_compose_frame(camera_frame, second_frame, content_frame, (640, 360))
            try:
                photo = frame_to_photo(composed, (max(1, self.broadcast_preview.winfo_width()),
                                                   max(1, self.broadcast_preview.winfo_height())),
                                       self.broadcast_preview)
                if photo:
                    self.broadcast_preview_photo = photo
                    self.broadcast_preview.configure(image=photo, text="")
            except (tk.TclError, ValueError):
                pass
        elif self.broadcast_thread is None:
            self.broadcast_preview.configure(image="", text="Output kamera akan muncul di sini")
            self.broadcast_preview_photo = None
        # Tk image conversion is comparatively expensive; refresh preview at 2 FPS.
        self.after(500, self._broadcast_refresh_preview)

    def _broadcast_scroll_over_choice(self, event):
        """Route mouse-wheel events above a dropdown to panel scrolling, not selection."""
        if event.num == 4:
            step = -1
        elif event.num == 5:
            step = 1
        else:
            step = int(-event.delta / 120) if event.delta else -1
        self.broadcast_controls_canvas.yview_scroll(step or (-1 if event.delta > 0 else 1), "units")
        return "break"

    def _broadcast_set_source_controls(self):
        source = getattr(self, "broadcast_source_choice", None)
        source = source.get() if source is not None else "Kamera"
        content_mode = source in {"Item Susunan", "Kamera + Item Susunan"}
        camera_enabled = source in {"Kamera", "Kamera + Item Susunan"}
        camera_state = "readonly" if camera_enabled else "disabled"
        camera2_state = "readonly" if source == "Kamera" else "disabled"
        self.broadcast_camera1_choice.configure(state=camera_state)
        self.broadcast_camera2_choice.configure(state=camera2_state)
        self.broadcast_mode_choice.configure(state=camera_state)
        # Keep the item list available even while Camera is selected. Choosing an item
        # should switch the broadcast source automatically, rather than looking disabled.
        item_enabled = source in {"Item Susunan", "Kamera + Item Susunan"}
        self.broadcast_source_choice.configure(state="readonly")
        self.broadcast_item_choice.configure(state="readonly" if item_enabled else "disabled")
        for button in (self.broadcast_prev_button, self.broadcast_next_button):
            button.configure(state="normal" if content_mode and self.broadcast_content_frames else "disabled")
        self._broadcast_update_mode_options()

    def _broadcast_source_selected(self, _event=None):
        self._broadcast_set_source_controls()
        if self.broadcast_source_choice.get() in {"Item Susunan", "Kamera + Item Susunan"}:
            self._broadcast_item_selected()
        else:
            self.broadcast_content_mode = False
            self.broadcast_content_frame = None
            if self.broadcast_content_capture is not None:
                self.broadcast_content_capture.close()
                self.broadcast_content_capture = None
            self.broadcast_content_frames = []
            self.broadcast_preview.configure(image="", text="Output kamera akan muncul di sini")
            self.broadcast_preview_photo = None
            count = len(getattr(self, "broadcast_content_rows", []))
            self.broadcast_content_status.configure(
                text=(f"{count} item tersedia · pilih item untuk menyiarkannya langsung."
                      if count else "Belum ada lagu, ayat, atau media di susunan ini."))
        if self.broadcast_thread and self.broadcast_thread.is_alive():
            source = self.broadcast_source_choice.get()
            self.broadcast_content_mode = source in {"Item Susunan", "Kamera + Item Susunan"}
            self.broadcast_camera_enabled = source in {"Kamera", "Kamera + Item Susunan"}
            self._broadcast_mode_selected()
            if self.broadcast_camera_enabled:
                self._broadcast_camera_choice_changed()
            else:
                self._broadcast_camera_switch_generation += 1
                old_captures = (self.broadcast_capture, self.broadcast_second_capture)
                self.broadcast_capture = self.broadcast_second_capture = None
                self.broadcast_primary_source = self.broadcast_secondary_source = None
                for capture in old_captures:
                    if capture is not None:
                        capture.close()

    def _broadcast_refresh_item_choices(self):
        if not hasattr(self, "broadcast_item_choice"):
            return
        rows = self.store.items(self.active_schedule_id) if getattr(self, "active_schedule_id", None) else []
        self.broadcast_content_rows = [row for row in rows if row["kind"] in
                                       {"Lagu", "Ayat", "Media", "PowerPoint"}]
        self.broadcast_item_labels = {
            f"{index + 1}. {row['title']} · {row['kind']}": index
            for index, row in enumerate(self.broadcast_content_rows)
        }
        values = list(self.broadcast_item_labels)
        self.broadcast_item_choice["values"] = values
        if values:
            current = self.broadcast_item_choice.get()
            self.broadcast_item_choice.set(current if current in self.broadcast_item_labels else values[0])
            self._broadcast_set_source_controls()
            if self.broadcast_source_choice.get() in {"Item Susunan", "Kamera + Item Susunan"}:
                self._broadcast_item_selected()
            else:
                self.broadcast_content_status.configure(
                    text=f"{len(values)} item tersedia · pilih item untuk menyiarkannya langsung.")
        else:
            self.broadcast_item_choice.set("")
            self.broadcast_content_frames = []
            self.broadcast_content_frame = None
            if hasattr(self, "broadcast_content_status"):
                self.broadcast_content_status.configure(text="Belum ada lagu, ayat, atau media di susunan ini.")

    def _broadcast_item_selected(self, _event=None):
        label = self.broadcast_item_choice.get()
        index = getattr(self, "broadcast_item_labels", {}).get(label)
        if index is None:
            return
        if self.broadcast_source_choice.get() == "Kamera":
            self.broadcast_source_choice.set("Item Susunan")
        self._broadcast_set_source_controls()
        row = self.broadcast_content_rows[index]
        if self.broadcast_source_choice.get() == "Kamera + Item Susunan":
            if row["kind"] == "Lagu":
                self.broadcast_mode_choice.set("Lagu Overlay Kamera")
            elif self.broadcast_mode_choice.get() == "Lagu Overlay Kamera":
                self.broadcast_mode_choice.set("Kamera + Item 50:50")
        self.broadcast_content_mode = True
        self.broadcast_content_index = 0
        self.broadcast_content_frames = []
        self.broadcast_content_frame = None
        if self.broadcast_content_capture is not None:
            self.broadcast_content_capture.close()
            self.broadcast_content_capture = None
        try:
            kind = row["kind"]
            if kind == "PowerPoint":
                paths = render_powerpoint_slides(row["media_path"] or "", (1280, 720))
                if not paths:
                    raise RuntimeError("Slide PowerPoint tidak dapat dibaca. Pastikan PowerPoint tersedia.")
                import cv2
                self.broadcast_content_frames = [cv2.imread(str(path)) for path in paths]
                self.broadcast_content_frames = [frame for frame in self.broadcast_content_frames if frame is not None]
            elif kind == "Media":
                path = row["media_path"] or ""
                ext = Path(path).suffix.lower()
                if ext in {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".webm", ".m4v"}:
                    self.broadcast_content_capture = VideoCaptureSource(path, loop=True)
                    self.broadcast_content_status.configure(text=f"Video diputar berulang · {row['title']}")
                    self.broadcast_content_frame = None
                    self.broadcast_content_frames = []
                else:
                    import cv2
                    frame = cv2.imread(str(path))
                    if frame is None:
                        raise RuntimeError("File gambar tidak dapat dibaca.")
                    self.broadcast_content_frames = [self._broadcast_fit_frame(frame)]
            elif kind in {"Lagu", "Ayat"}:
                texts = self.make_slides(row["content"] or "") or [row["content"] or ""]
                self.broadcast_content_frames = [
                    self._broadcast_text_frame(
                        text,
                        horizontal=self.broadcast_text_halign.get(),
                        vertical=self.broadcast_text_valign.get(),
                        margin_x=self.broadcast_margin_x.get(),
                        margin_y=self.broadcast_margin_y.get(),
                        text_color=(self.broadcast_song_text_color if kind == "Lagu" else
                                    self.broadcast_verse_text_color if kind == "Ayat" else "#FFFFFF"))
                    for text in texts if text.strip()]
            if self.broadcast_content_frames:
                self.broadcast_content_frame = self.broadcast_content_frames[0]
                self.broadcast_content_status.configure(
                    text=f"Slide 1/{len(self.broadcast_content_frames)} · {row['title']}")
            elif self.broadcast_content_capture is None:
                raise RuntimeError("Item ini tidak memiliki konten yang dapat ditampilkan.")
        except Exception as error:
            self.broadcast_content_frames = []
            self.broadcast_content_frame = None
            self.broadcast_content_status.configure(text=f"Konten gagal dimuat: {error}")
            messagebox.showerror("Konten Broadcast tidak dapat dimuat", str(error), parent=self)
        self._broadcast_set_source_controls()
        self._broadcast_refresh_preview_content()
        if self.broadcast_thread and self.broadcast_thread.is_alive():
            self._broadcast_mode_selected()

    @staticmethod
    def _broadcast_fit_frame(frame):
        import cv2
        height, width = frame.shape[:2]
        scale = min(1280 / width, 720 / height)
        resized = cv2.resize(frame, (max(1, int(width * scale)), max(1, int(height * scale))))
        canvas = __import__("numpy").zeros((720, 1280, 3), dtype="uint8")
        x, y = (1280 - resized.shape[1]) // 2, (720 - resized.shape[0]) // 2
        canvas[y:y + resized.shape[0], x:x + resized.shape[1]] = resized
        return canvas

    @staticmethod
    def _broadcast_text_frame(text, horizontal="Tengah", vertical="Tengah", margin_x=6, margin_y=8, text_color="#FFFFFF"):
        import cv2
        import numpy as np
        import textwrap
        import unicodedata
        text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode("ascii")
        text = text.replace("’", "'").replace("“", '"').replace("”", '"')
        paragraphs = text.splitlines() or [text]
        scale = 1.3
        thickness = 3
        max_width = 1120
        lines = []
        for paragraph in paragraphs:
            if not paragraph.strip():
                lines.append("")
                continue
            words = paragraph.split()
            current = ""
            for word in words:
                candidate = f"{current} {word}".strip()
                if cv2.getTextSize(candidate, cv2.FONT_HERSHEY_DUPLEX, scale, thickness)[0][0] > max_width and current:
                    lines.append(current)
                    current = word
                else:
                    current = candidate
            if current:
                lines.append(current)
        while len(lines) * 58 > 590 and scale > 0.72:
            scale -= 0.08
            return EasyWorshipApp._broadcast_text_frame_scaled(
                text, scale, horizontal, vertical, margin_x, margin_y, text_color)
        frame = np.zeros((720, 1280, 3), dtype=np.uint8)
        rgb = tuple(int(text_color[index:index + 2], 16) for index in (1, 3, 5))
        bgr = (rgb[2], rgb[1], rgb[0])
        line_height = max(36, int(58 * scale / 1.3))
        text_height = len(lines) * line_height
        if vertical == "Atas":
            y = int(720 * margin_y / 100) + line_height
        elif vertical == "Bawah":
            y = 720 - int(720 * margin_y / 100) - text_height + line_height
        else:
            y = (720 - text_height) // 2 + line_height
        for line in lines:
            size = cv2.getTextSize(line, cv2.FONT_HERSHEY_DUPLEX, scale, thickness)[0]
            if horizontal == "Kiri":
                x = max(8, int(1280 * margin_x / 100))
            elif horizontal == "Kanan":
                x = min(1280 - size[0] - 8, 1280 - int(1280 * margin_x / 100) - size[0])
            else:
                x = max(0, (1280 - size[0]) // 2)
            cv2.putText(frame, line, (x, y), cv2.FONT_HERSHEY_DUPLEX, scale,
                        bgr, thickness, cv2.LINE_AA)
            y += line_height
        return frame

    @staticmethod
    def _broadcast_text_frame_scaled(text, scale, horizontal="Tengah", vertical="Tengah",
                                     margin_x=6, margin_y=8, text_color="#FFFFFF"):
        import cv2
        import numpy as np
        import unicodedata
        text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode("ascii")
        lines = []
        for paragraph in text.splitlines() or [text]:
            words, current = paragraph.split(), ""
            for word in words:
                candidate = f"{current} {word}".strip()
                if cv2.getTextSize(candidate, cv2.FONT_HERSHEY_DUPLEX, scale, 2)[0][0] > 1120 and current:
                    lines.append(current); current = word
                else:
                    current = candidate
            lines.append(current)
        thickness = max(1, round(scale * 2.2))
        frame = np.zeros((720, 1280, 3), dtype=np.uint8)
        rgb = tuple(int(text_color[index:index + 2], 16) for index in (1, 3, 5))
        bgr = (rgb[2], rgb[1], rgb[0])
        line_height = max(30, int(58 * scale / 1.3))
        text_height = len(lines) * line_height
        if vertical == "Atas":
            y = int(720 * margin_y / 100) + line_height
        elif vertical == "Bawah":
            y = 720 - int(720 * margin_y / 100) - text_height + line_height
        else:
            y = (720 - text_height) // 2 + line_height
        for line in lines:
            width = cv2.getTextSize(line, cv2.FONT_HERSHEY_DUPLEX, scale, thickness)[0][0]
            if horizontal == "Kiri":
                x = max(8, int(1280 * margin_x / 100))
            elif horizontal == "Kanan":
                x = min(1280 - width - 8, 1280 - int(1280 * margin_x / 100) - width)
            else:
                x = max(0, (1280 - width) // 2)
            cv2.putText(frame, line, (x, y), cv2.FONT_HERSHEY_DUPLEX,
                        scale, bgr, thickness, cv2.LINE_AA)
            y += line_height
        return frame

    def _broadcast_step_content(self, offset):
        if not self.broadcast_content_frames:
            return
        self.broadcast_content_index = max(0, min(len(self.broadcast_content_frames) - 1,
                                                  self.broadcast_content_index + offset))
        self.broadcast_content_frame = self.broadcast_content_frames[self.broadcast_content_index]
        self._broadcast_refresh_preview_content()

    def _broadcast_refresh_preview_content(self):
        if self.broadcast_content_capture is None and self.broadcast_content_frame is None:
            return
        count = len(self.broadcast_content_frames)
        if count:
            self.broadcast_content_status.configure(
                text=f"Slide {self.broadcast_content_index + 1}/{count} · {self.broadcast_item_choice.get()}")
        content_mode = self.broadcast_source_choice.get() in {"Item Susunan", "Kamera + Item Susunan"}
        self.broadcast_prev_button.configure(
            state="normal" if content_mode and count and self.broadcast_content_index > 0 else "disabled")
        self.broadcast_next_button.configure(
            state="normal" if content_mode and count and self.broadcast_content_index < count - 1 else "disabled")

    def _broadcast_mode_selected(self, _event=None):
        if not (self.broadcast_thread and self.broadcast_thread.is_alive()):
            return
        mode = {"Kamera 1 penuh": "primary", "Kamera 2 penuh": "secondary",
                "Dua kamera 50:50": "split", "Kamera penuh": "primary",
                "Item penuh": "content", "Kamera + Item 50:50": "camera_content_split",
                "Lagu Overlay Kamera": "song_overlay"}.get(
                    self.broadcast_mode_choice.get(), "primary")
        self.broadcast_output_mode = mode
        self.broadcast_status.configure(text=f"Aktif · tampilan {self.broadcast_mode_choice.get()} · OBS Virtual Camera")
        if mode in {"secondary", "split"} and self.broadcast_second_capture is None:
            self._broadcast_camera_choice_changed()

    def _broadcast_camera_choice_changed(self, _event=None):
        """Switch camera inputs while the virtual camera stays live."""
        if not (self.broadcast_thread and self.broadcast_thread.is_alive()) or self.broadcast_camera_stopped:
            return
        if not self.broadcast_camera_enabled:
            return
        try:
            primary_label = self.broadcast_camera1_choice.get()
            primary_source = self.camera_devices[primary_label]
        except (KeyError, tk.TclError):
            self.broadcast_status.configure(text="Pilih Kamera 1 sebelum mengganti sumber siaran.")
            return
        needs_second = self.broadcast_output_mode in {"secondary", "split"}
        secondary_label = self.broadcast_camera2_choice.get() if needs_second else "(Tidak ada)"
        secondary_source = self.camera_devices.get(secondary_label) if secondary_label != "(Tidak ada)" else None
        if needs_second and secondary_source is None:
            self.broadcast_status.configure(text="Pilih Kamera 2 untuk tata letak ini.")
            return
        if secondary_source is not None and secondary_source == primary_source:
            self.broadcast_status.configure(text="Kamera 1 dan Kamera 2 harus berbeda.")
            return
        if (primary_source == self.broadcast_primary_source and
                secondary_source == self.broadcast_secondary_source):
            return

        self._broadcast_camera_switch_generation += 1
        generation = self._broadcast_camera_switch_generation
        old_primary, old_secondary = self.broadcast_capture, self.broadcast_second_capture
        old_primary_source, old_secondary_source = self.broadcast_primary_source, self.broadcast_secondary_source

        def open_capture(source):
            if source is None:
                return None
            if self.live_window and self.live_window.winfo_exists() and source == self.live_source_current:
                return self.live_window.capture.acquire()
            if self.live_window and self.live_window.winfo_exists() and source == self.live_second_source_current:
                active = self.live_window.second_capture
                if active is not None:
                    return active.acquire()
            return VideoCaptureSource(source)

        def switch_worker():
            created = []
            try:
                # Reuse already-open captures when only changing their layout/role.
                new_primary = (old_primary if primary_source == old_primary_source else
                              old_secondary if primary_source == old_secondary_source and old_secondary else None)
                if new_primary is None:
                    new_primary = open_capture(primary_source)
                    if new_primary is not None and new_primary not in {old_primary, old_secondary}:
                        created.append(new_primary)
                new_secondary = (old_secondary if secondary_source == old_secondary_source else
                                old_primary if secondary_source == old_primary_source and old_primary else None)
                if secondary_source is not None and new_secondary is None:
                    new_secondary = open_capture(secondary_source)
                    if new_secondary is not None and new_secondary not in {old_primary, old_secondary}:
                        created.append(new_secondary)
                def apply_switch():
                    if generation != self._broadcast_camera_switch_generation or self._closing:
                        for capture in created:
                            capture.close()
                        return
                    self.broadcast_capture = new_primary
                    self.broadcast_second_capture = new_secondary
                    self.broadcast_primary_source = primary_source
                    self.broadcast_secondary_source = secondary_source
                    for capture in (old_primary, old_secondary):
                        if capture is not None and capture not in {new_primary, new_secondary}:
                            capture.close()
                    self.broadcast_status.configure(text="Sumber kamera diperbarui · output tetap aktif.")
                self.after(0, apply_switch)
            except Exception as error:
                for capture in created:
                    capture.close()
                try:
                    self.after(0, lambda message=str(error): self.broadcast_status.configure(
                        text=f"Kamera tidak dapat diganti: {message}"))
                except tk.TclError:
                    pass

        threading.Thread(target=switch_worker, daemon=True).start()

    def _broadcast_update_mode_options(self, _event=None):
        """Hide layouts that need camera 2 when no distinct second source is selected."""
        has_second = self.broadcast_camera2_choice.get() not in {"", "(Tidak ada)"}
        source = self.broadcast_source_choice.get()
        item_index = getattr(self, "broadcast_item_labels", {}).get(self.broadcast_item_choice.get())
        selected_kind = (self.broadcast_content_rows[item_index]["kind"]
                         if item_index is not None and item_index < len(self.broadcast_content_rows) else None)
        if source == "Item Susunan":
            modes = ("Item penuh",)
        elif source == "Kamera + Item Susunan":
            if selected_kind == "Lagu":
                modes = ("Lagu Overlay Kamera", "Kamera + Item 50:50", "Kamera penuh", "Item penuh")
            else:
                modes = ("Kamera + Item 50:50", "Kamera penuh", "Item penuh")
        else:
            modes = (("Kamera 1 penuh", "Kamera 2 penuh", "Dua kamera 50:50")
                     if has_second else ("Kamera 1 penuh",))
        self.broadcast_mode_choice.configure(values=modes)
        if self.broadcast_mode_choice.get() not in modes:
            self.broadcast_mode_choice.set(modes[0])
            if self.broadcast_thread and self.broadcast_thread.is_alive():
                self._broadcast_mode_selected()

    def _broadcast_schedule_text_style_changed(self):
        if self.broadcast_text_style_after is not None:
            try:
                self.after_cancel(self.broadcast_text_style_after)
            except tk.TclError:
                pass
        self.broadcast_text_style_after = self.after(140, self._broadcast_text_style_changed)

    def _broadcast_text_style_changed(self, _event=None):
        self.broadcast_text_style_after = None
        """Re-render the selected song/verse slide when text placement changes."""
        if not hasattr(self, "broadcast_content_rows"):
            return
        index = getattr(self, "broadcast_item_labels", {}).get(self.broadcast_item_choice.get())
        if index is None:
            return
        row = self.broadcast_content_rows[index]
        if row["kind"] not in {"Lagu", "Ayat"}:
            return
        texts = self.make_slides(row["content"] or "") or [row["content"] or ""]
        self.broadcast_content_frames = [
            self._broadcast_text_frame(
                text,
                horizontal=self.broadcast_text_halign.get(),
                vertical=self.broadcast_text_valign.get(),
                margin_x=self.broadcast_margin_x.get(),
                margin_y=self.broadcast_margin_y.get(),
                text_color=(self.broadcast_song_text_color if row["kind"] == "Lagu" else
                            self.broadcast_verse_text_color if row["kind"] == "Ayat" else "#FFFFFF"))
            for text in texts if text.strip()]
        if self.broadcast_content_frames:
            self.broadcast_content_index = min(getattr(self, "broadcast_content_index", 0),
                                               len(self.broadcast_content_frames) - 1)
            self.broadcast_content_frame = self.broadcast_content_frames[self.broadcast_content_index]
            self._broadcast_refresh_preview_content()

    def _broadcast_choose_song_text_color(self):
        _rgb, color = colorchooser.askcolor(
            color=self.broadcast_song_text_color, title="Pilih warna teks lagu", parent=self)
        if not color:
            return
        self.broadcast_song_text_color = color.upper()
        self.broadcast_song_color_button.configure(
            text=f"Warna teks lagu · {self.broadcast_song_text_color}")
        self._broadcast_text_style_changed()

    def _broadcast_choose_verse_text_color(self):
        _rgb, color = colorchooser.askcolor(
            color=self.broadcast_verse_text_color, title="Pilih warna teks ayat Alkitab", parent=self)
        if not color:
            return
        self.broadcast_verse_text_color = color.upper()
        self.broadcast_verse_color_button.configure(
            text=f"Warna teks ayat · {self.broadcast_verse_text_color}")
        self._broadcast_text_style_changed()

    def _broadcast_flip_toggled(self):
        self.broadcast_flip_enabled = bool(self.broadcast_flip_horizontal.get())
        self._broadcast_refresh_preview()

    def _broadcast_compose_frame(self, camera, second_camera, content, size):
        """Build the same unmirrored/mirrored composition for preview and virtual camera."""
        import cv2
        import numpy as np
        mode = getattr(self, "broadcast_output_mode", "primary")
        width, height = size
        if mode == "content":
            frame = self._broadcast_fit_frame(content) if content is not None else np.zeros((height, width, 3), np.uint8)
            frame = cv2.resize(frame, (width, height))
        elif mode == "camera_content_split":
            frame = np.zeros((height, width, 3), dtype=np.uint8)
            half = width // 2
            self._broadcast_place_frame(frame, camera, 0, 0, half, height, contain=True)
            self._broadcast_place_frame(frame, content, half, 0, width - half, height, contain=True)
            frame[:, half - 1:half + 1] = (210, 216, 224)
        elif mode == "song_overlay":
            frame = self._compose_broadcast_frame(camera, None, "primary", (width, height))
            if content is not None:
                overlay = content
                if overlay.shape[1] != width or overlay.shape[0] != height:
                    overlay = cv2.resize(overlay, (width, height))
                mask = overlay.max(axis=2).astype(np.uint8)
                x, y, box_width, box_height = cv2.boundingRect(mask)
                rgb = tuple(int(self.broadcast_song_text_color[index:index + 2], 16)
                            for index in (1, 3, 5))
                alpha_max = max(rgb)
                if box_width and box_height and alpha_max:
                    overlay_roi = overlay[y:y + box_height, x:x + box_width].astype(np.float32)
                    alpha = (mask[y:y + box_height, x:x + box_width].astype(np.float32) /
                             alpha_max).clip(0.0, 1.0)[..., None]
                    target = frame[y:y + box_height, x:x + box_width].astype(np.float32)
                    frame[y:y + box_height, x:x + box_width] = (
                        target * (1.0 - alpha) + overlay_roi * alpha).astype(np.uint8)
        elif mode in {"secondary", "split"}:
            frame = self._compose_broadcast_frame(camera, second_camera, mode, (width, height))
        else:
            frame = self._compose_broadcast_frame(camera, None, "primary", (width, height))
        if getattr(self, "broadcast_flip_enabled", False):
            frame = cv2.flip(frame, 1)
        return frame

    @staticmethod
    def _broadcast_place_frame(output, frame, x, y, width, height, contain=False):
        if frame is None:
            return
        import cv2
        src_h, src_w = frame.shape[:2]
        scale = min(width / src_w, height / src_h) if contain else max(width / src_w, height / src_h)
        rw, rh = max(1, int(src_w * scale)), max(1, int(src_h * scale))
        resized = cv2.resize(frame, (rw, rh), interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
        if contain:
            dx, dy = (width - rw) // 2, (height - rh) // 2
            output[y + dy:y + dy + rh, x + dx:x + dx + rw] = resized
        else:
            sx, sy = max(0, (rw - width) // 2), max(0, (rh - height) // 2)
            output[y:y + height, x:x + width] = resized[sy:sy + height, sx:sx + width]

    @staticmethod
    def _compose_broadcast_frame(primary, secondary, mode, size):
        import cv2
        import numpy as np

        width, height = size
        output = np.zeros((height, width, 3), dtype=np.uint8)

        def place(frame, x, y, region_width, region_height, contain=False):
            if frame is None or region_width < 2 or region_height < 2:
                return
            src_h, src_w = frame.shape[:2]
            scale = (min(region_width / src_w, region_height / src_h) if contain else
                     max(region_width / src_w, region_height / src_h))
            resized_w = max(1, int(src_w * scale))
            resized_h = max(1, int(src_h * scale))
            resized = cv2.resize(frame, (resized_w, resized_h),
                                 interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
            if contain:
                # Use black letterboxing instead of a per-frame blurred background.
                dx, dy = (region_width - resized_w) // 2, (region_height - resized_h) // 2
                output[y + dy:y + dy + resized_h, x + dx:x + dx + resized_w] = resized
            else:
                sx, sy = max(0, (resized_w - region_width) // 2), max(0, (resized_h - region_height) // 2)
                output[y:y + region_height, x:x + region_width] = resized[sy:sy + region_height,
                                                                           sx:sx + region_width]

        if mode == "secondary":
            place(secondary, 0, 0, width, height)
        elif mode == "split":
            half = width // 2
            place(primary, 0, 0, half, height, contain=True)
            place(secondary, half, 0, width - half, height, contain=True)
        else:
            place(primary, 0, 0, width, height)
        if mode == "split":
            output[:, width // 2 - 1:width // 2 + 1] = (210, 216, 224)
        return output

    def refresh_broadcast_microphones(self):
        try:
            import sounddevice as sd
            devices = sd.query_devices()
        except (ImportError, OSError) as error:
            self.broadcast_mic_choice["values"] = []
            self.broadcast_mic_status.config(text=f"Input mikrofon tidak tersedia: {error}")
            return
        options = {}
        for index, device in enumerate(devices):
            if int(device.get("max_input_channels", 0)) < 1:
                continue
            try:
                host = sd.query_hostapis(device["hostapi"])["name"]
            except Exception:
                host = "Audio"
            label = f"{device['name']} · {host} · {index}"
            options[label] = index
        self.broadcast_mic_devices = options
        labels = list(options)
        self.broadcast_mic_choice["values"] = labels
        if labels:
            preferred = self.store.get_setting("broadcast_microphone", "")
            try:
                default_input = sd.default.device[0]
            except Exception:
                default_input = None
            default_label = next((label for label, index in options.items() if index == default_input), labels[0])
            selected = next((label for label in labels if label == preferred), default_label)
            self.broadcast_mic_choice.set(selected)
            self.broadcast_mic_status.config(text=f"{len(labels)} input mikrofon tersedia.")
        else:
            self.broadcast_mic_choice.set("")
            self.broadcast_mic_status.config(text="Tidak ada input mikrofon yang terdeteksi.")

    def toggle_broadcast_microphone(self):
        self.broadcast_mic_started_for_output = False
        if self.broadcast_mic_stream is not None:
            try:
                self.broadcast_mic_stream.stop()
                self.broadcast_mic_stream.close()
            except Exception:
                pass
            self.broadcast_mic_stream = None
            self.broadcast_mic_meter["value"] = 0
            self.broadcast_mic_button.config(text="Aktifkan Mikrofon")
            self.broadcast_mic_status.config(text="Mikrofon tidak sedang diuji.")
            return
        try:
            import sounddevice as sd
            device = self.broadcast_mic_devices[self.broadcast_mic_choice.get()]
            sample_rate = int(sd.query_devices(device)["default_samplerate"] or 44100)
            self.store.set_setting("broadcast_microphone", self.broadcast_mic_choice.get())
            last_update = [0.0]

            def on_audio(indata, _frames, _time_info, _status):
                level = min(1.0, float((indata * indata).mean() ** 0.5) * 4.0)
                now = time.monotonic()
                if now - last_update[0] >= 0.12:
                    last_update[0] = now
                    try:
                        self.after(0, lambda value=level: self.broadcast_mic_meter.configure(value=value * 100))
                    except tk.TclError:
                        pass

            self.broadcast_mic_stream = sd.InputStream(device=device, channels=1,
                                                       samplerate=sample_rate, blocksize=1024,
                                                       callback=on_audio)
            self.broadcast_mic_stream.start()
            self.broadcast_mic_button.config(text="Matikan Mikrofon")
            self.broadcast_mic_status.config(text="Input mikrofon aktif. Pilih perangkat ini juga di aplikasi rapat.")
        except Exception as error:
            self.broadcast_mic_stream = None
            messagebox.showerror("Mikrofon tidak dapat diaktifkan", str(error), parent=self)
            self.broadcast_mic_status.config(text="Mikrofon gagal dibuka. Periksa perangkat dan izin Windows.")

    def start_broadcast(self):
        if self.broadcast_thread and self.broadcast_thread.is_alive():
            return
        source_mode = self.broadcast_source_choice.get()
        content_mode = source_mode in {"Item Susunan", "Kamera + Item Susunan"}
        camera_enabled = source_mode in {"Kamera", "Kamera + Item Susunan"}
        mode_labels = {"Kamera 1 penuh": "primary", "Kamera 2 penuh": "secondary",
                       "Dua kamera 50:50": "split", "Kamera penuh": "primary",
                       "Item penuh": "content", "Kamera + Item 50:50": "camera_content_split",
                       "Lagu Overlay Kamera": "song_overlay"}
        mode = mode_labels.get(self.broadcast_mode_choice.get(), "primary")
        primary_source = secondary_source = None
        if content_mode:
            if self.broadcast_content_capture is None and not self.broadcast_content_frames:
                self._broadcast_item_selected()
            if self.broadcast_content_capture is None and not self.broadcast_content_frames:
                messagebox.showinfo("Pilih konten", "Pilih item susunan yang bisa ditampilkan terlebih dahulu.", parent=self)
                return
        if camera_enabled:
            primary_label = self.broadcast_camera1_choice.get()
            secondary_label = (self.broadcast_camera2_choice.get()
                               if source_mode == "Kamera" else "(Tidak ada)")
            try:
                primary_source = self.camera_devices[primary_label]
            except (KeyError, tk.TclError):
                messagebox.showinfo("Pilih kamera", "Tekan Cari dan pilih Kamera 1 terlebih dahulu.", parent=self)
                return
            secondary_source = self.camera_devices.get(secondary_label) if secondary_label != "(Tidak ada)" else None
            if mode in {"secondary", "split"} and secondary_source is None:
                messagebox.showinfo("Pilih kamera kedua", "Pilih Kamera 2 untuk mode ini.", parent=self)
                return
            if secondary_source is not None and secondary_source == primary_source:
                messagebox.showwarning("Input sama", "Pilih dua input kamera yang berbeda.", parent=self)
                return
        mic_started_here = self.broadcast_mic_stream is None
        if self.broadcast_mic_stream is None:
            try:
                import sounddevice as sd  # noqa: F401
                if self.broadcast_mic_choice.get():
                    self.toggle_broadcast_microphone()
            except ImportError:
                messagebox.showerror("Mikrofon belum tersedia",
                                     "Pasang dependensi sounddevice dari requirements.txt terlebih dahulu.", parent=self)
                return
        if self.broadcast_mic_choice.get() and self.broadcast_mic_stream is None:
            return

        live_running = bool(self.live_window and self.live_window.winfo_exists())
        try:
            if not camera_enabled:
                self.broadcast_capture = None
                self.broadcast_second_capture = None
            elif live_running:
                if primary_source != self.live_source_current:
                    raise RuntimeError("Kamera 1 pada proyeksi live berbeda. Pilih sumber yang sama atau hentikan proyeksi.")
                if secondary_source is not None and secondary_source != self.live_second_source_current:
                    raise RuntimeError("Kamera 2 pada proyeksi live berbeda. Pilih sumber yang sama atau hentikan proyeksi.")
                self.broadcast_capture = self.live_window.capture.acquire()
                if secondary_source is not None:
                    active_secondary = self.live_window.second_capture
                    if active_secondary is None:
                        raise RuntimeError("Proyeksi live belum memakai Kamera 2. Hentikan lalu mulai proyeksi dengan input kedua.")
                    self.broadcast_second_capture = active_secondary.acquire()
            else:
                self.broadcast_capture = VideoCaptureSource(primary_source)
                if secondary_source is not None:
                    self.broadcast_second_capture = VideoCaptureSource(secondary_source)
        except Exception as error:
            if self.broadcast_capture:
                self.broadcast_capture.close()
                self.broadcast_capture = None
            if self.broadcast_second_capture:
                self.broadcast_second_capture.close()
                self.broadcast_second_capture = None
            if mic_started_here and self.broadcast_mic_stream is not None:
                self.toggle_broadcast_microphone()
            self.broadcast_capture = None
            self.broadcast_second_capture = None
            messagebox.showerror("Kamera tidak siap", str(error), parent=self)
            return
        self.broadcast_content_mode = content_mode
        self.broadcast_camera_enabled = camera_enabled
        self.broadcast_output_mode = mode
        self.broadcast_primary_source = primary_source
        self.broadcast_secondary_source = secondary_source
        self.broadcast_camera_stopped = False
        self.broadcast_mic_started_for_output = mic_started_here and self.broadcast_mic_stream is not None
        self.broadcast_stop_event = threading.Event()
        self.broadcast_start_button.configure(state="disabled")
        self.broadcast_camera_stop_button.configure(state="disabled")
        self.broadcast_stop_button.configure(state="normal")
        self.broadcast_status.configure(text="Menyambungkan ke OBS Virtual Camera…")
        self._broadcast_set_source_controls()
        self.broadcast_thread = threading.Thread(target=self._broadcast_video_loop,
                                                 args=(self.broadcast_stop_event,), daemon=True)
        self.broadcast_thread.start()

    def _broadcast_video_loop(self, stop_event):
        error_message = None
        try:
            import pyvirtualcam
            fps, width, height = 25, 1280, 720
            with pyvirtualcam.Camera(width=width, height=height, fps=fps,
                                     fmt=pyvirtualcam.PixelFormat.BGR, backend="obs") as camera:
                self.after(0, lambda: self._broadcast_output_ready(camera.device, width, height, fps))
                while not stop_event.is_set():
                    if self.broadcast_camera_stopped:
                        import numpy as np
                        frame = np.zeros((height, width, 3), dtype=np.uint8)
                    elif self.broadcast_content_mode and self.broadcast_camera_enabled:
                        content = (self.broadcast_content_capture.latest_frame()
                                   if self.broadcast_content_capture else self.broadcast_content_frame)
                        primary = self.broadcast_capture.latest_frame() if self.broadcast_capture else None
                        frame = self._broadcast_compose_frame(primary, None, content, (width, height))
                    elif self.broadcast_content_mode:
                        content = (self.broadcast_content_capture.latest_frame()
                                   if self.broadcast_content_capture else self.broadcast_content_frame)
                        if content is None:
                            import numpy as np
                            frame = np.zeros((height, width, 3), dtype=np.uint8)
                        else:
                            frame = self._broadcast_compose_frame(None, None, content, (width, height))
                    else:
                        primary = self.broadcast_capture.latest_frame() if self.broadcast_capture else None
                        secondary = self.broadcast_second_capture.latest_frame() if self.broadcast_second_capture else None
                        if primary is None:
                            import numpy as np
                            frame = np.zeros((height, width, 3), dtype=np.uint8)
                        else:
                            frame = self._broadcast_compose_frame(primary, secondary, None, (width, height))
                    camera.send(frame)
                    camera.sleep_until_next_frame()
        except Exception as error:
            error_message = str(error)
        finally:
            if self.broadcast_capture:
                self.broadcast_capture.close()
                self.broadcast_capture = None
            if self.broadcast_second_capture:
                self.broadcast_second_capture.close()
                self.broadcast_second_capture = None
            if self.broadcast_content_capture:
                self.broadcast_content_capture.close()
                self.broadcast_content_capture = None
            try:
                self.after(0, lambda message=error_message: self._broadcast_video_finished(message))
            except tk.TclError:
                pass

    def _broadcast_output_ready(self, device, width, height, fps):
        if self._closing:
            return
        self.broadcast_status.configure(text=f"Aktif · {device} · {width}×{height} · {fps} FPS")
        self.broadcast_camera_stop_button.configure(state="normal")

    def stop_broadcast_camera(self):
        if not self.broadcast_thread or not self.broadcast_thread.is_alive():
            messagebox.showinfo(
                "Output belum aktif",
                "Mulai output siaran terlebih dahulu. Tombol ini akan mengirim layar hitam "
                "tanpa memutus siaran saat output sudah aktif.", parent=self)
            return
        if self.broadcast_camera_stopped:
            messagebox.showinfo("Kamera sudah dihentikan", "Output hitam masih dikirim ke rapat.", parent=self)
            return
        if not messagebox.askyesno(
                "Hentikan kamera Broadcast?",
                "Sumber siaran akan dihentikan dan layar hitam dikirim ke OBS Virtual Camera. "
                "Output tetap tersambung ke rapat sampai Anda memilih Hentikan Output.\n\nLanjutkan?",
                parent=self):
            return
        self.broadcast_camera_stopped = True
        if self.broadcast_capture:
            self.broadcast_capture.close()
            self.broadcast_capture = None
        if self.broadcast_second_capture:
            self.broadcast_second_capture.close()
            self.broadcast_second_capture = None
        if self.broadcast_content_capture:
            self.broadcast_content_capture.close()
            self.broadcast_content_capture = None
        self.broadcast_content_frame = None
        self.broadcast_camera_stop_button.configure(
            text="▶  Mulai Kembali Sumber", state="normal",
            command=self.resume_broadcast_source)
        self.broadcast_status.configure(text="Sumber dihentikan · layar hitam dikirim; output rapat tetap aktif.")
        self.broadcast_preview.configure(bg="#000000", image="", text="")
        self.broadcast_preview_photo = None

    def resume_broadcast_source(self):
        """Reopen the selected sources while keeping the virtual-camera output alive."""
        if not self.broadcast_thread or not self.broadcast_thread.is_alive():
            self.broadcast_camera_stop_button.configure(
                text="■  Hentikan Sumber · Tampilkan Hitam", command=self.stop_broadcast_camera)
            self.start_broadcast()
            return
        if not self.broadcast_camera_stopped:
            return

        content_mode = self.broadcast_content_mode
        camera_enabled = self.broadcast_camera_enabled
        primary_source = self.broadcast_primary_source
        secondary_source = self.broadcast_secondary_source
        if content_mode:
            # Rebuild the selected slide or reopen its video after the stop action cleared it.
            self._broadcast_item_selected()

        self.broadcast_camera_stop_button.configure(text="Menyambungkan sumber…", state="disabled")
        self.broadcast_status.configure(text="Mengaktifkan kembali sumber siaran…")
        if not camera_enabled:
            self.broadcast_camera_stopped = False
            self._broadcast_source_resumed()
            return

        def open_sources():
            primary = secondary = None
            try:
                if self.live_window and self.live_window.winfo_exists() and primary_source == self.live_source_current:
                    primary = self.live_window.capture.acquire()
                else:
                    primary = VideoCaptureSource(primary_source)
                if secondary_source is not None:
                    if (self.live_window and self.live_window.winfo_exists() and
                            secondary_source == self.live_second_source_current and
                            self.live_window.second_capture is not None):
                        secondary = self.live_window.second_capture.acquire()
                    else:
                        secondary = VideoCaptureSource(secondary_source)
            except Exception as error:
                if primary is not None:
                    primary.close()
                if secondary is not None:
                    secondary.close()
                try:
                    self.after(0, lambda message=str(error): self._broadcast_source_resume_failed(message))
                except tk.TclError:
                    pass
                return

            def apply_sources():
                if self._closing or not (self.broadcast_thread and self.broadcast_thread.is_alive()):
                    primary.close()
                    if secondary is not None:
                        secondary.close()
                    return
                self.broadcast_capture = primary
                self.broadcast_second_capture = secondary
                self.broadcast_camera_stopped = False
                self._broadcast_source_resumed()
            try:
                self.after(0, apply_sources)
            except tk.TclError:
                primary.close()
                if secondary is not None:
                    secondary.close()

        threading.Thread(target=open_sources, daemon=True).start()

    def _broadcast_source_resumed(self):
        self.broadcast_camera_stop_button.configure(
            text="■  Hentikan Sumber · Tampilkan Hitam", state="normal",
            command=self.stop_broadcast_camera)
        self.broadcast_status.configure(text="Sumber siaran aktif kembali.")
        self._broadcast_refresh_preview()

    def _broadcast_source_resume_failed(self, error):
        self.broadcast_camera_stopped = True
        self.broadcast_camera_stop_button.configure(
            text="▶  Coba Lagi Aktifkan Sumber", state="normal",
            command=self.resume_broadcast_source)
        self.broadcast_status.configure(text="Sumber belum dapat dibuka kembali; output tetap hitam.")
        messagebox.showerror("Sumber tidak dapat diaktifkan", error, parent=self)

    def _broadcast_video_finished(self, error=None):
        self.broadcast_thread = None
        self.broadcast_stop_event = None
        self.broadcast_primary_source = None
        self.broadcast_secondary_source = None
        if self._closing:
            return
        self.broadcast_start_button.configure(state="normal")
        self.broadcast_source_choice.configure(state="readonly")
        self.broadcast_camera_stop_button.configure(
            text="■  Hentikan Sumber · Tampilkan Hitam", state="normal",
            command=self.stop_broadcast_camera)
        self._broadcast_set_source_controls()
        self.broadcast_stop_button.configure(state="disabled")
        self.broadcast_camera_stopped = False
        if error and getattr(self, "broadcast_mic_started_for_output", False):
            self.toggle_broadcast_microphone()
        self.broadcast_mic_started_for_output = False
        if error:
            self.broadcast_status.configure(text="Output kamera berhenti.")
            if "camera" in error.casefold() or "obs" in error.casefold() or "virtual" in error.casefold():
                messagebox.showerror(
                    "OBS Virtual Camera tidak tersedia",
                    "Pastikan OBS Virtual Camera sudah terpasang. Jika kamera OBS belum pernah diaktifkan, buka OBS, klik Start Virtual Camera sekali, lalu coba lagi.\n\n"
                    + error, parent=self)
            else:
                messagebox.showerror("Output kamera berhenti", error, parent=self)
        else:
            self.broadcast_status.configure(text="Siap · OBS Virtual Camera sudah dihentikan.")

    def stop_broadcast(self):
        self._broadcast_camera_switch_generation += 1
        if self.broadcast_stop_event:
            self.broadcast_stop_event.set()
        if hasattr(self, "broadcast_camera_stop_button"):
            self.broadcast_camera_stop_button.configure(
                text="■  Hentikan Sumber · Tampilkan Hitam", state="disabled",
                command=self.stop_broadcast_camera)
        if self.broadcast_mic_stream is not None:
            try:
                self.broadcast_mic_stream.stop()
                self.broadcast_mic_stream.close()
            except Exception:
                pass
            self.broadcast_mic_stream = None
            self.broadcast_mic_button.configure(text="Aktifkan Mikrofon")
            self.broadcast_mic_meter["value"] = 0
            self.broadcast_mic_status.configure(text="Mikrofon tidak sedang diuji.")
        self.broadcast_mic_started_for_output = False

    def _update_live_navigation(self):
        for slides, index, label, previous, following, noun in (
                (self.live_song_slides, self.live_song_index, self.live_song_position,
                 self.live_song_previous, self.live_song_next, "lirik"),
                (self.live_verse_slides, self.live_verse_index, self.live_verse_position,
                 self.live_verse_previous, self.live_verse_next, "ayat")):
            total = len(slides)
            if total and 0 <= index < total:
                label.config(text=f"{noun.title()}  {index + 1} / {total}")
                previous.state(["!disabled"] if index > 0 else ["disabled"])
                following.state(["!disabled"] if index < total - 1 else ["disabled"])
            else:
                label.config(text=f"Belum ada {noun}")
                previous.state(["disabled"])
                following.state(["disabled"])

    def _scroll_camera_panel(self, event):
        current = self.tabs.select()
        if current == str(self.camera_tab) and hasattr(self, "overlay_canvas"):
            canvas = self.overlay_canvas
        elif current == str(self.verses_tab) and hasattr(self, "verses_canvas"):
            canvas = self.verses_canvas
        elif current == str(self.broadcast_tab) and hasattr(self, "broadcast_controls_canvas"):
            canvas = self.broadcast_controls_canvas
        else:
            return None
        left, top = canvas.winfo_rootx(), canvas.winfo_rooty()
        if not (left <= event.x_root <= left + canvas.winfo_width()
                and top <= event.y_root <= top + canvas.winfo_height()):
            return None
            
        # Tangani KeyError 'popdown' jika mouse berada di atas dropdown Combobox
        try:
            target_widget = self.winfo_containing(event.x_root, event.y_root)
            if isinstance(target_widget, (tk.Text, ttk.Treeview)):
                return None
        except KeyError:
            return None

        step = int(-event.delta / 120) if event.delta else -1
        canvas.yview_scroll(step or -1, "units")
        return "break"

    def _on_tab_changed(self, _event=None):
        if hasattr(self, "display_choice") and self.tabs.select() == str(self.camera_tab):
            self.refresh_display_choices()

    def _close_main_window(self):
        if self._closing:
            return
        self._closing = True
        self.stop_broadcast()
        if self.broadcast_thread and self.broadcast_thread.is_alive():
            # Keep Tk's event loop running while output releases its camera.
            self.after(50, self._finish_main_window_close)
            return
        self._finish_main_window_close()

    def _finish_main_window_close(self):
        if self.broadcast_thread and self.broadcast_thread.is_alive():
            self.after(50, self._finish_main_window_close)
            return
        if self.live_window and self.live_window.winfo_exists():
            self.live_window.close_projector()
            self.live_window = None
        self.destroy()

    def _display_choice_selected(self, _event=None):
        index = self.display_choice.current()
        if 0 <= index < len(self.display_options):
            self._display_preference = self.display_options[index]

    def refresh_display_choices(self):
        """Re-read the monitors so a screen connected after start-up is usable."""
        displays = available_displays()
        labels = [f"Layar {i + 1} · {monitor[2]}×{monitor[3]}"
                  for i, monitor in enumerate(displays)]
        if not labels:
            return
        self.display_options = displays
        self.display_labels = labels
        self.display_choice["values"] = labels
        preference = getattr(self, "_display_preference", None)
        if preference in displays:
            index = displays.index(preference)
        elif len(labels) > 1:
            # No explicit choice: prefer the extended screen over the operator screen.
            index = 1
        else:
            index = 0
        self.display_choice.current(index)

    def refresh_camera_devices(self):
        if self._closing or self._camera_scan_in_progress:
            return
        try:
            import cv2
        except ImportError:
            self.camera_choice["values"] = []
            self.camera_status.config(text="OpenCV belum terpasang. Pasang dependensi dari requirements.txt untuk mengaktifkan kamera.")
            return
        self.camera_status.config(text="Sedang mencari kamera dan capture device…")
        self._camera_scan_in_progress = True

        def scan():
            found = []
            # Sahabat Ibadah currently accepts at most two camera inputs. Probing
            # indexes 2–7 created OpenCV warnings for devices the UI cannot use.
            get_log_level = getattr(cv2, "getLogLevel", None)
            set_log_level = getattr(cv2, "setLogLevel", None)
            previous_log_level = get_log_level() if get_log_level and set_log_level else None
            try:
                if previous_log_level is not None:
                    set_log_level(0)  # OpenCV ERROR level; expected missing-index probes stay quiet.
                for index in range(2):
                    capture = cv2.VideoCapture(index, cv2.CAP_DSHOW)
                    try:
                        if capture.isOpened():
                            ok, _frame = capture.read()
                            if ok:
                                found.append(index)
                    finally:
                        capture.release()
            finally:
                if previous_log_level is not None:
                    set_log_level(previous_log_level)
            try:
                self.after(0, lambda: self._camera_scan_finished(found))
            except tk.TclError:
                pass

        threading.Thread(target=scan, daemon=True).start()

    def _camera_scan_finished(self, camera_ids):
        self._camera_scan_in_progress = False
        if self._closing:
            return
        self.camera_devices = {f"Kamera / Input {index}": index for index in camera_ids}
        choices = list(self.camera_devices)
        if self.video_input_path:
            video_label = f"Video · {Path(self.video_input_path).name}"
            self.camera_devices[video_label] = self.video_input_path
            choices.append(video_label)
        self.camera_choice["values"] = choices
        self.camera2_choice["values"] = ["(Tidak ada)"] + choices
        self.broadcast_camera1_choice["values"] = choices
        self.broadcast_camera2_choice["values"] = ["(Tidak ada)"] + choices
        self.camera2_choice.current(0)
        self.broadcast_camera2_choice.current(0)
        self._broadcast_update_mode_options()
        if choices:
            self.camera_choice.current(0)
            self.broadcast_camera1_choice.current(0)
            self.camera_status.config(text=f"{len(camera_ids)} kamera/capture device ditemukan. Perangkat video eksternal akan muncul sebagai input kamera.")
        else:
            self.camera_status.config(text="Tidak ada kamera ditemukan. Coba tombol Cari atau pilih file video.")

    def choose_video_file(self):
        path = filedialog.askopenfilename(parent=self, title="Pilih file video",
                                          filetypes=[("Video", "*.mp4 *.avi *.mov *.mkv *.wmv *.m4v"),
                                                     ("Semua file", "*.*")])
        if not path:
            return
        self.video_input_path = path
        video_label = f"Video · {Path(path).name}"
        self.camera_devices[video_label] = path
        values = list(self.camera_choice["values"])
        if video_label not in values:
            values.append(video_label)
        self.camera_choice["values"] = values
        self.camera2_choice["values"] = ["(Tidak ada)"] + values
        self.broadcast_camera1_choice["values"] = values
        self.broadcast_camera2_choice["values"] = ["(Tidak ada)"] + values
        self.broadcast_camera2_choice.current(0)
        self._broadcast_update_mode_options()
        self.camera_choice.set(video_label)
        self.broadcast_camera1_choice.set(video_label)
        self.camera_status.config(text="File video siap diputar berulang sebagai sumber live.")

    def start_live_projection(self):
        if self.live_window and self.live_window.winfo_exists():
            self.live_window.lift()
            return
        try:
            source = self.camera_devices[self.camera_choice.get()]
        except (KeyError, tk.TclError):
            messagebox.showinfo("Pilih sumber video", "Cari dan pilih kamera, atau pilih file video terlebih dahulu.")
            return
        try:
            self.refresh_display_choices()
            try:
                display_index = self.display_labels.index(self.display_choice.get())
            except ValueError:
                display_index = 1 if len(self.display_options) > 1 else 0
            display = self.display_options[display_index]
            second_label = self.camera2_choice.get()
            second_source = self.camera_devices.get(second_label) if second_label != "(Tidak ada)" else None
            if second_label != "(Tidak ada)" and second_source == self.camera_devices.get(self.camera_choice.get()):
                messagebox.showwarning("Pilih input berbeda", "Kamera pertama dan kedua harus berupa input yang berbeda.", parent=self)
                return
            self.live_window = LiveProjector(self, source, display,
                                             lambda: self.live_text.get("1.0", "end-1c"),
                                             self._live_layout_state,
                                             self.step_live_content,
                                             self.live_image_path, self.live_image_mode,
                                             second_source=second_source,
                                             camera_mode=self.live_camera_mode,
                                             overlay_text_color=lambda: self.live_text_color,
                                             overlay_stroke=self._live_stroke_style,
                                             overlay_background=lambda: self.live_text_background)
            self.live_source_current = source
            self.live_second_source_current = second_source
        except (RuntimeError, IndexError) as error:
            messagebox.showerror("Tidak dapat memulai live", str(error))

    def stop_live_projection(self):
        if self.live_window and self.live_window.winfo_exists():
            self.live_window.close_projector()
            self.live_window = None
            self.live_source_current = None
            self.live_second_source_current = None
            self.camera_status.config(text="Kamera / video telah dihentikan.")
        else:
            self.live_window = None
            self.live_source_current = None
            self.live_second_source_current = None
            self.camera_status.config(text="Kamera / video sedang tidak berjalan.")

    def set_live_camera_mode(self, mode):
        if mode in {"secondary", "split"} and self.camera2_choice.get() in {"", "(Tidak ada)"}:
            messagebox.showinfo("Pilih kamera kedua", "Pilih input Kamera kedua terlebih dahulu.", parent=self)
            return
        self.live_camera_mode = mode
        if self.live_window and self.live_window.winfo_exists():
            try:
                self.live_window.set_camera_mode(mode)
            except RuntimeError as error:
                messagebox.showinfo("Tampilan kamera", str(error), parent=self)
                return
        labels = {"primary": "Kamera 1 tampil penuh", "secondary": "Kamera 2 tampil penuh",
                  "split": "Dua kamera tampil berdampingan 50:50", "hidden": "Kamera disembunyikan"}
        self.camera_status.config(text=labels[mode])

    def choose_live_media(self):
        path = filedialog.askopenfilename(
            parent=self, title="Pilih gambar atau presentasi PowerPoint untuk tampilan live",
            filetypes=[("Gambar atau PowerPoint", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff *.pptx *.ppt"),
                       ("Gambar", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"),
                       ("PowerPoint", "*.pptx *.ppt"),
                       ("Semua file", "*.*")])
        if not path:
            return
        if Path(path).suffix.casefold() in {".pptx", ".ppt"}:
            self.load_live_presentation(path)
            return
        try:
            if self.live_window and self.live_window.winfo_exists():
                self.live_window.set_still_image(path)
            else:
                import cv2
                try:
                    valid_image = cv2.imread(path, cv2.IMREAD_COLOR)
                except Exception as error:
                    raise RuntimeError("File gambar tidak dapat dibaca oleh OpenCV.") from error
                if valid_image is None:
                    raise RuntimeError("File gambar tidak dapat dibuka. Gunakan PNG, JPG, BMP, atau WebP.")
            self.live_media_kind = "image"
            self.live_presentation_path = None
            self.live_presentation_slides = []
            self.live_presentation_index = -1
            self._hide_slide_controls()
            self.live_image_path = path
            if self.live_image_mode != "camera":
                self.live_image_status.config(text=self._live_media_status_text())
            else:
                self.live_image_status.config(text=f"Gambar siap: {Path(path).name}")
        except Exception as error:
            messagebox.showerror("Gambar tidak dapat dibuka", str(error), parent=self)

    def load_live_presentation(self, path):
        """Render a PPT/PPTX deck and show its slides as live media."""
        self.refresh_display_choices()
        try:
            display = self.display_options[self.display_choice.current()]
        except (IndexError, tk.TclError):
            display = available_displays()[0]
        size = (max(640, display[2]), max(480, display[3]))
        self.live_image_status.config(text=f"Merender presentasi… {Path(path).name}")
        self.update_idletasks()
        try:
            slides = render_powerpoint_slides(path, size)
        except PowerPointRenderError as error:
            self.live_image_status.config(text="Presentasi tidak dapat dirender.")
            messagebox.showerror("PowerPoint tidak dapat ditampilkan", str(error), parent=self)
            return
        if not slides:
            self.live_image_status.config(text="Presentasi tidak memiliki slide.")
            return
        self.live_media_kind = "presentation"
        self.live_presentation_path = path
        self.live_presentation_slides = slides
        self.live_presentation_index = 0
        self.live_image_path = slides[0]
        if self.live_window and self.live_window.winfo_exists():
            try:
                self.live_window.set_still_image(slides[0])
            except RuntimeError as error:
                messagebox.showerror("Slide tidak dapat ditampilkan", str(error), parent=self)
                return
        self._refresh_slide_choices()
        self._show_slide_controls()
        if self.live_image_mode != "camera":
            self.live_image_status.config(text=self._live_media_status_text())
        else:
            self.live_image_status.config(
                text=f"Presentasi siap: {Path(path).name} · {len(slides)} slide")

    def _show_slide_controls(self):
        if not self.slide_actions.winfo_manager():
            self.slide_actions.pack(fill="x", pady=(6, 0), after=self.image_actions)

    def _hide_slide_controls(self):
        if self.slide_actions.winfo_manager():
            self.slide_actions.pack_forget()

    def _refresh_slide_choices(self):
        total = len(self.live_presentation_slides)
        self.live_slide_choice["values"] = [f"Slide {index + 1}" for index in range(total)]
        if 0 <= self.live_presentation_index < total:
            self.live_slide_choice.current(self.live_presentation_index)
            self.live_slide_status.config(
                text=f"{self.live_presentation_index + 1}/{total} · "
                     f"{Path(self.live_presentation_path).name}")
        else:
            self.live_slide_choice.set("")
            self.live_slide_status.config(text="")

    def set_live_presentation_slide(self, index):
        if not (0 <= index < len(self.live_presentation_slides)):
            return
        path = self.live_presentation_slides[index]
        if self.live_window and self.live_window.winfo_exists():
            try:
                self.live_window.set_still_image(path)
            except RuntimeError as error:
                messagebox.showerror("Slide tidak dapat ditampilkan", str(error), parent=self)
                return
        self.live_presentation_index = index
        self.live_image_path = path
        self.live_slide_choice.current(index)
        self._refresh_slide_choices()
        if self.live_image_mode != "camera":
            self.live_image_status.config(text=self._live_media_status_text())

    def step_live_presentation(self, offset):
        if not self.live_presentation_slides:
            return
        target = max(0, min(self.live_presentation_index + offset,
                            len(self.live_presentation_slides) - 1))
        self.set_live_presentation_slide(target)

    def select_live_presentation_slide(self, _event=None):
        index = self.live_slide_choice.current()
        if index >= 0:
            self.set_live_presentation_slide(index)

    def _live_media_status_text(self):
        if self.live_image_mode == "split":
            text = "Mode kamera dan media berdampingan (50/50)."
        elif self.live_image_mode == "image":
            text = "Mode media layar penuh; kamera tetap berjalan di latar."
        else:
            text = "Media disembunyikan; kamera tampil kembali."
        if self.live_media_kind == "presentation" and self.live_presentation_slides:
            text += (f"  ·  slide {self.live_presentation_index + 1}/"
                     f"{len(self.live_presentation_slides)}")
        elif self.live_image_path:
            text += f"  ·  {Path(self.live_image_path).name}"
        return text

    def set_live_image_mode(self, mode):
        if mode != "camera" and not self.live_image_path:
            messagebox.showinfo("Pilih media", "Pilih gambar atau presentasi terlebih dahulu.", parent=self)
            return
        try:
            if self.live_window and self.live_window.winfo_exists():
                self.live_window.set_image_mode(mode)
            self.live_image_mode = mode
            self.live_image_status.config(text=self._live_media_status_text())
        except RuntimeError as error:
            messagebox.showerror("Tampilan media tidak dapat diubah", str(error), parent=self)

    def show_live_split(self):
        self.set_live_image_mode("split")

    def show_live_image_full(self):
        self.set_live_image_mode("image")

    def hide_live_image(self):
        self.set_live_image_mode("camera")

    def load_live_song(self):
        selected = self.schedule_tree.selection()
        if not selected:
            messagebox.showinfo("Pilih lagu", "Pilih sebuah lagu pada tab Susunan Ibadah terlebih dahulu.")
            self.tabs.select(self.schedule_tab)
            return
        row = self.store.db.execute("SELECT * FROM schedule_items WHERE id=? AND schedule_id=?",
                                    (int(selected[0]), self.active_schedule["id"])).fetchone()
        if not row or row["kind"] != "Lagu":
            messagebox.showinfo("Item bukan lagu", "Pilih item lagu pada susunan ibadah untuk memuat liriknya.")
            return
        self.live_song_slides = self.make_slides(row["content"])
        self.live_song_index = 0 if self.live_song_slides else -1
        if self.live_song_index < 0:
            messagebox.showinfo("Lirik kosong", "Item lagu ini belum memiliki lirik.")
            return
        self.live_mode = "song"
        self._show_live_song_slide()
        self._update_live_navigation()
        self.tabs.select(self.camera_tab)

    def load_live_verse(self):
        rows = self.store.db.execute(
            "SELECT title, content FROM schedule_items WHERE schedule_id=? AND kind='Ayat' ORDER BY position, id",
            (self.active_schedule["id"],)).fetchall()
        self.live_verse_slides = []
        for row in rows:
            for part in self.make_slides(row["content"]):
                first_line = part.splitlines()[0].strip() if part.splitlines() else ""
                if re.fullmatch(r".+\s+\d+:\d+", first_line):
                    self.live_verse_slides.append(part)
                else:
                    self.live_verse_slides.append(f"{row['title']}\n{part}")
        self.live_verse_index = 0 if self.live_verse_slides else -1
        if self.live_verse_index < 0:
            messagebox.showinfo("Ayat belum tersedia", "Tambahkan ayat ke susunan ibadah terlebih dahulu.")
            self.tabs.select(self.verses_tab)
            return
        self.live_mode = "verse"
        self._refresh_live_verse_choices()
        self._show_live_verse_slide()
        self._update_live_navigation()
        self.tabs.select(self.camera_tab)

    def _refresh_live_verse_choices(self):
        choices = []
        for index, slide in enumerate(self.live_verse_slides, start=1):
            reference = slide.splitlines()[0].strip() if slide.splitlines() else f"Ayat {index}"
            choices.append(f"{index}. {reference}")
        self.live_verse_choice["values"] = choices

    def select_live_verse(self, _event=None):
        index = self.live_verse_choice.current()
        if 0 <= index < len(self.live_verse_slides):
            self.live_mode = "verse"
            self.live_verse_index = index
            self._show_live_verse_slide()
            self._update_live_navigation()

    def _show_live_verse_slide(self):
        if 0 <= self.live_verse_index < len(self.live_verse_slides):
            self.live_text.delete("1.0", "end")
            self.live_text.insert("1.0", self.live_verse_slides[self.live_verse_index])
            if self.live_verse_choice.current() != self.live_verse_index:
                self.live_verse_choice.current(self.live_verse_index)

    def step_live_verse(self, offset):
        if not self.live_verse_slides:
            messagebox.showinfo("Ayat belum dimuat", "Tekan Muat Ayat dari Susunan untuk menampilkan ayat dari susunan ibadah.")
            return
        self.live_mode = "verse"
        self.live_verse_index = max(0, min(self.live_verse_index + offset, len(self.live_verse_slides) - 1))
        self._show_live_verse_slide()
        self._update_live_navigation()

    def step_live_content(self, offset):
        if self.live_mode == "verse":
            self.step_live_verse(offset)
        elif self.live_song_slides:
            self.step_live_song(offset)
        elif self.live_presentation_slides:
            # No lyrics loaded: arrow keys/space step through the live presentation.
            self.step_live_presentation(offset)

    def _show_live_song_slide(self):
        if 0 <= self.live_song_index < len(self.live_song_slides):
            self.live_text.delete("1.0", "end")
            self.live_text.insert("1.0", self.live_song_slides[self.live_song_index])

    def step_live_song(self, offset):
        if not self.live_song_slides:
            messagebox.showinfo("Lirik belum dimuat", "Pilih lagu di Susunan Ibadah lalu tekan Muat Lagu Terpilih.")
            return
        self.live_song_index = max(0, min(self.live_song_index + offset, len(self.live_song_slides) - 1))
        self.live_mode = "song"
        self._show_live_song_slide()
        self._update_live_navigation()

    def clear_live_text(self):
        self.live_text.delete("1.0", "end")
        self.live_song_slides = []
        self.live_song_index = -1
        self.live_verse_slides = []
        self.live_verse_index = -1
        self.live_verse_choice["values"] = []
        self.live_verse_choice.set("")
        self.live_mode = None
        self._update_live_navigation()

    def choose_live_text_color(self):
        chosen = colorchooser.askcolor(color=self.live_text_color,
                                       title="Warna teks overlay live", parent=self)
        if chosen and chosen[1]:
            self.live_text_color = chosen[1]
            self.live_text_color_button.config(text=f"Warna Teks · {self.live_text_color}")

    def _live_stroke_style(self):
        """The live projector calls this every frame, so changes apply instantly."""
        if self.live_stroke_enabled:
            return self.live_stroke_color, self.live_stroke_width
        return None, 0

    def open_live_stroke_dialog(self):
        self._live_stroke_dialog = self._reuse_dialog(
            self._live_stroke_dialog,
            lambda: TextStrokeDialog(self, "Garis Tepi Teks Live", self.live_text_color,
                                     self.live_stroke_enabled, self.live_stroke_color,
                                     self.live_stroke_width, self.apply_live_stroke))

    def apply_live_stroke(self, state):
        self.live_stroke_enabled = bool(state["enabled"])
        self.live_stroke_color = state["color"]
        self.live_stroke_width = int(state["width"])
        self.live_stroke_button.config(
            text=f"Garis Tepi · {self.live_stroke_color}" if self.live_stroke_enabled
            else "Garis Tepi Teks…")

    def set_live_text_background(self):
        self.live_text_background = bool(self.live_background_var.get())
        self.store.set_setting("live_text_background", "1" if self.live_text_background else "0")

    @staticmethod
    def _clamp_int(value, low, high, fallback):
        try:
            return max(low, min(high, int(value)))
        except (TypeError, ValueError):
            return fallback

    @staticmethod
    def _clamp_float(value, low, high, fallback):
        try:
            return max(low, min(high, float(value)))
        except (TypeError, ValueError):
            return fallback

    def _load_live_layout_settings(self):
        position = self.store.get_setting("live_overlay_position", "Bawah")
        self.live_text_position = position if position in OVERLAY_POSITIONS else "Bawah"
        align = self.store.get_setting("live_text_align", "Tengah")
        self.live_text_align = align if align in OVERLAY_ALIGNS else "Tengah"
        self.live_text_scale = self._clamp_int(
            self.store.get_setting("live_text_scale", "100"), 40, 150, 100)
        self.live_margin_left = self._clamp_int(
            self.store.get_setting("live_margin_left", "6"), 0, 45, 6)
        self.live_margin_right = self._clamp_int(
            self.store.get_setting("live_margin_right", "6"), 0, 45, 6)
        self.live_margin_top = self._clamp_int(
            self.store.get_setting("live_margin_top", "6"), 0, 45, 6)
        self.live_margin_bottom = self._clamp_int(
            self.store.get_setting("live_margin_bottom", "6"), 0, 45, 6)
        self.live_offset_x = self._clamp_float(
            self.store.get_setting("live_offset_x", "0"), -80, 80, 0.0)
        self.live_offset_y = self._clamp_float(
            self.store.get_setting("live_offset_y", "0"), -80, 80, 0.0)

    def _live_layout_state(self):
        return {"position": self.live_text_position, "align": self.live_text_align,
                "scale": self.live_text_scale, "left": self.live_margin_left,
                "right": self.live_margin_right, "top": self.live_margin_top,
                "bottom": self.live_margin_bottom,
                "offset_x": self.live_offset_x, "offset_y": self.live_offset_y}

    def set_live_overlay_position(self, _event=None):
        value = self.live_overlay_position.get()
        if value not in OVERLAY_POSITIONS:
            return
        self.live_text_position = value
        self.store.set_setting("live_overlay_position", value)
        self._update_live_layout_status()

    def _update_live_layout_status(self):
        if not hasattr(self, "live_layout_status"):
            return
        status = (f"{self.live_text_position} · {self.live_text_align} · "
                  f"teks {self.live_text_scale}% · margin "
                  f"{self.live_margin_left}/{self.live_margin_right}/"
                  f"{self.live_margin_top}/{self.live_margin_bottom}%")
        if self.live_offset_x or self.live_offset_y:
            status += f" · geser {self.live_offset_x:+.0f}%/{self.live_offset_y:+.0f}%"
        self.live_layout_status.config(text=status)

    def apply_live_layout(self, state):
        self.live_text_position = state["position"]
        self.live_text_align = state["align"]
        self.live_text_scale = int(state["scale"])
        self.live_margin_left = int(state["left"])
        self.live_margin_right = int(state["right"])
        self.live_margin_top = int(state["top"])
        self.live_margin_bottom = int(state["bottom"])
        self.live_offset_x = round(float(state.get("offset_x", 0)), 1)
        self.live_offset_y = round(float(state.get("offset_y", 0)), 1)
        self.live_overlay_position.set(self.live_text_position)
        self.store.set_setting("live_overlay_position", self.live_text_position)
        self.store.set_setting("live_text_align", self.live_text_align)
        self.store.set_setting("live_text_scale", str(self.live_text_scale))
        self.store.set_setting("live_margin_left", str(self.live_margin_left))
        self.store.set_setting("live_margin_right", str(self.live_margin_right))
        self.store.set_setting("live_margin_top", str(self.live_margin_top))
        self.store.set_setting("live_margin_bottom", str(self.live_margin_bottom))
        self.store.set_setting("live_offset_x", f"{self.live_offset_x:g}")
        self.store.set_setting("live_offset_y", f"{self.live_offset_y:g}")
        self._update_live_layout_status()

    def open_live_layout_dialog(self):
        self._live_layout_dialog = self._reuse_dialog(
            self._live_layout_dialog,
            lambda: TextLayoutDialog(self, self.live_text_color,
                                     self.live_stroke_color if self.live_stroke_enabled else None,
                                     self.live_stroke_width, self._live_layout_state(),
                                     self.apply_live_layout))

    def _build_schedule(self):
        pane = ttk.Panedwindow(self.schedule_tab, orient="horizontal")
        pane.pack(fill="both", expand=True)
        left = ttk.Frame(pane, style="Card.TFrame", padding=14)
        right = ttk.Frame(pane, style="Card.TFrame", padding=18)
        pane.add(left, weight=2); pane.add(right, weight=3)
        ttk.Label(left, text="Susunan Ibadah", font=("Segoe UI", 15, "bold"), background="white").pack(anchor="w")
        schedule_bar = ttk.Frame(left, style="Card.TFrame")
        schedule_bar.pack(fill="x", pady=(8, 10))
        self.schedule_choice = ttk.Combobox(schedule_bar, state="readonly", width=24)
        self.schedule_choice.pack(side="left", fill="x", expand=True)
        self.schedule_choice.bind("<<ComboboxSelected>>", self.select_schedule)
        ttk.Button(schedule_bar, text="＋ Jadwal", command=self.new_schedule).pack(side="left", padx=(6, 0))
        schedule_menu_button = ttk.Menubutton(schedule_bar, text="⋯", width=3)
        schedule_menu = tk.Menu(schedule_menu_button, tearoff=0)
        schedule_menu.add_command(label="Gandakan jadwal", command=self.duplicate_schedule)
        schedule_menu.add_command(label="Hapus jadwal ini", command=self.delete_schedule)
        schedule_menu.add_separator()
        schedule_menu.add_command(label="Simpan jadwal ke file…", command=self.save_schedule_file)
        schedule_menu.add_command(label="Buka jadwal dari file…", command=self.open_schedule_file)
        schedule_menu_button["menu"] = schedule_menu
        schedule_menu_button.pack(side="left", padx=(5, 0))
        self.schedule_tree = ttk.Treeview(left, columns=("title", "kind", "operator_note"),
                                          show="headings", selectmode="browse")
        self.schedule_tree.heading("title", text="No. / Urutan")
        self.schedule_tree.heading("kind", text="Jenis")
        self.schedule_tree.heading("operator_note", text="Catatan Operator")
        self.schedule_tree.column("title", width=250, stretch=True)
        self.schedule_tree.column("kind", width=82, stretch=False)
        self.schedule_tree.column("operator_note", width=160, stretch=True)
        self.schedule_tree.pack(fill="both", expand=True)
        ttk.Label(left, text="Seret untuk mengurutkan · Edit untuk menambah catatan operator (privat).", background="white",
                  foreground="#65758b").pack(anchor="w", pady=(4, 0))
        actions = ttk.Frame(left, style="Card.TFrame")
        actions.pack(fill="x", pady=(12, 0))
        add_actions = ttk.Frame(actions, style="Card.TFrame")
        add_actions.pack(fill="x")
        add_menu_button = ttk.Menubutton(add_actions, text="＋ Tambah ▾")
        add_menu = tk.Menu(add_menu_button, tearoff=0)
        add_menu.add_command(label="Lagu…", command=self.add_song_to_schedule)
        add_menu.add_command(label="Ayat…", command=self.add_verse_to_schedule)
        add_menu.add_separator()
        add_menu.add_command(label="Media / gambar / video…", command=self.add_media_to_schedule)
        add_menu.add_command(label="URL website…", command=self.add_url_to_schedule)
        add_menu_button["menu"] = add_menu
        add_menu_button.pack(side="left")
        manage_actions = ttk.Frame(actions, style="Card.TFrame")
        manage_actions.pack(fill="x", pady=(6, 0))
        ttk.Button(manage_actions, text="Hapus dari Susunan", style="Danger.TButton",
                   command=self.remove_schedule_item).pack(side="left")
        ttk.Button(manage_actions, text="Edit", command=self.edit_schedule_item).pack(side="right")
        ttk.Button(manage_actions, text="↑", width=3, command=lambda: self.move_schedule_item(-1)).pack(side="right", padx=(0, 5))
        ttk.Button(manage_actions, text="↓", width=3, command=lambda: self.move_schedule_item(1)).pack(side="right", padx=(0, 5))
        preview_heading = ttk.Frame(right, style="Card.TFrame")
        preview_heading.pack(fill="x")
        ttk.Label(preview_heading, text="Preview & Live", font=("Segoe UI", 15, "bold"), background="white").pack(side="left")
        ttk.Button(preview_heading, text="▶ Tampilkan Preview ke Live",
                   style="Accent.TButton", command=self.go_live_preview).pack(side="right")
        self.projector_stop_button = ttk.Button(preview_heading, text="■ Hentikan Proyeksi",
                                                command=lambda: self.projector_window and self.projector_window.close_projector())
        self.projector_stop_button.pack(side="right", padx=(0, 6))
        self.projector_stop_button.state(["disabled"])
        toolbar = ttk.Frame(right, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(7, 0))
        self.presentation_menu_button = ttk.Menubutton(toolbar, text="Tampilan ▾")
        style_menu = tk.Menu(self.presentation_menu_button, tearoff=0)
        style_menu.add_command(label="Latar Gambar / Video…", command=self.choose_presentation_background)
        style_menu.add_command(label="Warna Latar…", command=self.choose_presentation_background_color)
        style_menu.add_command(label="Warna Teks…", command=self.choose_presentation_text_color)
        style_menu.add_command(label="Garis Tepi Teks…", command=self.open_presentation_stroke_dialog)
        style_menu.add_separator()
        self.transition_var = tk.StringVar(value=self.transition_style)
        transition_menu = tk.Menu(style_menu, tearoff=0)
        for option in ("Potong", "Fade"):
            transition_menu.add_radiobutton(label=option, value=option, variable=self.transition_var,
                                            command=self.set_transition_style)
        style_menu.add_cascade(label="Transisi", menu=transition_menu)
        self.background_loop_var = tk.BooleanVar(value=self.presentation_background_loop)
        style_menu.add_checkbutton(label="Ulangi video latar", variable=self.background_loop_var,
                                   command=self.set_background_loop)
        style_menu.add_command(label="Jeda / Putar video latar", command=self.toggle_background_video)
        style_menu.add_separator()
        self.youtube_fullscreen_var = tk.BooleanVar(value=self.youtube_fullscreen)
        style_menu.add_checkbutton(label="YouTube otomatis layar penuh",
                                   variable=self.youtube_fullscreen_var,
                                   command=self.set_youtube_fullscreen)
        style_menu.add_command(label="Reset Tampilan", command=self.reset_presentation_style)
        self.presentation_menu_button["menu"] = style_menu
        self.presentation_menu_button.pack(side="left")
        ttk.Button(toolbar, text="Clear", command=self.clear_projector_output).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Blackout", style="Danger.TButton",
                   command=self.blackout_projector).pack(side="left", padx=(5, 0))
        self.presentation_style_status = ttk.Label(right, text="Latar standar · teks putih", style="Sub.TLabel")
        self.presentation_style_status.pack(anchor="w", pady=(4, 0))
        output_pane = ttk.Panedwindow(right, orient="horizontal")
        output_pane.pack(fill="both", expand=True, pady=14)
        parts_card = ttk.Frame(output_pane, style="Card.TFrame", padding=(0, 0, 8, 0))
        preview_card = ttk.Frame(output_pane, style="Card.TFrame", padding=(0, 0, 6, 0))
        live_card = ttk.Frame(output_pane, style="Card.TFrame", padding=(6, 0, 0, 0))
        output_pane.add(parts_card, weight=1)
        output_pane.add(preview_card, weight=2)
        output_pane.add(live_card, weight=2)
        ttk.Label(parts_card, text="BAGIAN ITEM", background="white", foreground="#176b87",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.parts_tree = ttk.Treeview(parts_card, columns=("number", "text"),
                                       show="headings", selectmode="browse", height=6)
        self.parts_tree.heading("number", text="#")
        self.parts_tree.heading("text", text="Teks")
        self.parts_tree.column("number", width=34, stretch=False, anchor="center")
        self.parts_tree.column("text", width=160, stretch=True)
        self.parts_tree.pack(fill="both", expand=True)
        self.parts_tree.bind("<<TreeviewSelect>>", self.select_preview_part)
        self.parts_tree.bind("<Double-1>", self.send_preview_part_live)
        parts_actions = ttk.Frame(parts_card, style="Card.TFrame")
        parts_actions.pack(fill="x", pady=(6, 0))
        ttk.Button(parts_actions, text="Edit Slide…", command=self.edit_preview_slide).pack(fill="x")
        ttk.Label(parts_card, text="Klik untuk pratinjau, klik dua kali untuk tampil ke Live.",
                  background="white", foreground="#65758b", wraplength=190,
                  justify="left").pack(anchor="w", pady=(4, 0))
        ttk.Label(preview_card, text="PREVIEW · PILIHAN OPERATOR", background="white",
                  foreground="#176b87", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.preview_position = ttk.Label(preview_card, text="", style="Sub.TLabel")
        self.preview_position.pack(fill="x", anchor="w", pady=(0, 5))
        self.preview = SlideCanvas(preview_card, bg="#112233")
        self.preview.pack(fill="both", expand=True)
        self.preview.show_text("Pilih item susunan untuk melihat preview", color="white",
                               background="#112233", font=("Segoe UI", 16, "bold"))
        ttk.Label(live_card, text="LIVE · TAMPILAN PROYEKTOR", background="white",
                  foreground="#b44b4b", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        self.live_position = ttk.Label(live_card, text="Belum ada proyeksi aktif", style="Sub.TLabel")
        self.live_position.pack(anchor="w", pady=(0, 4))
        self.live_preview = SlideCanvas(live_card, bg="black")
        self.live_preview.pack(fill="both", expand=True)
        self.live_preview.show_text("Tampilan Live akan terlihat di sini saat proyeksi dimulai",
                                    color="white", background="black", font=("Segoe UI", 16, "bold"))
        self.schedule_tree.bind("<<TreeviewSelect>>", self.preview_schedule_item)
        self.schedule_tree.bind("<Delete>", self.remove_schedule_item)
        self.schedule_tree.bind("<ButtonPress-1>", self._schedule_drag_start, add="+")
        self.schedule_tree.bind("<ButtonRelease-1>", self._schedule_drag_end, add="+")

    def _build_resources(self):
        card = ttk.Frame(self.resources_tab, style="Card.TFrame", padding=20)
        card.pack(fill="both", expand=True)
        ttk.Label(card, text="Resources", font=("Segoe UI", 17, "bold"), background="white").pack(anchor="w")
        ttk.Label(card, text="Cari lagu, ayat AYT, gambar, video, dan PowerPoint dari satu tempat. Hasil bisa dimasukkan ke susunan ibadah.",
                  background="white", foreground="#65758b").pack(anchor="w", pady=(5, 12))
        controls = ttk.Frame(card, style="Card.TFrame")
        controls.pack(fill="x", pady=(0, 8))
        self.resource_query = ttk.Entry(controls, font=("Segoe UI", 12))
        self.resource_query.pack(side="left", fill="x", expand=True)
        self.resource_query.bind("<KeyRelease>", self.search_resources)
        self.resource_kind = ttk.Combobox(controls, state="readonly", width=14,
                                          values=("Semua", "Lagu", "Ayat", "Gambar", "Video", "PowerPoint", "Media"))
        self.resource_kind.set("Semua")
        self.resource_kind.pack(side="left", padx=(8, 0))
        self.resource_kind.bind("<<ComboboxSelected>>", self.search_resources)
        self.resource_tree = ttk.Treeview(card, columns=("kind", "detail"), show="headings", selectmode="extended")
        self.resource_tree.heading("kind", text="Jenis")
        self.resource_tree.heading("detail", text="Judul / Referensi dan isi")
        self.resource_tree.column("kind", width=90, stretch=False)
        self.resource_tree.column("detail", width=750, stretch=True)
        self.resource_tree.pack(fill="both", expand=True)
        self.resource_tree.bind("<Double-1>", self.add_single_resource_to_schedule)
        footer = ttk.Frame(card, style="Card.TFrame")
        footer.pack(fill="x", pady=(10, 0))
        self.resource_status = ttk.Label(footer, text="Ketik judul lagu, potongan lirik, referensi, atau kata dalam ayat.",
                                         style="Sub.TLabel")
        self.resource_status.pack(side="left", fill="x", expand=True)
        ttk.Button(footer, text="＋ Media dari File…", command=self.add_media_to_schedule).pack(side="right", padx=(6, 0))
        ttk.Button(footer, text="Tambahkan ke Susunan", style="Accent.TButton",
                   command=self.add_resource_to_schedule).pack(side="right")

    def search_resources(self, _event=None):
        if not hasattr(self, "resource_tree"):
            return
        term = self.resource_query.get().strip()
        for item in self.resource_tree.get_children():
            self.resource_tree.delete(item)
        if not term:
            self.resource_status.config(text="Ketik judul lagu, potongan lirik, referensi, atau kata dalam ayat.")
            return
        songs, verses, media = self.store.search_resources(term, self.resource_kind.get())
        for song in songs:
            excerpt = re.sub(r"\s+", " ", song["lyrics"]).strip()
            self.resource_tree.insert("", "end", iid=f"song:{song['id']}",
                                      values=("Lagu", f"{song['title']} — {song['author']}  ·  {excerpt[:180]}"))
        for verse in verses:
            reference = f"{verse['book']} {verse['chapter']}:{verse['verse']}"
            self.resource_tree.insert("", "end", iid=f"verse:{verse['id']}",
                                      values=(f"Ayat {verse['version']}", f"{reference}  ·  {verse['text']}"))
        for resource in media:
            self.resource_tree.insert("", "end", iid=f"media:{resource['id']}",
                                      values=(resource["category"],
                                              f"{resource['title']}  ·  {resource['media_path']}"))
        total = len(songs) + len(verses) + len(media)
        self.resource_status.config(text=f"{total} hasil ditemukan · pilih beberapa ayat dengan Ctrl/Shift. Klik dua kali untuk satu item.")

    def add_single_resource_to_schedule(self, event):
        item = self.resource_tree.identify_row(event.y)
        if item:
            self.resource_tree.selection_set(item)
            self.add_resource_to_schedule([item])
        return "break"

    def add_resource_to_schedule(self, item_ids=None):
        selected = list(item_ids if item_ids is not None else self.resource_tree.selection())
        if not selected:
            self.resource_status.config(text="Pilih satu atau beberapa hasil terlebih dahulu.")
            return
        selected_set = set(selected)
        selected = [iid for iid in self.resource_tree.get_children() if iid in selected_set]
        verse_rows = []
        added_item_ids = []
        for iid in selected:
            resource_type, raw_id = iid.split(":", 1)
            if resource_type == "song":
                row = self.store.db.execute("SELECT * FROM songs WHERE id=?", (int(raw_id),)).fetchone()
                if row:
                    added_item_ids.append(self.store.add_item(
                        self.active_schedule["id"], "Lagu", row["title"], row["lyrics"]))
            elif resource_type == "verse":
                row = self.store.db.execute("SELECT * FROM bible_verses WHERE id=?", (int(raw_id),)).fetchone()
                if row:
                    verse_rows.append(row)
            elif resource_type == "media":
                row = self.store.db.execute("SELECT media_path FROM schedule_items WHERE id=? AND media_path<>''",
                                            (int(raw_id),)).fetchone()
                if row and Path(row["media_path"]).is_file():
                    added_item_ids.append(self.store.add_media(self.active_schedule["id"], row["media_path"]))
                elif row:
                    self.resource_status.config(text=f"File media tidak ditemukan: {row['media_path']}")
        if verse_rows:
            references = [f"{row['book']} {row['chapter']}:{row['verse']}" for row in verse_rows]
            consecutive = (len(verse_rows) > 1
                           and all(row["book"] == verse_rows[0]["book"]
                                   and row["chapter"] == verse_rows[0]["chapter"]
                                   for row in verse_rows)
                           and all(verse_rows[i]["verse"] + 1 == verse_rows[i + 1]["verse"]
                                   for i in range(len(verse_rows) - 1)))
            if len(verse_rows) == 1:
                title = references[0]
            elif consecutive:
                title = f"{verse_rows[0]['book']} {verse_rows[0]['chapter']}:{verse_rows[0]['verse']}–{verse_rows[-1]['verse']}"
            else:
                title = f"{len(verse_rows)} ayat terpilih"
            content = "\n\n".join(f"{ref}\n{row['text']}" for ref, row in zip(references, verse_rows))
            added_item_ids.append(self.store.add_item(self.active_schedule["id"], "Ayat", title, content))
        if added_item_ids:
            self.refresh_schedule()
            self.tabs.select(self.schedule_tab)
            last_item = str(added_item_ids[-1])
            self.schedule_tree.selection_set(last_item)
            self.schedule_tree.focus(last_item)

    def _schedule_drag_start(self, event):
        self._schedule_drag_iid = self.schedule_tree.identify_row(event.y)

    def _schedule_drag_end(self, event):
        source = getattr(self, "_schedule_drag_iid", "")
        target = self.schedule_tree.identify_row(event.y)
        self._schedule_drag_iid = ""
        if not source or not target or source == target:
            return
        rows = self.store.items(self.active_schedule["id"])
        order = [row["id"] for row in rows]
        try:
            source_index, target_index = order.index(int(source)), order.index(int(target))
        except ValueError:
            return
        item_id = order.pop(source_index)
        order.insert(target_index, item_id)
        with self.store.db:
            for position, current_id in enumerate(order, start=1):
                self.store.db.execute("UPDATE schedule_items SET position=? WHERE id=? AND schedule_id=?",
                                      (position, current_id, self.active_schedule["id"]))
        self.refresh_schedule()
        self.schedule_tree.selection_set(str(item_id))
        self.schedule_tree.focus(str(item_id))

    def _build_songs(self):
        pane = ttk.Panedwindow(self.songs_tab, orient="horizontal")
        pane.pack(fill="both", expand=True)
        left = ttk.Frame(pane, style="Card.TFrame", padding=12)
        right = ttk.Frame(pane, style="Card.TFrame", padding=18)
        pane.add(left, weight=2); pane.add(right, weight=3)
        ttk.Label(left, text="Lagu", font=("Segoe UI", 15, "bold"), background="white").pack(anchor="w", pady=(0, 10))
        self.song_tree = ttk.Treeview(left, columns=("author",), show="tree headings", selectmode="browse")
        self.song_tree.heading("#0", text="Judul"); self.song_tree.heading("author", text="Penulis")
        self.song_tree.column("#0", width=220); self.song_tree.column("author", width=110)
        self.song_tree.pack(fill="both", expand=True)
        self.song_tree.bind("<<TreeviewSelect>>", self.load_selected_song)
        ttk.Button(left, text="＋ Lagu Baru", style="Accent.TButton", command=self.new_song).pack(anchor="w", pady=(12, 0))
        song_actions = ttk.Frame(left, style="Card.TFrame")
        song_actions.pack(fill="x", pady=(7, 0))
        ttk.Button(song_actions, text="Impor CSV", command=self.import_songs).pack(side="left")
        ttk.Button(song_actions, text="Ekspor CSV", command=self.export_songs).pack(side="left", padx=6)
        ttk.Button(song_actions, text="Hapus Lagu", command=self.delete_song).pack(side="right")
        ttk.Label(right, text="Detail Lagu", font=("Segoe UI", 15, "bold"), background="white").pack(anchor="w", pady=(0, 12))
        ttk.Label(right, text="Judul", background="white").pack(anchor="w")
        self.song_title = ttk.Entry(right); self.song_title.pack(fill="x", pady=(3, 10))
        ttk.Label(right, text="Penulis / Sumber", background="white").pack(anchor="w")
        self.song_author = ttk.Entry(right); self.song_author.pack(fill="x", pady=(3, 10))
        ttk.Label(right, text="Lirik (pisahkan bagian dengan satu baris kosong; nomor bagian bisa dipilih saat proyeksi)",
                  background="white", wraplength=460, justify="left").pack(anchor="w")
        self.song_lyrics = tk.Text(right, height=15, wrap="word", font=("Segoe UI", 11), relief="solid", bd=1)
        self.song_lyrics.pack(fill="both", expand=True, pady=(4, 12))
        ttk.Button(right, text="Simpan Lagu", style="Accent.TButton", command=self.save_song).pack(anchor="e")

    def _build_verses(self):
        host = ttk.Frame(self.verses_tab, style="Card.TFrame")
        host.pack(fill="both", expand=True)
        self.verses_canvas = tk.Canvas(host, bg="#f2f5f8", highlightthickness=0, bd=0)
        verses_scrollbar = ttk.Scrollbar(host, orient="vertical",
                                         command=self.verses_canvas.yview)
        self.verses_canvas.configure(yscrollcommand=verses_scrollbar.set)
        self.verses_canvas.pack(side="left", fill="both", expand=True)
        verses_scrollbar.pack(side="right", fill="y")
        card = ttk.Frame(self.verses_canvas, style="Card.TFrame", padding=22)
        self.verses_inner = card
        card_window = self.verses_canvas.create_window((0, 0), window=card, anchor="nw")
        card.bind("<Configure>", lambda _event: self.verses_canvas.configure(
            scrollregion=self.verses_canvas.bbox("all")))
        self.verses_canvas.bind("<Configure>", lambda event: self.verses_canvas.itemconfigure(
            card_window, width=event.width))
        ttk.Label(card, text="Tampilkan Ayat Alkitab", font=("Segoe UI", 17, "bold"), background="white").pack(anchor="w")
        ttk.Label(card, text="Impor beberapa versi Alkitab CSV, pilih versi, cari ayat, lalu tambahkan ke susunan.",
                  background="white", foreground="#65758b").pack(anchor="w", pady=(5, 12))
        ttk.Label(card, text="Tampilkan Alkitab", font=("Segoe UI", 12, "bold"),
                  background="white").pack(anchor="w")
        picker = ttk.Frame(card, style="Card.TFrame")
        picker.pack(fill="x", pady=(4, 0))
        ttk.Label(picker, text="Versi:").pack(side="left")
        self.bible_version_choice = ttk.Combobox(picker, state="readonly", width=8)
        self.bible_version_choice.pack(side="left", padx=(5, 12))
        self.bible_version_choice.bind("<<ComboboxSelected>>", self.on_bible_version_changed)
        ttk.Label(picker, text="Kitab:").pack(side="left")
        self.bible_book_choice = ttk.Combobox(picker, width=22)
        self.bible_book_choice.pack(side="left", padx=(5, 12))
        self.bible_book_choice.bind("<<ComboboxSelected>>", self._bible_book_chosen)
        self.bible_book_choice.bind("<Return>", self._bible_book_chosen)
        self.bible_book_choice.bind("<FocusOut>", self._bible_book_chosen)
        self.bible_book_choice.bind("<KeyRelease>", self._filter_bible_books)
        ttk.Label(picker, text="Pasal:").pack(side="left")
        self.bible_chapter_choice = ttk.Combobox(picker, state="readonly", width=6)
        self.bible_chapter_choice.pack(side="left", padx=(5, 12))
        self.bible_chapter_choice.bind("<<ComboboxSelected>>", self._bible_chapter_chosen)
        ttk.Label(picker, text="Ayat:").pack(side="left")
        self.bible_verse_choice = ttk.Combobox(picker, state="readonly", width=9)
        self.bible_verse_choice.pack(side="left", padx=(5, 12))
        ttk.Button(picker, text="Tampilkan", command=self.show_bible_passage).pack(side="left")
        ttk.Label(card, text="Ketik nama kitab (contoh: Mat) lalu Enter, pilih pasal dan ayat, "
                             "atau SEMUA untuk seluruh pasal.",
                  background="white", foreground="#65758b").pack(anchor="w", pady=(3, 10))
        library_actions = ttk.Frame(card, style="Card.TFrame")
        library_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(library_actions, text="Impor Dataset Alkitab CSV…", command=self.import_bible_dataset).pack(side="left")
        self.bible_count = ttk.Label(library_actions, text="", style="Sub.TLabel")
        self.bible_count.pack(side="left", padx=10)
        ttk.Button(library_actions, text="Tambah Ayat Terpilih", command=self.add_selected_bible_verse).pack(side="right")
        ttk.Label(card, text="Cari satu ayat atau rentang, contoh: Kejadian 1:4-9. Pilih beberapa baris dengan Ctrl/Shift; setiap ayat menjadi slide bernomor.",
                  background="white", foreground="#65758b").pack(anchor="w", pady=(0, 5))
        self.bible_search = ttk.Entry(card, font=("Segoe UI", 11))
        self.bible_search.pack(fill="x", pady=(0, 6))
        self.bible_search.insert(0, "Cari kitab, referensi, atau kata dalam ayat…")
        self.bible_search.bind("<FocusIn>", self._clear_bible_search_hint)
        self.bible_search.bind("<KeyRelease>", self.search_bible_library)
        self.bible_tree = ttk.Treeview(card, columns=("reference", "text"), show="headings", height=6,
                                       selectmode="extended")
        self.bible_tree.heading("reference", text="Referensi")
        self.bible_tree.heading("text", text="Teks Ayat")
        self.bible_tree.column("reference", width=190, stretch=False)
        self.bible_tree.column("text", width=650, stretch=True)
        self.bible_tree.pack(fill="x", pady=(0, 8))
        self.bible_tree.bind("<<TreeviewSelect>>", self.load_bible_selection)
        self.bible_tree.bind("<Double-1>", lambda _event: self.add_verse_to_schedule())
        ttk.Label(card, text="Referensi (contoh: Mazmur 23:1)", background="white").pack(anchor="w")
        self.verse_ref = ttk.Entry(card, font=("Segoe UI", 12)); self.verse_ref.pack(fill="x", pady=(4, 12))
        ttk.Label(card, text="Teks ayat", background="white").pack(anchor="w")
        self.verse_text = tk.Text(card, height=6, wrap="word", font=("Segoe UI", 12), relief="solid", bd=1)
        self.verse_text.pack(fill="both", expand=True, pady=(4, 12))
        verse_actions = ttk.Frame(card, style="Card.TFrame")
        verse_actions.pack(fill="x")
        ttk.Button(verse_actions, text="Tambah Rentang ke Susunan", command=self.add_verse_range_to_schedule).pack(side="left")
        ttk.Button(verse_actions, text="Tambahkan Ayat Manual ke Susunan", style="Accent.TButton",
                   command=self.add_verse_to_schedule).pack(side="right")
        self._bible_books = []
        self._bible_abbrs = {}
        self._bible_current_book = None
        self.refresh_bible_picker()

    def bible_count_text(self):
        try:
            total = self.store.db.execute("SELECT COUNT(*) FROM bible_verses").fetchone()[0]
            version = self.bible_version_choice.get() if hasattr(self, "bible_version_choice") else "AYT"
            count = self.store.bible_version_count(version) if version else 0
        except sqlite3.OperationalError:
            total = count = 0
        return f"{version or 'AYT'}: {count:,} ayat · total {total:,}"

    def _clear_bible_search_hint(self, _event=None):
        if self.bible_search.get() == "Cari kitab, referensi, atau kata dalam ayat…":
            self.bible_search.delete(0, "end")

    def refresh_bible_picker(self):
        if not hasattr(self, "bible_book_choice"):
            return
        versions = self.store.bible_versions()
        self.bible_version_choice["values"] = versions
        current_version = self.bible_version_choice.get()
        if current_version not in versions:
            current_version = "AYT" if "AYT" in versions else versions[0]
        self.bible_version_choice.set(current_version)
        self.bible_count.config(text=self.bible_count_text())
        self._bible_books = self.store.bible_books(current_version)
        self._bible_abbrs = self.store.bible_abbreviations(current_version)
        self.bible_book_choice["values"] = self._bible_books
        current = self._bible_current_book
        if current not in self._bible_books:
            current = self._bible_books[0] if self._bible_books else None
        self._bible_current_book = current
        if current:
            self.bible_book_choice.set(current)
            self._refresh_bible_chapters(current)
        else:
            self.bible_book_choice.set("")
            self.bible_chapter_choice["values"] = []
            self.bible_chapter_choice.set("")
            self.bible_verse_choice["values"] = []
            self.bible_verse_choice.set("")

    def on_bible_version_changed(self, _event=None):
        self._bible_current_book = None
        self.refresh_bible_picker()
        self.search_bible_library()

    def _resolve_bible_book(self, text):
        text = (text or "").strip()
        if not text:
            return None
        lowered = text.casefold().rstrip(".")
        for book in getattr(self, "_bible_books", []):
            if book.casefold() == lowered:
                return book
        version = self.bible_version_choice.get() if hasattr(self, "bible_version_choice") else "AYT"
        normalized = normalize_bible_book(text, version)
        for book in getattr(self, "_bible_books", []):
            if book.casefold() == normalized.casefold():
                return book
        abbreviation = getattr(self, "_bible_abbrs", {}).get(lowered)
        if abbreviation:
            return abbreviation
        books = getattr(self, "_bible_books", [])
        prefix = [book for book in books if book.casefold().startswith(lowered)]
        if len(prefix) == 1:
            return prefix[0]
        contains = [book for book in books if lowered in book.casefold()]
        if len(contains) == 1:
            return contains[0]
        return None

    def _filter_bible_books(self, _event=None):
        if not hasattr(self, "bible_book_choice"):
            return
        text = self.bible_book_choice.get().strip().casefold()
        books = getattr(self, "_bible_books", [])
        if not text:
            self.bible_book_choice["values"] = books
            return
        matches = [book for book in books if text in book.casefold()]
        self.bible_book_choice["values"] = matches or books

    def _bible_book_chosen(self, _event=None):
        if not hasattr(self, "bible_book_choice"):
            return
        book = self._resolve_bible_book(self.bible_book_choice.get())
        if book is None:
            if self._bible_current_book:
                self.bible_book_choice.set(self._bible_current_book)
            return
        self._bible_current_book = book
        self.bible_book_choice.set(book)
        self._refresh_bible_chapters(book)

    def _refresh_bible_chapters(self, book):
        chapters = [str(chapter) for chapter in self.store.bible_chapters(book, self.bible_version_choice.get())]
        self.bible_chapter_choice["values"] = chapters
        current = self.bible_chapter_choice.get()
        if current not in chapters:
            current = chapters[0] if chapters else ""
        self.bible_chapter_choice.set(current)
        if current.isdigit():
            self._refresh_bible_verses(book, int(current))
        else:
            self.bible_verse_choice["values"] = []
            self.bible_verse_choice.set("")

    def _bible_chapter_chosen(self, _event=None):
        book = self._bible_current_book
        chapter = self.bible_chapter_choice.get().strip()
        if book and chapter.isdigit():
            self._refresh_bible_verses(book, int(chapter))

    def _refresh_bible_verses(self, book, chapter):
        verses = [str(verse) for verse in self.store.bible_verse_numbers(
            book, chapter, self.bible_version_choice.get())]
        values = ["SEMUA"] + verses
        self.bible_verse_choice["values"] = values
        if self.bible_verse_choice.get() not in values:
            self.bible_verse_choice.set("SEMUA")

    def show_bible_passage(self):
        book = self._resolve_bible_book(self.bible_book_choice.get())
        if book is None:
            messagebox.showinfo("Pilih kitab", "Ketik atau pilih nama kitab terlebih dahulu.", parent=self)
            return
        self._bible_current_book = book
        self.bible_book_choice.set(book)
        chapter_text = self.bible_chapter_choice.get().strip()
        if not chapter_text.isdigit():
            messagebox.showinfo("Pilih pasal", "Pilih pasal terlebih dahulu.", parent=self)
            return
        chapter = int(chapter_text)
        verse_text = self.bible_verse_choice.get().strip().casefold()
        verse = int(verse_text) if verse_text.isdigit() else None
        rows = self.store.bible_chapter_verses(book, chapter, verse,
                                               version=self.bible_version_choice.get())
        if not rows:
            messagebox.showinfo("Ayat tidak ditemukan",
                                f"Tidak ada ayat untuk {book} {chapter}.", parent=self)
            return
        for item in self.bible_tree.get_children():
            self.bible_tree.delete(item)
        for row in rows:
            reference = f"{row['book']} {row['chapter']}:{row['verse']}"
            self.bible_tree.insert("", "end", iid=str(row["id"]), values=(reference, row["text"]))
        items = self.bible_tree.get_children()
        if items:
            self.bible_tree.selection_set(items if verse is None else items[:1])
            self.bible_tree.focus(items[0])
            self.bible_tree.see(items[0])
        first = rows[0]
        self.verse_ref.delete(0, "end")
        self.verse_ref.insert(0, f"{first['book']} {first['chapter']}:{first['verse']}")
        self.verse_text.delete("1.0", "end")
        self.verse_text.insert("1.0", first["text"])
        self.bible_search.delete(0, "end")

    def import_bible_dataset(self):
        path = filedialog.askopenfilename(
            parent=self, title="Pilih dataset Alkitab CSV",
            filetypes=[("Dataset Alkitab", "*.csv"), ("Semua file", "*.*")])
        if not path:
            return
        filename = Path(path).name.casefold()
        if "kjv" in filename:
            version = "KJV"
        elif "amp" in filename:
            version = "AMP"
        elif "niv" in filename:
            version = "NIV"
        elif "ayt" in filename:
            version = "AYT"
        else:
            version = simpledialog.askstring(
                "Nama Versi Alkitab", "Masukkan kode versi untuk dataset ini (contoh: TB, BIS, KJV):",
                initialvalue=self.bible_version_choice.get() or "TB", parent=self)
            if not version or not version.strip():
                return
            version = version.strip().upper()
        try:
            count = self.store.import_bible_csv(path, version)
        except (OSError, UnicodeError, csv.Error, ValueError, sqlite3.Error) as error:
            messagebox.showerror("Impor dataset gagal", str(error), parent=self)
            return
        self.bible_version_choice.set(version)
        self._bible_current_book = None
        self.refresh_bible_picker()
        self.search_bible_library()
        messagebox.showinfo("Impor selesai", f"{count:,} ayat versi {version} berhasil dimasukkan atau diperbarui.", parent=self)

    def import_ayt_dataset(self):
        # Keep the old handler name for callers from earlier versions.
        return self.import_bible_dataset()

    def search_bible_library(self, _event=None):
        if not hasattr(self, "bible_tree"):
            return
        term = self.bible_search.get().strip()
        if term == "Cari kitab, referensi, atau kata dalam ayat…":
            term = ""
        for item in self.bible_tree.get_children():
            self.bible_tree.delete(item)
        rows = self.store.search_bible_verses(term, version=self.bible_version_choice.get())
        for row in rows:
            reference = f"{row['book']} {row['chapter']}:{row['verse']}"
            self.bible_tree.insert("", "end", iid=str(row["id"]), values=(reference, row["text"]))

    def load_bible_selection(self, _event=None):
        selected = self.bible_tree.selection()
        if not selected:
            return
        row = self.store.db.execute("SELECT * FROM bible_verses WHERE id=?", (int(selected[0]),)).fetchone()
        if row:
            self.verse_ref.delete(0, "end")
            self.verse_ref.insert(0, f"{row['book']} {row['chapter']}:{row['verse']} [{row['version']}]")
            self.verse_text.delete("1.0", "end")
            self.verse_text.insert("1.0", row["text"])

    def add_selected_bible_verse(self):
        selected = set(self.bible_tree.selection())
        if not selected:
            messagebox.showinfo("Pilih ayat", "Pilih satu ayat dari daftar terlebih dahulu.", parent=self)
            return
        rows = []
        for item_id in self.bible_tree.get_children():
            if item_id in selected:
                row = self.store.db.execute("SELECT * FROM bible_verses WHERE id=?", (int(item_id),)).fetchone()
                if row:
                    rows.append(row)
        if not rows:
            return
        version = self.bible_version_choice.get() or rows[0]["version"]
        refs = [f"{row['book']} {row['chapter']}:{row['verse']}" for row in rows]
        if len(rows) == 1:
            title = refs[0]
        elif all(row["book"] == rows[0]["book"] and row["chapter"] == rows[0]["chapter"]
                 for row in rows) and all(rows[i]["verse"] + 1 == rows[i + 1]["verse"]
                                          for i in range(len(rows) - 1)):
            title = f"{rows[0]['book']} {rows[0]['chapter']}:{rows[0]['verse']}–{rows[-1]['verse']}"
        else:
            title = f"{len(rows)} ayat terpilih"
        self.store.add_item(self.active_schedule["id"], "Ayat", f"{title} [{version}]",
                            "\n\n".join(f"{ref} [{version}]\n{row['text']}"
                                         for ref, row in zip(refs, rows)))
        self.refresh_schedule()
        self.tabs.select(self.schedule_tab)

    def add_verse_range_to_schedule(self):
        query = self.bible_search.get().strip()
        version = self.bible_version_choice.get() or "AYT"
        rows = self.store.search_bible_verses(query, version=version)
        if not re.fullmatch(r".+?\s+\d+:\d+\s*-\s*(?:\d+:)?\d+", query) or not rows:
            messagebox.showinfo("Rentang ayat", "Cari rentang dengan format seperti Kejadian 1:4-9, lalu coba lagi.", parent=self)
            return
        parts = [f"{row['book']} {row['chapter']}:{row['verse']} [{version}]\n{row['text']}" for row in rows]
        start_ref = f"{rows[0]['book']} {rows[0]['chapter']}:{rows[0]['verse']}"
        end_ref = f"{rows[-1]['chapter']}:{rows[-1]['verse']}"
        title = f"{start_ref}–{end_ref} [{version}]"
        self.store.add_item(self.active_schedule["id"], "Ayat", title, "\n\n".join(parts))
        self.refresh_schedule()
        self.tabs.select(self.schedule_tab)

    def refresh_songs(self):
        for item in self.song_tree.get_children(): self.song_tree.delete(item)
        for song in self.store.songs(): self.song_tree.insert("", "end", iid=str(song["id"]), text=song["title"], values=(song["author"],))

    def load_selected_song(self, _event=None):
        selected = self.song_tree.selection()
        if not selected: return
        song = self.store.db.execute("SELECT * FROM songs WHERE id=?", (int(selected[0]),)).fetchone()
        self.current_song_id = song["id"]
        self.song_title.delete(0, "end"); self.song_title.insert(0, song["title"])
        self.song_author.delete(0, "end"); self.song_author.insert(0, song["author"])
        self.song_lyrics.delete("1.0", "end"); self.song_lyrics.insert("1.0", song["lyrics"])

    def new_song(self):
        self.current_song_id = None
        self.song_tree.selection_remove(self.song_tree.selection())
        self.song_title.delete(0, "end"); self.song_author.delete(0, "end"); self.song_lyrics.delete("1.0", "end")
        self.song_title.focus_set()

    def save_song(self):
        title = self.song_title.get().strip()
        if not title:
            messagebox.showwarning("Judul diperlukan", "Isi judul lagu terlebih dahulu."); return
        self.store.save_song(self.current_song_id, title, self.song_author.get().strip(), self.song_lyrics.get("1.0", "end").strip())
        self.refresh_songs()
        if hasattr(self, "resource_query") and self.resource_query.get().strip():
            self.search_resources()
        messagebox.showinfo("Tersimpan", "Lagu berhasil disimpan di database lokal.")

    def refresh_schedule(self):
        schedules = self.store.schedules()
        if not hasattr(self, "active_schedule_id") or not any(s["id"] == self.active_schedule_id for s in schedules):
            self.active_schedule_id = schedules[0]["id"]
        self.active_schedule = next(s for s in schedules if s["id"] == self.active_schedule_id)
        self.schedule_labels = {f"{s['title']}  (#{index})": s["id"] for index, s in enumerate(schedules, start=1)}
        self.schedule_choice["values"] = list(self.schedule_labels)
        label = next(k for k, v in self.schedule_labels.items() if v == self.active_schedule_id)
        self.schedule_choice.set(label)
        for item in self.schedule_tree.get_children(): self.schedule_tree.delete(item)
        for number, row in enumerate(self.store.items(self.active_schedule["id"]), start=1):
            note = re.sub(r"\s+", " ", row["operator_note"] or "").strip()
            self.schedule_tree.insert("", "end", iid=str(row["id"]),
                                      values=(f"{number}.  {row['title']}", row["kind"], note))
        if hasattr(self, "broadcast_item_choice"):
            self._broadcast_refresh_item_choices()

    def select_schedule(self, _event=None):
        schedule_id = self.schedule_labels.get(self.schedule_choice.get())
        if schedule_id:
            self.active_schedule_id = schedule_id
            self.refresh_schedule()

    def new_schedule(self):
        title = simpledialog.askstring("Jadwal Baru", "Nama jadwal ibadah:", parent=self)
        if not title or not title.strip():
            return
        self.active_schedule_id = self.store.create_schedule(title.strip())
        self.refresh_schedule()

    def save_schedule_file(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Simpan susunan ibadah", defaultextension=".sahabat.json",
            filetypes=[("Susunan Sahabat Ibadah", "*.sahabat.json"), ("JSON", "*.json")],
            initialfile=f"{self.active_schedule['title']}.sahabat.json")
        if not path:
            return
        rows = self.store.items(self.active_schedule["id"])
        document = {
            "format": "sahabat-ibadah-schedule", "version": 1,
            "title": self.active_schedule["title"],
            "items": [{key: (row[key] or "") for key in
                       ("kind", "title", "content", "media_path", "operator_note")} for row in rows],
        }
        try:
            with open(path, "w", encoding="utf-8") as output:
                json.dump(document, output, ensure_ascii=False, indent=2)
            messagebox.showinfo("Susunan tersimpan", f"Susunan berhasil disimpan ke:\n{path}", parent=self)
        except OSError as error:
            messagebox.showerror("Gagal menyimpan susunan", str(error), parent=self)

    def open_schedule_file(self):
        path = filedialog.askopenfilename(
            parent=self, title="Buka susunan ibadah", filetypes=[("Susunan Sahabat Ibadah", "*.json"),
                                                                    ("Semua file", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as source:
                document = json.load(source)
            items = document.get("items")
            if (document.get("format") != "sahabat-ibadah-schedule" or document.get("version") != 1
                    or not isinstance(items, list) or len(items) > 2000):
                raise ValueError("File bukan susunan Sahabat Ibadah yang didukung.")
            normalized = []
            for row in items:
                if not isinstance(row, dict):
                    raise ValueError("Ada item susunan dengan format yang tidak valid.")
                kind = row.get("kind", "")
                if kind not in {"Lagu", "Ayat", "Media", "PowerPoint", "URL"}:
                    raise ValueError(f"Jenis item tidak didukung: {kind}")
                values = [row.get(key, "") for key in ("title", "content", "media_path", "operator_note")]
                if not all(isinstance(value, str) for value in values) or not values[0].strip():
                    raise ValueError("Judul atau isi item susunan tidak valid.")
                normalized.append((kind, *values))
            title = document.get("title", "Susunan Ibadah")
            if not isinstance(title, str) or not title.strip():
                title = "Susunan Ibadah"
        except (OSError, json.JSONDecodeError, ValueError) as error:
            messagebox.showerror("File susunan tidak dapat dibuka", str(error), parent=self)
            return
        db = self.store.db
        try:
            cursor = db.execute("INSERT INTO schedules(title) VALUES (?)", (title.strip(),))
            schedule_id = cursor.lastrowid
            db.executemany("INSERT INTO schedule_items(schedule_id,kind,title,content,position,media_path,operator_note) "
                           "VALUES (?,?,?,?,?,?,?)",
                           [(schedule_id, kind, item_title, content, index, media_path, note)
                            for index, (kind, item_title, content, media_path, note) in enumerate(normalized, 1)])
            db.commit()
            self.active_schedule_id = schedule_id
            self.refresh_schedule()
            missing = [media_path for _kind, _title, _content, media_path, _note in normalized
                       if media_path and not Path(media_path).exists()]
            message = f"Susunan '{title}' berhasil dibuka ({len(normalized)} item)."
            if missing:
                message += f"\n\n{len(missing)} file media belum ditemukan di lokasi aslinya."
            messagebox.showinfo("Susunan dibuka", message, parent=self)
        except sqlite3.Error as error:
            db.rollback()
            messagebox.showerror("Gagal membuka susunan", str(error), parent=self)

    def open_data_tools(self):
        dialog = tk.Toplevel(self)
        dialog.title("Pencadangan dan Pemulihan Data")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Data Sahabat Ibadah", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(body, text="Cadangkan database lagu, ayat, dan susunan ibadah ke satu file ZIP.\n"
                            "Pemulihan akan mengganti data lokal dengan isi file cadangan.",
                  justify="left", wraplength=430).pack(anchor="w", pady=(7, 16))
        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Buat Cadangan…", style="Accent.TButton",
                   command=self.backup_application_data).pack(side="left")
        ttk.Button(buttons, text="Pulihkan Cadangan…", command=self.restore_application_data).pack(side="left", padx=8)
        ttk.Button(body, text="Tutup", command=dialog.destroy).pack(anchor="e", pady=(16, 0))

    def backup_application_data(self):
        suggested = f"Sahabat-Ibadah-{time.strftime('%Y%m%d-%H%M%S')}.zip"
        path = filedialog.asksaveasfilename(parent=self, title="Buat cadangan data", defaultextension=".zip",
                                            initialfile=suggested, filetypes=[("Cadangan ZIP", "*.zip")])
        if not path:
            return
        try:
            with tempfile.TemporaryDirectory(prefix="sahabat-backup-") as temp_dir:
                db_copy = Path(temp_dir) / "sahabat_ibadah.db"
                with sqlite3.connect(db_copy) as target:
                    self.store.db.backup(target)
                manifest = {"format": "sahabat-ibadah-backup", "version": 1,
                            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "includes": ["sahabat_ibadah.db"]}
                with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
                    archive.write(db_copy, "sahabat_ibadah.db")
                    ayt_path = APP_DIR / "ayt.csv"
                    if ayt_path.exists():
                        archive.write(ayt_path, "ayt.csv")
                        manifest["includes"].append("ayt.csv")
                    archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            messagebox.showinfo("Cadangan selesai", f"Data berhasil dicadangkan ke:\n{path}", parent=self)
        except (OSError, sqlite3.Error, zipfile.BadZipFile) as error:
            messagebox.showerror("Cadangan gagal", str(error), parent=self)

    def restore_application_data(self):
        if self.projector_window and self.projector_window.winfo_exists():
            messagebox.showwarning("Proyeksi sedang aktif", "Tutup proyeksi sebelum memulihkan data.", parent=self)
            return
        path = filedialog.askopenfilename(parent=self, title="Pilih file cadangan",
                                          filetypes=[("Cadangan Sahabat Ibadah", "*.zip"), ("Semua file", "*.*")])
        if not path:
            return
        if not messagebox.askyesno("Ganti data lokal?",
                                   "Pemulihan akan mengganti seluruh data lokal saat ini. Lanjutkan?", parent=self):
            return
        temp_db = None
        try:
            with zipfile.ZipFile(path, "r") as archive:
                names = set(archive.namelist())
                if "sahabat_ibadah.db" not in names or "manifest.json" not in names:
                    raise ValueError("File ZIP ini bukan cadangan Sahabat Ibadah yang valid.")
                manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
                if manifest.get("format") != "sahabat-ibadah-backup" or manifest.get("version") != 1:
                    raise ValueError("Versi atau format file cadangan tidak didukung.")
                info = archive.getinfo("sahabat_ibadah.db")
                if info.file_size > 1_000_000_000:
                    raise ValueError("Ukuran database cadangan melebihi batas 1 GB.")
                fd, temp_name = tempfile.mkstemp(prefix="sahabat-restore-", suffix=".db", dir=APP_DIR)
                os.close(fd)
                temp_db = Path(temp_name)
                temp_db.write_bytes(archive.read("sahabat_ibadah.db"))
            with sqlite3.connect(f"file:{temp_db.as_posix()}?mode=ro", uri=True) as check:
                integrity = check.execute("PRAGMA integrity_check").fetchone()[0]
                tables = {row[0] for row in check.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if integrity != "ok" or not {"songs", "schedules", "schedule_items"}.issubset(tables):
                    raise ValueError("Database dalam cadangan rusak atau tidak lengkap.")
            self.store.db.close()
            self.store = None
            os.replace(temp_db, DB_PATH)
            temp_db = None
            self.store = Store()
            self.live_text_background = self.store.get_setting("live_text_background", "0") == "1"
            if hasattr(self, "live_background_var"):
                self.live_background_var.set(self.live_text_background)
            self._load_live_layout_settings()
            if hasattr(self, "live_overlay_position"):
                self.live_overlay_position.set(self.live_text_position)
                self._update_live_layout_status()
            self.refresh_bible_picker()
            self.refresh_songs()
            self.active_schedule_id = None
            self.refresh_schedule()
            if hasattr(self, "resource_query") and self.resource_query.get().strip():
                self.search_resources()
            messagebox.showinfo("Pemulihan selesai", "Data dari cadangan berhasil dipulihkan.", parent=self)
        except (OSError, sqlite3.Error, zipfile.BadZipFile, json.JSONDecodeError, KeyError, ValueError) as error:
            if self.store is None:
                self.store = Store()
            messagebox.showerror("Pemulihan gagal", str(error), parent=self)
        finally:
            if temp_db and temp_db.exists():
                temp_db.unlink(missing_ok=True)

    def blackout_projector(self):
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.toggle_blackout()
        else:
            messagebox.showinfo("Proyeksi belum aktif", "Mulai proyeksi terlebih dahulu.", parent=self)

    def clear_projector_output(self):
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.clear_output()
        else:
            messagebox.showinfo("Proyeksi belum aktif", "Mulai proyeksi terlebih dahulu.", parent=self)

    def duplicate_schedule(self):
        self.active_schedule_id = self.store.duplicate_schedule(self.active_schedule["id"])
        self.refresh_schedule()

    def delete_schedule(self):
        """Merespons klik menu Hapus Jadwal Ini di kelas UI SahabatIbadahApp."""
        schedules = self.store.schedules()
        if len(schedules) <= 1:
            messagebox.showwarning(
                "Tidak Dapat Dihapus",
                "Aplikasi harus memiliki minimal satu jadwal ibadah.",
                parent=self
            )
            return

        active_title = self.active_schedule["title"]
        if not messagebox.askyesno(
            "Konfirmasi Hapus Jadwal",
            f"Apakah Anda yakin ingin menghapus jadwal '{active_title}'?\n\n"
            "Seluruh item susunan dalam jadwal ini akan dihapus.",
            parent=self
        ):
            return

        # Panggil fungsi hapus dari kelas Store menggunakan ID jadwal yang aktif
        self.store.delete_schedule_by_id(self.active_schedule["id"])
        
        # Reset ID aktif agar aplikasi otomatis memilih jadwal lain yang tersisa
        self.active_schedule_id = None
        self.refresh_schedule()
        messagebox.showinfo("Berhasil", f"Jadwal '{active_title}' berhasil dihapus.", parent=self)

    def edit_schedule_item(self):
        selected = self.schedule_tree.selection()
        if not selected:
            messagebox.showinfo("Pilih item", "Pilih item jadwal yang ingin diedit terlebih dahulu.")
            return
        row = self.store.db.execute("SELECT * FROM schedule_items WHERE id=? AND schedule_id=?",
                                    (int(selected[0]), self.active_schedule["id"])).fetchone()
        if not row:
            return
        dialog = tk.Toplevel(self)
        dialog.title("Edit Item Jadwal")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("600x590")
        dialog.minsize(520, 480)
        dialog.configure(bg="#f2f5f8")
        body = ttk.Frame(dialog, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"{row['kind']} · Judul", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        title_field = ttk.Entry(body)
        title_field.pack(fill="x", pady=(5, 12))
        title_field.insert(0, row["title"])
        ttk.Label(body, text="Teks (untuk media, teks dapat dikosongkan)").pack(anchor="w")
        content_field = tk.Text(body, height=10, wrap="word", font=("Segoe UI", 11))
        content_field.pack(fill="both", expand=True, pady=(5, 10))
        content_field.insert("1.0", row["content"])
        ttk.Label(body, text="Catatan operator (hanya terlihat di aplikasi operator)",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        note_field = tk.Text(body, height=4, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1)
        note_field.pack(fill="x", pady=(5, 12))
        note_field.insert("1.0", row["operator_note"] or "")

        def save_changes():
            title = title_field.get().strip()
            if not title:
                messagebox.showwarning("Judul diperlukan", "Judul item tidak boleh kosong.", parent=dialog)
                return
            self.store.update_item(row["id"], self.active_schedule["id"], title,
                                   content_field.get("1.0", "end").strip(),
                                   note_field.get("1.0", "end").strip())
            self.refresh_schedule()
            self.schedule_tree.selection_set(str(row["id"]))
            dialog.destroy()

        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Batal", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="Simpan", style="Accent.TButton", command=save_changes).pack(side="right", padx=7)
        title_field.focus_set()

    def add_song_to_schedule(self):
        songs = self.store.songs()
        if not songs:
            messagebox.showinfo("Belum ada lagu", "Tambahkan lagu di Perpustakaan Lagu terlebih dahulu."); return
        dialog = tk.Toplevel(self)
        dialog.title("Pilih Lagu")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("620x460")
        dialog.minsize(460, 320)
        body = ttk.Frame(dialog, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Cari berdasarkan sebagian judul atau nama penyanyi", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        query = tk.StringVar()
        search = ttk.Entry(body, textvariable=query, font=("Segoe UI", 12))
        search.pack(fill="x", pady=(7, 10))
        listing = ttk.Treeview(body, columns=("author",), show="tree headings", selectmode="browse")
        listing.heading("#0", text="Judul Lagu")
        listing.heading("author", text="Penyanyi / Sumber")
        listing.column("#0", width=350, stretch=True)
        listing.column("author", width=200, stretch=True)
        listing.pack(fill="both", expand=True)
        status = ttk.Label(body, text="", style="Sub.TLabel")
        status.pack(anchor="w", pady=(6, 0))
        songs_by_id = {str(song["id"]): song for song in songs}

        def refresh_matches(*_args):
            term = query.get().strip().casefold()
            for existing in listing.get_children():
                listing.delete(existing)
            matches = [song for song in songs
                       if not term or term in song["title"].casefold() or term in song["author"].casefold()]
            for song in matches:
                listing.insert("", "end", iid=str(song["id"]), text=song["title"], values=(song["author"],))
            status.config(text=f"{len(matches)} lagu ditemukan · pilih lagu lalu tekan Tambahkan")
            if len(matches) == 1:
                listing.selection_set(str(matches[0]["id"]))
                listing.focus(str(matches[0]["id"]))

        def add_selected(_event=None):
            selected = listing.selection()
            if not selected:
                status.config(text="Pilih salah satu lagu dari daftar terlebih dahulu.")
                return "break"
            song = songs_by_id.get(selected[0])
            if not song:
                return "break"
            self.store.add_item(self.active_schedule["id"], "Lagu", song["title"], song["lyrics"])
            self.refresh_schedule()
            dialog.destroy()
            return "break"

        def focus_first_match(_event=None):
            rows = listing.get_children()
            if rows:
                listing.selection_set(rows[0])
                listing.focus(rows[0])
                listing.focus_set()
            return "break"

        query.trace_add("write", refresh_matches)
        listing.bind("<Double-1>", add_selected)
        search.bind("<Return>", add_selected)
        search.bind("<Down>", focus_first_match)
        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Batal", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="Tambahkan ke Susunan", style="Accent.TButton",
                   command=add_selected).pack(side="right", padx=7)
        refresh_matches()
        search.focus_set()

    def add_verse_to_schedule(self):
        ref = self.verse_ref.get().strip(); text = self.verse_text.get("1.0", "end").strip()
        if not ref or not text:
            self.tabs.select(self.verses_tab)
            messagebox.showinfo("Ayat belum lengkap", "Isi referensi dan teks ayat terlebih dahulu."); return
        self.store.add_item(self.active_schedule["id"], "Ayat", ref, text)
        self.refresh_schedule(); self.tabs.select(self.schedule_tab)

    def add_media_to_schedule(self):
        path = filedialog.askopenfilename(parent=self, title="Pilih media untuk proyeksi",
                                          filetypes=[("PowerPoint, gambar, atau video", "*.pptx *.png *.jpg *.jpeg *.bmp *.gif *.webp *.mp4 *.avi *.mov *.mkv *.wmv *.m4v"),
                                                     ("PowerPoint", "*.pptx"),
                                                     ("Gambar", "*.png *.jpg *.jpeg *.bmp *.gif *.webp"),
                                                     ("Video", "*.mp4 *.avi *.mov *.mkv *.wmv *.m4v")])
        if path:
            if Path(path).suffix.casefold() == ".pptx":
                try:
                    render_powerpoint_slides(path)
                except PowerPointRenderError as error:
                    messagebox.showerror("PowerPoint tidak bisa ditampilkan", str(error))
                    return
            item_id = self.store.add_media(self.active_schedule["id"], path)
            self.refresh_schedule()
            self.schedule_tree.selection_set(str(item_id))
            self.schedule_tree.focus(str(item_id))
            self.tabs.select(self.schedule_tab)

    def add_url_to_schedule(self):
        url = simpledialog.askstring("Tambahkan Website", "Masukkan URL website:", parent=self)
        if not url or not url.strip():
            return
        url = url.strip()
        if not urlparse(url).scheme:
            url = "https://" + url
        parsed = urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            messagebox.showwarning("URL tidak valid", "Masukkan alamat website yang diawali http:// atau https://.", parent=self)
            return
        title = simpledialog.askstring("Nama Website", "Nama yang ditampilkan di susunan:",
                                       initialvalue=parsed.netloc, parent=self)
        if title is None:
            return
        item_id = self.store.add_item(self.active_schedule["id"], "URL", title.strip() or parsed.netloc, url)
        self.refresh_schedule()
        self.schedule_tree.selection_set(str(item_id))
        self.schedule_tree.focus(str(item_id))
        self.tabs.select(self.schedule_tab)

    def move_schedule_item(self, offset):
        selected = self.schedule_tree.selection()
        if not selected:
            return
        self.store.move_item(int(selected[0]), offset, self.active_schedule["id"])
        self.refresh_schedule()
        rows = self.store.items(self.active_schedule["id"])
        moved = next((i for i, row in enumerate(rows) if row["id"] == int(selected[0])), None)
        if moved is not None:
            self.schedule_tree.selection_set(str(rows[moved]["id"]))

    def import_songs(self):
        path = filedialog.askopenfilename(parent=self, title="Impor perpustakaan lagu", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        count = 0
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as source:
                reader = csv.DictReader(source)
                if not {"title", "lyrics"}.issubset(reader.fieldnames or []):
                    raise ValueError("CSV harus memiliki kolom title dan lyrics (author opsional).")
                for row in reader:
                    if row.get("title", "").strip():
                        self.store.save_song(None, row["title"].strip(), row.get("author", "").strip(), row.get("lyrics", "").strip())
                        count += 1
        except (OSError, csv.Error, ValueError) as error:
            messagebox.showerror("Impor gagal", str(error)); return
        self.refresh_songs()
        messagebox.showinfo("Impor selesai", f"{count} lagu berhasil diimpor.")

    def export_songs(self):
        path = filedialog.asksaveasfilename(parent=self, title="Ekspor perpustakaan lagu", defaultextension=".csv",
                                           filetypes=[("CSV", "*.csv")], initialfile="lagu.csv")
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as target:
            writer = csv.DictWriter(target, fieldnames=["title", "author", "lyrics"])
            writer.writeheader()
            for song in self.store.songs():
                writer.writerow({key: song[key] for key in ("title", "author", "lyrics")})
        messagebox.showinfo("Ekspor selesai", "Perpustakaan lagu berhasil diekspor.")

    def delete_song(self):
        selected = self.song_tree.selection()
        if not selected:
            return
        song_id = int(selected[0])
        if not messagebox.askyesno("Hapus lagu", "Hapus lagu ini dari perpustakaan? Item yang sudah ada di jadwal tetap tersimpan."):
            return
        self.store.db.execute("DELETE FROM songs WHERE id=?", (song_id,)); self.store.db.commit()
        self.new_song(); self.refresh_songs()

    def remove_schedule_item(self, _event=None):
        selected = self.schedule_tree.selection()
        if not selected:
            messagebox.showinfo("Pilih item", "Pilih lagu, ayat, atau media pada susunan ibadah terlebih dahulu.")
            return "break"
        item = self.store.db.execute("SELECT title,kind FROM schedule_items WHERE id=? AND schedule_id=?",
                                     (int(selected[0]), self.active_schedule["id"])).fetchone()
        if not item:
            return "break"
        if not messagebox.askyesno("Hapus dari Susunan",
                                   f"Hapus {item['kind'].lower()} ‘{item['title']}’ dari susunan ibadah ini?\n\n"
                                   "Lagu di perpustakaan dan file media asli tidak akan dihapus."):
            return "break"
        self.store.db.execute("DELETE FROM schedule_items WHERE id=? AND schedule_id=?",
                              (int(selected[0]), self.active_schedule["id"]))
        self.store.db.commit()
        self.refresh_schedule()
        return "break"

    def choose_presentation_background(self):
        path = filedialog.askopenfilename(
            parent=self, title="Pilih latar proyeksi",
            filetypes=[("Gambar atau video", "*.png *.jpg *.jpeg *.bmp *.webp *.mp4 *.avi *.mov *.mkv *.wmv *.m4v"),
                       ("Gambar", "*.png *.jpg *.jpeg *.bmp *.webp"),
                       ("Video", "*.mp4 *.avi *.mov *.mkv *.wmv *.m4v"), ("Semua file", "*.*")])
        if not path:
            return
        try:
            import cv2
            if Path(path).suffix.casefold() in {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"}:
                capture = cv2.VideoCapture(path)
                valid = capture.isOpened()
                capture.release()
            else:
                valid = cv2.imread(path, cv2.IMREAD_COLOR) is not None
            if not valid:
                raise RuntimeError("File gambar atau video tidak dapat dibaca.")
            previous_path = self.presentation_background_path
            self.presentation_background_path = path
            self.presentation_background_paused = False
            try:
                self.apply_presentation_style()
            except Exception:
                self.presentation_background_path = previous_path
                raise
        except Exception as error:
            messagebox.showerror("Latar tidak dapat digunakan", str(error), parent=self)

    def choose_presentation_background_color(self):
        chosen = colorchooser.askcolor(color=self.presentation_background_color,
                                       title="Warna latar proyeksi", parent=self)
        if chosen and chosen[1]:
            self.presentation_background_color = chosen[1]
            self.apply_presentation_style()

    def choose_presentation_text_color(self):
        chosen = colorchooser.askcolor(color=self.presentation_text_color,
                                       title="Warna teks proyeksi", parent=self)
        if chosen and chosen[1]:
            self.presentation_text_color = chosen[1]
            self.apply_presentation_style()

    def _presentation_stroke_style(self):
        if self.presentation_stroke_enabled:
            return self.presentation_stroke_color, self.presentation_stroke_width
        return None, 0

    def _reuse_dialog(self, existing, factory):
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return existing
        return factory()

    def open_presentation_stroke_dialog(self):
        self._presentation_stroke_dialog = self._reuse_dialog(
            self._presentation_stroke_dialog,
            lambda: TextStrokeDialog(self, "Garis Tepi Teks Proyeksi", self.presentation_text_color,
                                     self.presentation_stroke_enabled, self.presentation_stroke_color,
                                     self.presentation_stroke_width, self.apply_presentation_stroke))

    def apply_presentation_stroke(self, state):
        self.presentation_stroke_enabled = bool(state["enabled"])
        self.presentation_stroke_color = state["color"]
        self.presentation_stroke_width = int(state["width"])
        self.presentation_stroke_button.config(
            text=f"Garis Teks · {self.presentation_stroke_color}"
            if self.presentation_stroke_enabled else "Garis Teks…")
        self.apply_presentation_style()

    def reset_presentation_style(self):
        self.presentation_background_path = None
        self.presentation_background_color = "#000000"
        self.presentation_text_color = "#ffffff"
        self.presentation_background_loop = True
        self.presentation_stroke_enabled = False
        self.presentation_stroke_color = DEFAULT_STROKE_COLOR
        self.presentation_stroke_width = DEFAULT_STROKE_WIDTH
        if hasattr(self, "presentation_stroke_button"):
            self.presentation_stroke_button.config(text="Garis Teks…")
        self.background_loop_var.set(True)
        self.apply_presentation_style()

    @staticmethod
    def _is_video_path(path):
        return bool(path and Path(path).suffix.casefold() in
                    {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"})

    def set_transition_style(self, _event=None):
        self.transition_style = self.transition_choice.get() or "Potong"
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.transition_style = self.transition_style

    def set_background_loop(self):
        self.presentation_background_loop = bool(self.background_loop_var.get())
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.set_background_video_loop(self.presentation_background_loop)

    def set_youtube_fullscreen(self):
        self.youtube_fullscreen = bool(self.youtube_fullscreen_var.get())
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.youtube_fullscreen = self.youtube_fullscreen

    def toggle_background_video(self):
        if not self._is_video_path(self.presentation_background_path):
            messagebox.showinfo("Latar video belum dipilih", "Pilih file video pada menu Latar Gambar / Video terlebih dahulu.", parent=self)
            return
        if self.projector_window and self.projector_window.winfo_exists():
            paused = self.projector_window.toggle_background_video_pause()
        else:
            self.presentation_background_paused = not getattr(self, "presentation_background_paused", False)
            paused = self.presentation_background_paused
            
        # Aman dari AttributeError jika background_pause_button belum dibuat
        if hasattr(self, "background_pause_button"):
            self.background_pause_button.config(text="Putar Latar" if paused else "Jeda Latar")

    def apply_presentation_style(self):
        if self.presentation_background_path:
            label = Path(self.presentation_background_path).name
            loop_suffix = " · video berulang" if Path(label).suffix.casefold() in {
                ".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"} else ""
            status = f"Latar: {label}{loop_suffix} · teks {self.presentation_text_color}"
        else:
            status = f"Latar {self.presentation_background_color} · teks {self.presentation_text_color}"
        if self.presentation_stroke_enabled:
            status += f" · garis tepi {self.presentation_stroke_color} {self.presentation_stroke_width}px"
        self.presentation_style_status.config(text=status)
        if self.projector_window and self.projector_window.winfo_exists():
            self.projector_window.set_style(self.presentation_background_path,
                                            self.presentation_background_color,
                                            self.presentation_text_color,
                                            loop=self.presentation_background_loop,
                                            stroke_color=self.presentation_stroke_color
                                            if self.presentation_stroke_enabled else None,
                                            stroke_width=self.presentation_stroke_width)
            if getattr(self, "presentation_background_paused", False):
                self.projector_window.set_background_video_paused(True)
        else:
            self.preview_schedule_item()
        is_video = self._is_video_path(self.presentation_background_path)
        if hasattr(self, "background_pause_button"):
            self.background_pause_button.config(state="normal" if is_video else "disabled",
                                                text="Putar Latar" if getattr(self, "presentation_background_paused", False) else "Jeda Latar")
        if hasattr(self, "background_loop_var"):
            self.background_loop_var.set(self.presentation_background_loop)

    def _preview_background_frame(self):
        if not self.presentation_background_path:
            return None
        try:
            import cv2
            path = self.presentation_background_path
            if Path(path).suffix.casefold() in {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"}:
                capture = cv2.VideoCapture(path)
                ok, frame = capture.read()
                capture.release()
                return frame if ok else None
            return cv2.imread(path, cv2.IMREAD_COLOR)
        except Exception:
            return None

    def preview_schedule_item(self, _event=None):
        selected = self.schedule_tree.selection()
        if not selected:
            self.preview_item_id = None
            self.preview_slides = []
            self._clear_preview_parts()
            self.preview.show_text("Pilih item susunan untuk melihat preview", color="white",
                                   background="#112233", font=("Segoe UI", 16, "bold"))
            self.preview_position.config(text="")
            return
        row = self.store.db.execute("SELECT * FROM schedule_items WHERE id=?", (int(selected[0]),)).fetchone()
        if not row:
            return
        self.preview_item_id = int(row["id"])
        self.preview_slide_index = 0
        if row["kind"] == "URL":
            youtube_note = ""
            if self.youtube_fullscreen and youtube_embed_url(row["content"]):
                youtube_note = "\n\nYouTube · player penuh otomatis"
            self.preview_slides = [{"image": "", "text": f"Website akan tampil di layar proyektor\n\n{row['content']}"
                                                         f"{youtube_note}\n\nEsc menutup website"}]
        elif row["media_path"] and Path(row["media_path"]).suffix.casefold() == ".pptx":
            try:
                images = render_powerpoint_slides(row["media_path"], (700, 450))
                self.preview_slides = [{"image": image, "text": ""} for image in images]
            except PowerPointRenderError as error:
                self.preview_slides = []
                self._clear_preview_parts()
                self.preview.show_text(f"PowerPoint tidak dapat ditampilkan\n{error}", color="white",
                                       background="#112233", font=("Segoe UI", 16, "bold"))
                self.preview_position.config(text=row["title"])
                return
        elif row["media_path"]:
            video_extensions = {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"}
            if Path(row["media_path"]).suffix.casefold() in video_extensions:
                try:
                    import cv2
                    capture = cv2.VideoCapture(row["media_path"])
                    ok, frame = capture.read()
                    capture.release()
                    if not ok:
                        raise RuntimeError("Video tidak dapat dibaca.")
                    self.preview_slides = [{"image": "", "frame": frame, "text": ""}]
                except Exception as error:
                    self.preview_slides = []
                    self._clear_preview_parts()
                    self.preview.show_text(f"Video tidak dapat dipratinjau\n{error}", color="white",
                                           background="black", font=("Segoe UI", 16, "bold"))
                    self.preview_position.config(text=row["title"])
                    return
            else:
                self.preview_slides = [{"image": row["media_path"], "text": ""}]
        else:
            slides = self.make_slides(row["content"]) or [row["title"]]
            self.preview_slides = [{"image": "", "text": slide} for slide in slides]
        self._populate_preview_parts()
        self.render_schedule_preview()

    def render_schedule_preview(self):
        if not self.preview_slides or self.preview_item_id is None:
            return
        row = self.store.db.execute("SELECT title FROM schedule_items WHERE id=?", (self.preview_item_id,)).fetchone()
        if not row:
            return
        slide = self.preview_slides[self.preview_slide_index]
        total = len(self.preview_slides)
        self.preview_position.config(text=f"{row['title']} · {self.preview_slide_index + 1}/{total}")
        if slide.get("frame") is not None:
            self.preview_image = frame_to_photo(
                slide["frame"], (max(200, self.preview.winfo_width() - 24),
                                 max(150, self.preview.winfo_height() - 24)), self.preview)
            self.preview.show_photo(self.preview_image, background="black")
            return
        if slide["image"]:
            try:
                self.preview_image = self.load_image(slide["image"], (max(200, self.preview.winfo_width() - 24),
                                                                      max(150, self.preview.winfo_height() - 24)))
                self.preview.show_photo(self.preview_image, background="black")
            except Exception as error:
                self.preview.show_text(f"Media tidak dapat dibuka\n{error}", color="white",
                                       background="black", font=("Segoe UI", 16, "bold"))
            return
        background_frame = self._preview_background_frame()
        width = max(200, self.preview.winfo_width() - 32)
        height = max(150, self.preview.winfo_height() - 32)
        preview_background = frame_to_photo(background_frame, (width, height), self.preview) if background_frame is not None else None
        self.preview_image = preview_background
        font_size = fit_text_font(self.preview, slide["text"],
                                  max(14, min(42, width // 18, height // 5)), 9,
                                  width, height, padding=12)
        stroke_color, stroke_width = self._presentation_stroke_style()
        self.preview.show_text(slide["text"], color=self.presentation_text_color,
                               background=self.presentation_background_color,
                               font=("Segoe UI", font_size, "bold"), wrap=width,
                               photo=preview_background, stroke_color=stroke_color,
                               stroke_width=stroke_width)

    def step_preview_slide(self, offset):
        if not self.preview_slides:
            return
        self.preview_slide_index = max(0, min(self.preview_slide_index + offset,
                                              len(self.preview_slides) - 1))
        self.render_schedule_preview()

    def update_projector_preview(self, slide, index, total, blackout=False, background_frame=None,
                                 background_color="#000000", text_color="#ffffff", video_frame=None,
                                 clear=False, stroke_color=None, stroke_width=0):
        if blackout or slide is None:
            clear_background = frame_to_photo(
                background_frame, (max(200, self.live_preview.winfo_width() - 24),
                                   max(150, self.live_preview.winfo_height() - 24)), self.live_preview
            ) if clear and background_frame is not None else None
            self.live_preview_image = clear_background
            if clear_background is not None:
                self.live_preview.show_photo(clear_background, background="black")
            else:
                self.live_preview.clear("black")
            self.live_position.config(text="Tampilan dibersihkan" if clear else f"Layar hitam · {index + 1}/{total}")
        elif video_frame is not None:
            self.live_preview_image = frame_to_photo(
                video_frame, (max(200, self.live_preview.winfo_width() - 24),
                              max(150, self.live_preview.winfo_height() - 24)), self.live_preview)
            self.live_preview.show_photo(self.live_preview_image, background="black")
            self.live_position.config(text=f"{slide['title']} · video berulang")
        elif slide.get("url"):
            self.live_preview_image = None
            message = f"Website dibuka pada layar proyektor\n\n{slide['url']}"
            if self.youtube_fullscreen and youtube_embed_url(slide["url"]):
                message += "\n\nYouTube · player penuh otomatis"
            message += "\n\nEsc menutup website"
            self.live_preview.show_text(message, color="white", background="#101820",
                                        font=("Segoe UI", 15, "bold"))
            self.live_position.config(text=f"Website · {slide['title']}")
        elif slide["image"]:
            try:
                preview_size = (max(200, self.live_preview.winfo_width() - 24),
                                max(150, self.live_preview.winfo_height() - 24))
                self.live_preview_image = self.load_image(slide["image"], preview_size)
                self.live_preview.show_photo(self.live_preview_image, background="black")
                self.live_position.config(text=f"{slide['title']} · {index + 1}/{total}")
            except Exception as error:
                self.live_preview.show_text(f"Tampilan Live tidak tersedia\n{error}", color="white",
                                            background="black", font=("Segoe UI", 16, "bold"))
        else:
            width = max(200, self.live_preview.winfo_width() - 48)
            height = max(150, self.live_preview.winfo_height() - 48)
            font_size = fit_text_font(self.live_preview, slide["text"],
                                      max(14, min(44, width // 20, height // 5)), 9,
                                      width, height, padding=12)
            preview_background = frame_to_photo(background_frame, (width, height), self.live_preview) if background_frame is not None else None
            self.live_preview_image = preview_background
            self.live_preview.show_text(slide["text"], color=text_color, background=background_color,
                                        font=("Segoe UI", font_size, "bold"), wrap=width,
                                        photo=preview_background, stroke_color=stroke_color,
                                        stroke_width=stroke_width)
            self.live_position.config(text=f"{slide['title']} · {index + 1}/{total}")

    def go_live_preview(self):
        projector = self.projector_window
        if not projector or not projector.winfo_exists():
            messagebox.showinfo("Proyeksi belum aktif", "Mulai proyeksi terlebih dahulu. Pilih item susunan untuk melihat preview, lalu tampilkan ke Live.", parent=self)
            return
        selected = self.schedule_tree.selection()
        if not selected:
            messagebox.showinfo("Pilih item", "Pilih item susunan yang ingin ditampilkan ke Live.", parent=self)
            return
        group = next((group for group in projector.slide_groups
                      if group.get("item_id") == int(selected[0])), None)
        if group:
            offset = self.preview_slide_index if self.preview_item_id == int(selected[0]) else 0
            projector.show_index(min(group["start"] + offset, group["end"] - 1))

    @staticmethod
    def make_slides(content):
        return [part.strip() for part in content.split("\n\n") if part.strip()]

    @staticmethod
    def load_image(path, size):
        try:
            from PIL import Image, ImageTk
            image = Image.open(path)
            image.thumbnail(size)
            return ImageTk.PhotoImage(image)
        except ImportError:
            try:
                import cv2
                image = cv2.imread(str(path), cv2.IMREAD_UNCHANGED)
                if image is not None:
                    height, width = image.shape[:2]
                    scale = min(size[0] / width, size[1] / height)
                    output_size = (max(1, int(width * scale)), max(1, int(height * scale)))
                    if output_size != (width, height):
                        image = cv2.resize(image, output_size,
                                           interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_LINEAR)
                    ok, encoded = cv2.imencode(".png", image)
                    if ok:
                        data = base64.b64encode(encoded.tobytes()).decode("ascii")
                        return tk.PhotoImage(data=data, format="png")
            except Exception:
                pass
            return tk.PhotoImage(file=path)
#tambahan
    def _clear_preview_parts(self):
        """Menghapus semua daftar bagian di panel bawah."""
        for item in self.parts_tree.get_children():
            self.parts_tree.delete(item)

    def _populate_preview_parts(self):
        """Mengisi daftar bagian dari slide preview yang aktif ke dalam Treeview."""
        self._clear_preview_parts()
        for index, slide in enumerate(self.preview_slides):
            # Ambil potongan teks pertama untuk ditampilkan di daftar (atau beri label media)
            text_preview = slide["text"].replace("\n", " ")[:80] if slide.get("text") else "(Media / Gambar / Video)"
            self.parts_tree.insert("", "end", iid=str(index), values=(str(index + 1), text_preview))
        
        # Otomatis pilih baris pertama jika ada
        if self.preview_slides:
            self.parts_tree.selection_set("0")
            self.parts_tree.focus("0")

    def select_preview_part(self, event=None):
        """Merespons ketika operator mengklik salah satu baris di daftar teks."""
        selected = self.parts_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        self.preview_slide_index = index
        self.render_schedule_preview()

    def send_preview_part_live(self, event=None):
        """Tampilkan ke proyektor langsung ketika operator klik ganda pada baris."""
        self.select_preview_part()
        self.go_live_preview()

    def edit_preview_slide(self):
        """Membuka dialog untuk memperbaiki teks secara cepat (typo) pada slide yang dipilih."""
        selected = self.parts_tree.selection()
        if not selected or not self.preview_item_id:
            messagebox.showinfo("Pilih Bagian", "Pilih bagian teks yang ingin diedit dari daftar di panel 'BAGIAN ITEM'.", parent=self)
            return
        
        index = int(selected[0])
        slide = self.preview_slides[index]

        if slide.get("image") or slide.get("url") or slide.get("frame") is not None:
            messagebox.showinfo("Tidak Dapat Diedit", "Bagian yang berupa media (gambar/video/URL) tidak dapat diedit teksnya.", parent=self)
            return

        row = self.store.db.execute("SELECT * FROM schedule_items WHERE id=?", (self.preview_item_id,)).fetchone()
        if not row:
            return

        dialog = tk.Toplevel(self)
        dialog.title("Edit Teks Slide Cepat")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("500x350")
        dialog.configure(bg="#f2f5f8")

        body = ttk.Frame(dialog, padding=15)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="Edit Teks Slide:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
        text_editor = tk.Text(body, wrap="word", font=("Segoe UI", 12), height=10)
        text_editor.pack(fill="both", expand=True)
        text_editor.insert("1.0", slide["text"])

        def save_edit():
            new_text = text_editor.get("1.0", "end").strip()
            slides_texts = self.make_slides(row["content"]) or [row["title"]]

            # Update teks di indeks bagian yang sesuai
            if index < len(slides_texts):
                slides_texts[index] = new_text
            else:
                slides_texts.append(new_text)

            new_content = "\n\n".join(slides_texts)
            
            # Simpan ke database
            self.store.update_item(row["id"], row["schedule_id"], row["title"], new_content, row["operator_note"])

            # Muat ulang pratinjau agar perubahan terlihat
            self.preview_schedule_item()
            
            # Kembalikan seleksi ke baris yang tadi diedit
            if str(index) in self.parts_tree.get_children():
                self.parts_tree.selection_set(str(index))
                self.parts_tree.see(str(index))
                self.select_preview_part()

            dialog.destroy()

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Batal", command=dialog.destroy).pack(side="right")
        ttk.Button(actions, text="Simpan", style="Accent.TButton", command=save_edit).pack(side="right", padx=8)
        text_editor.focus_set()

#tambahan

    def open_projector(self):
        item = None
        selected = self.schedule_tree.selection()
        if selected:
            item = self.store.db.execute("SELECT * FROM schedule_items WHERE id=?", (int(selected[0]),)).fetchone()
        if not item:
            rows = self.store.items(self.active_schedule["id"])
            if rows: item = rows[0]
        if not item:
            messagebox.showinfo("Susunan kosong", "Tambahkan lagu atau ayat ke susunan ibadah terlebih dahulu."); return
        rows = self.store.items(self.active_schedule["id"])
        displays = available_displays()
        target_display = displays[1] if len(displays) > 1 else displays[0]
        screen_size = (target_display[2], target_display[3])
        combined = []
        slide_groups = []
        selected_id = int(selected[0]) if selected else (rows[0]["id"] if rows else None)
        start_index = 0
        for row in rows:
            item_start = len(combined)
            if row["id"] == selected_id:
                start_index = len(combined)
            if row["media_path"] and Path(row["media_path"]).suffix.casefold() == ".pptx":
                try:
                    ppt_images = render_powerpoint_slides(row["media_path"], screen_size)
                    combined.extend({"text": "", "image": image,
                                     "title": f"{row['title']} · {index}/{len(ppt_images)}"}
                                    for index, image in enumerate(ppt_images, start=1))
                except PowerPointRenderError as error:
                    messagebox.showerror("PowerPoint tidak bisa ditampilkan", str(error))
                    return
            elif row["media_path"]:
                extension = Path(row["media_path"]).suffix.casefold()
                combined.append({"text": "", "image": row["media_path"],
                                 "video": extension in {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"},
                                 "title": row["title"]})
            elif row["kind"] == "URL":
                parsed_url = urlparse(row["content"])
                if parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc:
                    messagebox.showerror("URL tidak valid", f"Alamat website tidak valid:\n{row['content']}", parent=self)
                    return
                combined.append({"text": "", "image": "", "url": row["content"], "title": row["title"]})
            else:
                combined.extend({"text": slide, "image": "", "title": row["title"]}
                                for slide in (self.make_slides(row["content"]) or [row["title"]]))
            if len(combined) > item_start:
                slide_groups.append({"start": item_start, "end": len(combined), "title": row["title"],
                                     "item_id": row["id"]})
        selected_group = next((group for group in slide_groups if group.get("item_id") == selected_id), None)
        if selected_group and self.preview_item_id == selected_id:
            start_index = min(selected_group["start"] + self.preview_slide_index,
                              selected_group["end"] - 1)
        try:
            self.projector_window = Projector(self, self.active_schedule["title"], combined,
                                              start_index, slide_groups, target_display=target_display,
                                              background_path=self.presentation_background_path,
                                              background_color=self.presentation_background_color,
                                              text_color=self.presentation_text_color,
                                              background_loop=self.presentation_background_loop,
                                              transition_style=self.transition_style,
                                              stroke_color=self.presentation_stroke_color
                                              if self.presentation_stroke_enabled else None,
                                              stroke_width=self.presentation_stroke_width,
                                              youtube_fullscreen=self.youtube_fullscreen)
            self.projector_stop_button.state(["!disabled"])
        except Exception as error:
            self.projector_window = None
            messagebox.showerror("Proyeksi tidak dapat dimulai", str(error), parent=self)


class Projector(tk.Toplevel):
    def __init__(self, parent, title, slides, start_index=0, slide_groups=None, target_display=None,
                 background_path=None, background_color="#000000", text_color="#ffffff",
                 background_loop=True, transition_style="Potong", stroke_color=None,
                 stroke_width=0, youtube_fullscreen=True):
        super().__init__(parent)
        self.parent = parent
        self.slides = slides
        self.slide_groups = slide_groups or []
        displays = available_displays()
        self.target_display = target_display or (displays[1] if len(displays) > 1 else displays[0])
        self.number_buffer = ""
        self.number_after_id = None
        self.index = min(start_index, len(slides) - 1)
        self.blackout = False
        self.cleared = False
        self.background_loop = bool(background_loop)
        self.background_paused = False
        self.transition_style = transition_style
        self._transition_after_id = None
        self._fade_overlay = None
        self._browser_process = None
        self._browser_profile = None
        self._browser_url = None
        self._browser_position_after_id = None
        self._url_watch_after_id = None
        self._url_escape_down = False
        self._url_restoring = False
        self.youtube_fullscreen = bool(youtube_fullscreen)
        self._closing = False
        self._global_bindings = []
        self.background_path = None
        self.background_color = background_color
        self.text_color = text_color
        self.stroke_color = stroke_color
        self.stroke_width = int(stroke_width or 0)
        self.background_image = None
        self.background_capture = None
        self.background_photo = None
        self.background_after_id = None
        self.slide_capture = None
        self.slide_video_after_id = None
        self.slide_video_photo = None
        self._text_font_size = None
        self._resize_after_id = None
        self.configure(bg="black")
        left, top, width, height = self.target_display
        self.geometry(f"{width}x{height}{left:+d}{top:+d}")
        self.stage = SlideCanvas(self, bg="black")
        self.title_bar = tk.Frame(self, bg="black")
        self.title_bar.pack(side="top", fill="x", padx=18, pady=(12, 0))
        tk.Label(self.title_bar, text=title, bg="black", fg="#8da0b3", font=("Segoe UI", 10)).pack(side="left")
        self.exit_button = tk.Button(self.title_bar, text="Keluar dari Proyeksi  (Esc)", command=self.close_projector,
                                     bg="#263445", fg="white", activebackground="#40516a",
                                     activeforeground="white", relief="flat", padx=12, pady=6)
        self.exit_button.pack(side="right")
        self.stage.pack(fill="both", expand=True, padx=40, pady=30)
        self.hint = tk.Label(self, text="← → slide · nomor + Enter lompat · Home/End · B blackout · C clear · Esc keluar", bg="black", fg="#8da0b3", font=("Segoe UI", 11))
        self.hint.pack(side="bottom", pady=18)
        self._global_bindings.append(("<Right>", self.bind_all("<Right>", self.next_slide, add="+")))
        self._global_bindings.append(("<space>", self.bind_all("<space>", self.next_slide, add="+")))
        self._global_bindings.append(("<Left>", self.bind_all("<Left>", self.prev_slide, add="+")))
        self._global_bindings.append(("<Escape>", self.bind_all("<Escape>", self.close_projector, add="+")))
        self._global_bindings.append(("<Home>", self.bind_all("<Home>", self.first_slide, add="+")))
        self._global_bindings.append(("<End>", self.bind_all("<End>", self.last_slide, add="+")))
        self._global_bindings.append(("b", self.bind_all("b", self.toggle_blackout, add="+")))
        self._global_bindings.append(("c", self.bind_all("c", self.clear_output, add="+")))
        for digit in "0123456789":
            callback = lambda _event, value=digit: self.queue_number_jump(value)
            self._global_bindings.append((f"<KeyPress-{digit}>", self.bind_all(f"<KeyPress-{digit}>", callback, add="+")))
            self._global_bindings.append((f"<KP_{digit}>", self.bind_all(f"<KP_{digit}>", callback, add="+")))
        self._global_bindings.append(("<Return>", self.bind_all("<Return>", self.commit_number_jump, add="+")))
        self.bind("<Escape>", self.close_projector, add="+")
        self.stage.bind("<Escape>", self.close_projector, add="+")
        self.stage.bind("<Configure>", self._stage_resized, add="+")
        self.bind("<F11>", self.toggle_fullscreen)
        self.protocol("WM_DELETE_WINDOW", self.close_projector)
        try:
            self.set_style(background_path, background_color, text_color, redraw=False, loop=background_loop,
                           stroke_color=stroke_color, stroke_width=stroke_width)
        except Exception:
            self.destroy()
            raise
        self.after_idle(self._take_focus)
        self.show_slide()

    def set_style(self, background_path, background_color, text_color, redraw=True, loop=None,
                  stroke_color=None, stroke_width=0):
        new_image = None
        new_capture = None
        if background_path:
            import cv2
            extension = Path(background_path).suffix.casefold()
            if extension in {".mp4", ".avi", ".mov", ".mkv", ".wmv", ".m4v"}:
                new_capture = VideoCaptureSource(background_path, self.background_loop if loop is None else loop)
            else:
                new_image = cv2.imread(str(background_path), cv2.IMREAD_COLOR)
                if new_image is None:
                    raise RuntimeError("File gambar latar tidak dapat dibaca.")
        old_capture = self.background_capture
        self.background_path = background_path
        self.background_image = new_image
        self.background_capture = new_capture
        if loop is not None:
            self.background_loop = bool(loop)
        self.background_paused = False
        self.background_color = background_color
        self.text_color = text_color
        self.stroke_color = stroke_color
        self.stroke_width = int(stroke_width or 0)
        self.background_photo = None
        if old_capture is not None:
            old_capture.close()
        if redraw:
            self.show_slide()
        if new_capture is not None and self.background_after_id is None:
            self.background_after_id = self.after(50, self._update_background_video)

    def _stage_size(self):
        width, height = self.stage.winfo_width(), self.stage.winfo_height()
        if width < 10 or height < 10:
            width, height = self.target_display[2], self.target_display[3]
        return width, height

    def _paint_text_slide(self, slide, font_size=None):
        width, height = self._stage_size()
        text_width = max(400, width - 40)
        text_height = max(250, height - 40)
        if font_size is None:
            font_size = fit_text_font(
                self.stage, slide["text"], max(24, min(112, self.target_display[3] // 14)), 12,
                text_width, text_height, padding=18)
        self._text_font_size = font_size
        background_frame = self._current_background_frame()
        self.background_photo = frame_to_photo(
            background_frame, (self.target_display[2], self.target_display[3]), self
        ) if background_frame is not None else None
        self.stage.show_text(slide["text"], color=self.text_color, background=self.background_color,
                             font=("Segoe UI", font_size, "bold"), wrap=text_width,
                             photo=self.background_photo, stroke_color=self.stroke_color,
                             stroke_width=self.stroke_width)

    def _stage_resized(self, _event=None):
        if self._resize_after_id:
            try:
                self.after_cancel(self._resize_after_id)
            except tk.TclError:
                pass
        self._resize_after_id = self.after(150, self._repaint_text_after_resize)

    def _repaint_text_after_resize(self):
        self._resize_after_id = None
        if self._closing or not self.winfo_exists() or self.blackout or self.cleared or not self.slides:
            return
        slide = self.slides[self.index]
        if slide["image"] or slide.get("url"):
            return
        self._paint_text_slide(slide)

    def _stage_message(self, text, background="black", color="white", font_size=24):
        self.stage.show_text(text, color=color, background=background,
                             font=("Segoe UI", font_size, "bold"))

    def set_background_video_loop(self, enabled):
        self.background_loop = bool(enabled)
        if self.background_capture is not None:
            self.background_capture.set_loop(enabled)

    def set_background_video_paused(self, paused):
        self.background_paused = bool(paused)
        if self.background_capture is not None:
            self.background_capture.set_paused(paused)

    def toggle_background_video_pause(self):
        self.set_background_video_paused(not self.background_paused)
        return self.background_paused

    def _cancel_transition(self):
        if self._transition_after_id:
            try:
                self.after_cancel(self._transition_after_id)
            except tk.TclError:
                pass
            self._transition_after_id = None
        if self._fade_overlay is not None:
            try:
                self._fade_overlay.destroy()
            except tk.TclError:
                pass
            self._fade_overlay = None

    def _fade_step(self, value, target, steps, phase):
        if self._closing or not self.winfo_exists() or self._fade_overlay is None:
            return
        try:
            alpha = value / steps
            self._fade_overlay.attributes("-alpha", alpha)
        except tk.TclError:
            self._cancel_transition()
            self.index = target
            self.blackout = self.cleared = False
            self.show_slide()
            return
        if phase == "out" and value < steps:
            self._transition_after_id = self.after(25, self._fade_step, value + 1, target, steps, phase)
            return
        if phase == "out":
            self.index = target
            self.blackout = self.cleared = False
            self.show_slide()
            self._transition_after_id = self.after(25, self._fade_step, steps - 1, target, steps, "in")
            return
        if value > 0:
            self._transition_after_id = self.after(25, self._fade_step, value - 1, target, steps, phase)
        else:
            self._cancel_transition()
            self.focus_force()

    def navigate_to_index(self, index):
        if not self.slides:
            return
        target = max(0, min(int(index), len(self.slides) - 1))
        if target == self.index and not self.blackout and not self.cleared:
            return
        self._cancel_transition()
        if self.transition_style != "Fade":
            self.index = target
            self.blackout = self.cleared = False
            self.show_slide()
            return
        try:
            overlay = tk.Toplevel(self)
            overlay.overrideredirect(True)
            overlay.attributes("-topmost", True)
            overlay.configure(bg="black")
            overlay.geometry(f"{self.target_display[2]}x{self.target_display[3]}"
                             f"{self.target_display[0]:+d}{self.target_display[1]:+d}")
            tk.Frame(overlay, bg="black").pack(fill="both", expand=True)
            overlay.attributes("-alpha", 0.0)
            overlay.lift()
            self._fade_overlay = overlay
            self._transition_after_id = self.after(25, self._fade_step, 1, target, 8, "out")
        except tk.TclError:
            self.index = target
            self.blackout = self.cleared = False
            self.show_slide()

    def clear_output(self):
        self._cancel_transition()
        self.blackout = False
        self.cleared = True
        self.show_slide()
        return "break"

    def _current_background_frame(self):
        if self.background_capture is not None:
            return self.background_capture.latest_frame()
        return self.background_image

    def _browser_executable(self):
        candidates = [shutil.which("msedge"), shutil.which("chrome")]
        for variable, relative in (("PROGRAMFILES(X86)", "Microsoft/Edge/Application/msedge.exe"),
                                   ("PROGRAMFILES", "Microsoft/Edge/Application/msedge.exe"),
                                   ("PROGRAMFILES", "Google/Chrome/Application/chrome.exe"),
                                   ("LOCALAPPDATA", "Microsoft/Edge/Application/msedge.exe"),
                                   ("LOCALAPPDATA", "Google/Chrome/Application/chrome.exe")):
            base = os.environ.get(variable)
            if base:
                candidates.append(str(Path(base) / relative))
        return next((candidate for candidate in candidates if candidate and Path(candidate).is_file()), None)

    def _show_url_slide(self, slide):
        url = slide.get("url", "")
        embed_url = youtube_embed_url(url) if self.youtube_fullscreen else None
        player_url = None
        if embed_url:
            player_url = youtube_player_page_url(youtube_player_html(embed_url))
            if player_url is None:
                embed_url = None
        browser_key = embed_url or url
        if self._browser_process is not None and self._browser_url == browser_key:
            return
        self._close_url_window()
        executable = self._browser_executable()
        if not executable:
            self.deiconify()
            self._stage_message("Website tidak dapat ditampilkan\nMicrosoft Edge atau Google Chrome tidak ditemukan.")
            if self.parent.winfo_exists():
                self.parent.update_projector_preview(slide, self.index, len(self.slides))
            return
        self._browser_profile = tempfile.mkdtemp(prefix="sahabat-web-")
        target_url = player_url or url
        extra_flags = []
        if player_url:
            extra_flags.append("--autoplay-policy=no-user-gesture-required")
        left, top, width, height = self.target_display
        command = [executable, f"--app={target_url}", "--new-window", "--start-fullscreen",
                   "--no-first-run", "--no-default-browser-check", "--disable-sync",
                   "--disable-features=msImplicitSignin,msEdgeFirstRunExperience,msSyncPromo",
                   f"--user-data-dir={self._browser_profile}", f"--window-position={left},{top}",
                   f"--window-size={width},{height}", *extra_flags]
        if os.name == "nt":
            try:
                ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)  # ASFW_ANY
            except (AttributeError, OSError):
                pass
        try:
            self._browser_process = subprocess.Popen(
                command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            self._browser_url = browser_key
            self.withdraw()
            if self._browser_position_after_id:
                self.after_cancel(self._browser_position_after_id)
            self._browser_position_after_id = self.after(900, self._position_url_window, 5)
            self._start_url_watch()
            if self.parent.winfo_exists():
                self.parent.update_projector_preview(slide, self.index, len(self.slides))
                self.parent.live_position.config(text=f"Website aktif · {slide['title']} · Esc menutup website")
        except OSError as error:
            self._close_url_window()
            self.deiconify()
            self._stage_message(f"Website tidak dapat dibuka\n{error}")

    def _escape_pressed(self):
        if os.name != "nt":
            return False
        try:
            return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)  # VK_ESCAPE
        except (AttributeError, OSError):
            return False

    def _url_window_is_foreground(self):
        process = self._browser_process
        if process is None or os.name != "nt":
            return False
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return False
        pid = wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value == process.pid

    def _start_url_watch(self):
        """Watch the website window so Esc can close it and bring the projector back.

        The browser is a separate process and has the keyboard focus, so Tk's
        normal bindings never see Esc while a website is on screen.
        """
        self._stop_url_watch()
        self._url_escape_down = self._escape_pressed()
        self._url_watch_after_id = self.after(30, self._watch_url_window)

    def _stop_url_watch(self):
        if self._url_watch_after_id:
            try:
                self.after_cancel(self._url_watch_after_id)
            except tk.TclError:
                pass
            self._url_watch_after_id = None

    def _watch_url_window(self):
        self._url_watch_after_id = None
        process = self._browser_process
        if self._closing or process is None:
            return
        if process.poll() is not None:
            # The website window was closed (Esc, X, or the page itself).
            self._restore_after_website()
            return
        pressed = self._escape_pressed()
        if pressed and not self._url_escape_down and self._url_window_is_foreground():
            self._restore_after_website()
            return
        self._url_escape_down = pressed
        self._url_watch_after_id = self.after(30, self._watch_url_window)

    def _restore_after_website(self):
        """Close the website and give the projector window back to the operator."""
        if self._url_restoring:
            return
        self._url_restoring = True
        try:
            self._stop_url_watch()
            self._close_url_window()
            if self._closing or not self.winfo_exists():
                return
            self.deiconify()
            slide = self.slides[self.index] if self.slides else None
            message = "Tampilan website ditutup"
            if slide and slide.get("url"):
                message += f"\n{slide['url']}"
            message += "\n\n← → slide  ·  Esc keluar dari proyeksi"
            self._stage_message(message, font_size=20)
            self.hint.config(text="Website ditutup  ·  ← → slide  ·  Esc keluar dari proyeksi")
            if self.parent.winfo_exists():
                self.parent.live_position.config(text="Website ditutup · operator kembali aktif")
            self.after_idle(lambda: self._take_focus(force=True))
        finally:
            self._url_restoring = False

    @staticmethod
    def _force_foreground_window(hwnd):
        """Bring the website window to the front so keyboard input reaches it."""
        if os.name != "nt" or not hwnd:
            return False
        user32 = ctypes.windll.user32
        try:
            user32.AllowSetForegroundWindow(0xFFFFFFFF)  # ASFW_ANY
        except (AttributeError, OSError):
            pass
        if user32.GetForegroundWindow() == hwnd:
            return True
        current_thread = ctypes.windll.kernel32.GetCurrentThreadId()
        foreground = user32.GetForegroundWindow()
        foreground_thread = user32.GetWindowThreadProcessId(foreground, None) if foreground else 0
        attached = False
        try:
            if foreground_thread and foreground_thread != current_thread:
                attached = bool(user32.AttachThreadInput(current_thread, foreground_thread, True))
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
        finally:
            if attached:
                user32.AttachThreadInput(current_thread, foreground_thread, False)
        return user32.GetForegroundWindow() == hwnd

    def _position_url_window(self, attempts_left):
        self._browser_position_after_id = None
        process = self._browser_process
        if self._closing or process is None or process.poll() is not None:
            return
        if os.name == "nt":
            hwnd_found = []
            callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

            def collect(hwnd, _data):
                pid = wintypes.DWORD()
                ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value == process.pid and ctypes.windll.user32.IsWindowVisible(hwnd):
                    hwnd_found.append(hwnd)
                    return False
                return True

            callback = callback_type(collect)
            ctypes.windll.user32.EnumWindows(callback, 0)
            if hwnd_found:
                hwnd = hwnd_found[0]
                left, top, width, height = self.target_display
                ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                ctypes.windll.user32.SetWindowPos(hwnd, wintypes.HWND(-1), left, top, width, height,
                                                  0x0040 | 0x0200)  # SHOWWINDOW | FRAMECHANGED
                if not self._force_foreground_window(hwnd) and attempts_left > 0:
                    self._browser_position_after_id = self.after(400, self._position_url_window, attempts_left - 1)
                return
        if attempts_left > 0:
            self._browser_position_after_id = self.after(500, self._position_url_window, attempts_left - 1)

    def _close_url_window(self):
        self._stop_url_watch()
        if self._browser_position_after_id:
            try:
                self.after_cancel(self._browser_position_after_id)
            except tk.TclError:
                pass
            self._browser_position_after_id = None
        process = self._browser_process
        if process is not None:
            if os.name == "nt" and process.poll() is None:
                callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

                def close_window(hwnd, _data):
                    pid = wintypes.DWORD()
                    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value == process.pid:
                        ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE
                    return True

                ctypes.windll.user32.EnumWindows(callback_type(close_window), 0)
            try:
                process.wait(timeout=0.6)
            except subprocess.TimeoutExpired:
                try:
                    process.terminate()
                except OSError:
                    pass
                try:
                    process.wait(timeout=0.6)
                except subprocess.TimeoutExpired:
                    try:
                        process.kill()
                    except OSError:
                        pass
            self._browser_process = None
        self._browser_url = None
        if self._browser_profile:
            shutil.rmtree(self._browser_profile, ignore_errors=True)
            self._browser_profile = None

    def _update_background_video(self):
        self.background_after_id = None
        if self._closing or not self.winfo_exists() or self.background_capture is None:
            return
        frame = self.background_capture.latest_frame()
        current = self.slides[self.index] if self.slides else None
        if frame is not None and not self.blackout and current and not current["image"] and not current.get("url"):
            if self.cleared:
                self.background_photo = frame_to_photo(
                    frame, (self.target_display[2], self.target_display[3]), self)
                background = self.background_color if self.background_path or \
                    self.background_color.casefold() != "#000000" else "black"
                if self.background_photo is not None:
                    self.stage.show_photo(self.background_photo, background=background)
                else:
                    self.stage.clear(background)
            else:
                self._paint_text_slide(current, font_size=self._text_font_size)
            if self.parent.winfo_exists():
                if self.cleared:
                    self.parent.update_projector_preview(None, self.index, len(self.slides), clear=True,
                                                         background_frame=frame)
                else:
                    self.parent.update_projector_preview(
                        current, self.index, len(self.slides), background_frame=frame,
                        background_color=self.background_color, text_color=self.text_color,
                        stroke_color=self.stroke_color, stroke_width=self.stroke_width)
        if self.background_capture is not None and not self._closing:
            self.background_after_id = self.after(50, self._update_background_video)

    def _take_focus(self, force=False):
        if self.winfo_exists() and (force or not (self.slides and self.slides[self.index].get("url"))):
            place_window_on_display(self, self.target_display, fullscreen=True)
            self.focus_force()

    def toggle_fullscreen(self, _event=None):
        is_fullscreen = self.overrideredirect()
        place_window_on_display(self, self.target_display, fullscreen=not is_fullscreen)
        self.focus_force()
        return "break"

    def close_projector(self, _event=None):
        if self._closing:
            return "break"
        self._closing = True
        self._close_url_window()
        self._cancel_transition()
        if self._resize_after_id:
            try:
                self.after_cancel(self._resize_after_id)
            except tk.TclError:
                pass
            self._resize_after_id = None
        if self.background_after_id:
            self.after_cancel(self.background_after_id)
            self.background_after_id = None
        if self.background_capture is not None:
            self.background_capture.close()
            self.background_capture = None
        self._stop_slide_video()
        if self.number_after_id:
            self.after_cancel(self.number_after_id)
            self.number_after_id = None
        for sequence, binding_id in self._global_bindings:
            if binding_id:
                lines = self.tk.call("bind", "all", sequence).split("\n")
                prefix = f'if {{"[{binding_id} '
                remaining = "\n".join(line for line in lines if not line.startswith(prefix))
                self.tk.call("bind", "all", sequence, remaining)
                self.deletecommand(binding_id)
        self._global_bindings.clear()
        if self.winfo_exists():
            self.attributes("-fullscreen", False)
            self.destroy()
        if self.parent.winfo_exists():
            self.parent.deiconify()
            self.parent.lift()
            self.parent.focus_force()
            self.parent.projector_window = None
            if hasattr(self.parent, "projector_stop_button"):
                self.parent.projector_stop_button.state(["disabled"])
            self.parent.live_preview.show_text("Tampilan Live akan terlihat di sini saat proyeksi dimulai",
                                               color="white", background="black",
                                               font=("Segoe UI", 16, "bold"))
            self.parent.live_position.config(text="Belum ada proyeksi aktif")
            self.parent.live_preview_image = None
            self.parent.preview_schedule_item()
        return "break"

    def show_slide(self):
        self._stop_slide_video()
        if self.blackout:
            self._close_url_window()
            self.deiconify()
            self.stage.clear("black")
            self.hint.config(text="Layar hitam  ·  B tampilkan kembali  ·  Esc keluar")
            if self.parent.winfo_exists():
                self.parent.update_projector_preview(None, self.index, len(self.slides), blackout=True)
            return
        slide = self.slides[self.index]
        if self.cleared:
            self._close_url_window()
            self.deiconify()
            background = self.background_color if self.background_path or \
                self.background_color.casefold() != "#000000" else "black"
            if self.background_path or self.background_capture is not None:
                frame = self._current_background_frame()
                self.background_photo = frame_to_photo(
                    frame, (self.target_display[2], self.target_display[3]), self) if frame is not None else None
                if self.background_photo is not None:
                    self.stage.show_photo(self.background_photo, background=background)
                else:
                    self.stage.clear(background)
            else:
                self.stage.clear(background)
            self.hint.config(text="Tampilan dibersihkan · ← → untuk lanjut · Esc keluar")
            if self.parent.winfo_exists():
                self.parent.update_projector_preview(None, self.index, len(self.slides), clear=True,
                                                     background_frame=self._current_background_frame())
            return
        if slide.get("url"):
            self._show_url_slide(slide)
            return
        if self._browser_process is not None:
            self._close_url_window()
            self.deiconify()
            self.after_idle(self._take_focus)
        if slide["image"]:
            self.title_bar.pack_forget()
            self.hint.pack_forget()
            self.stage.pack_configure(padx=0, pady=0)
            if slide.get("video"):
                try:
                    self.slide_capture = VideoCaptureSource(slide["image"])
                    self._stage_message("Memuat video…")
                    self.slide_video_after_id = self.after(33, self._update_slide_video)
                except Exception as error:
                    self._stop_slide_video()
                    self._stage_message(f"Video tidak dapat dibuka\n{error}")
            else:
                try:
                    self.slide_image = SahabatIbadahApp.load_image(
                        slide["image"], (self.target_display[2], self.target_display[3]))
                    self.stage.show_photo(self.slide_image, background="black")
                except Exception as error:
                    self._stage_message(f"Gambar tidak dapat dibuka\n{slide['image']}\n{error}")
        else:
            has_custom_background = bool(self.background_path) or self.background_color.casefold() != "#000000"
            if has_custom_background:
                self.title_bar.pack_forget()
                self.hint.pack_forget()
                self.stage.pack_configure(padx=0, pady=0)
            else:
                if self.title_bar not in self.pack_slaves():
                    self.title_bar.pack(side="top", fill="x", padx=18, pady=(12, 0), before=self.stage)
                if self.hint not in self.pack_slaves():
                    self.hint.pack(side="bottom", pady=18, after=self.stage)
                self.stage.pack_configure(padx=40, pady=30)
            self.update_idletasks()
            self._paint_text_slide(slide)
        group = self.current_slide_group()
        if group:
            part_number = self.index - group["start"] + 1
            part_total = group["end"] - group["start"]
            part_hint = f"Bagian {part_number}/{part_total} · ketik nomor bagian + Enter untuk lompat"
        else:
            part_hint = "ketik nomor bagian + Enter untuk lompat"
        self.hint.config(text=f"{slide['title']}  ·  {part_hint}   ·   ← → slide   ·   Esc keluar")
        if self.parent.winfo_exists():
            self.parent.update_projector_preview(
                slide, self.index, len(self.slides),
                background_frame=self._current_background_frame(),
                background_color=self.background_color, text_color=self.text_color,
                stroke_color=self.stroke_color, stroke_width=self.stroke_width)

    def _update_slide_video(self):
        self.slide_video_after_id = None
        if self._closing or not self.winfo_exists() or self.slide_capture is None:
            return
        frame = self.slide_capture.latest_frame()
        if frame is not None:
            self.slide_video_photo = frame_to_photo(
                frame, (max(1, self.stage.winfo_width()), max(1, self.stage.winfo_height())), self)
            self.stage.show_photo(self.slide_video_photo, background="black")
            if self.parent.winfo_exists():
                self.parent.update_projector_preview(
                    self.slides[self.index], self.index, len(self.slides), video_frame=frame)
        if self.slide_capture is not None and not self._closing:
            self.slide_video_after_id = self.after(33, self._update_slide_video)

    def _stop_slide_video(self):
        if self.slide_video_after_id:
            try:
                self.after_cancel(self.slide_video_after_id)
            except tk.TclError:
                pass
            self.slide_video_after_id = None
        if self.slide_capture is not None:
            self.slide_capture.close()
            self.slide_capture = None
        self.slide_video_photo = None

    def show_index(self, index):
        if not self.slides:
            return
        self.number_buffer = ""
        self.navigate_to_index(index)

    def current_slide_group(self):
        return next((group for group in self.slide_groups
                     if group["start"] <= self.index < group["end"]), None)

    def queue_number_jump(self, digit):
        if self.number_after_id:
            self.after_cancel(self.number_after_id)
        self.number_buffer = (self.number_buffer + digit).lstrip("0") or "0"
        self.hint.config(text=f"Nomor bagian: {self.number_buffer}  ·  Enter untuk langsung memilih")
        self.number_after_id = self.after(650, self.commit_number_jump)
        return "break"

    def commit_number_jump(self, _event=None):
        if self.number_after_id:
            self.after_cancel(self.number_after_id)
            self.number_after_id = None
        if not self.number_buffer:
            return "break"
        item_number = int(self.number_buffer)
        self.number_buffer = ""
        group = self.current_slide_group()
        target = group["start"] + item_number - 1 if group and item_number > 0 else None
        if target is None or target >= group["end"]:
            available = group["end"] - group["start"] if group else 0
            self.hint.config(text=f"Bagian {item_number} tidak tersedia (item ini memiliki {available} bagian).")
            self.after(1400, self.show_slide)
            return "break"
        self.navigate_to_index(target)
        return "break"

    def next_slide(self, _event=None):
        self.navigate_to_index(min(self.index + 1, len(self.slides) - 1))
        return "break"

    def prev_slide(self, _event=None):
        self.navigate_to_index(max(0, self.index - 1))
        return "break"

    def first_slide(self, _event=None):
        self.navigate_to_index(0)
        return "break"

    def last_slide(self, _event=None):
        self.navigate_to_index(len(self.slides) - 1)
        return "break"

    def toggle_blackout(self, _event=None):
        self._cancel_transition()
        self.blackout = not self.blackout
        self.cleared = False
        self.show_slide()
        return "break"


if __name__ == "__main__":
    SahabatIbadahApp().mainloop()
