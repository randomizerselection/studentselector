import os
import csv
import random
import time
import tkinter as tk
import tkinter.font as tkfont
import ctypes

import ttkbootstrap as ttk
from ttkbootstrap import Style
from ttkbootstrap.dialogs import Messagebox

import sys


def resource_path(relative_path: str) -> str:
    """
    Works in dev (runs next to .py) and in PyInstaller --onefile (runs from _MEIPASS).
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# -------------------------
# Configuration / Assets
# -------------------------

THEME = "cosmo"

WINDOW_WIDTH = 560
TOP_RIGHT_PADDING_X = 20
TOP_PADDING_Y = 0

SLOT_PANEL_W = 900
SLOT_PANEL_H = 660
SLOT_PANEL_MIN_H = 560
SLOT_PANEL_MARGIN_RIGHT = 20
SLOT_PANEL_Y = 80

STUDENTS_CSV = resource_path(r"assets\students.csv")
MESSAGES_CSV = resource_path(r"assets\messages.csv")

INTRO_MUSIC = resource_path(r"assets\welcome.mp3")
CLOSING_MUSIC = resource_path(r"assets\closing.mp3")

SLOT_SOUND_SHORT = resource_path(r"assets\select_student.mp3")
SLOT_SOUND_MEDIUM = resource_path(r"assets\medium_slot.mp3")
SLOT_SOUND_LONG = resource_path(r"assets\long_slot.mp3")
TIMEUP_SOUND = resource_path(r"assets\timeup.mp3")

RATING_SOUNDS = {
    "A*": resource_path(r"assets\sound_a_star.mp3"),
    "A": resource_path(r"assets\sound_a.mp3"),
    "B": resource_path(r"assets\sound_b.mp3"),
    "C": resource_path(r"assets\sound_c.mp3"),
}

SHORT_MAX = 4.99
MEDIUM_MAX = 19.99

TICK_MIN_MS = 60
TICK_MAX_MS = 300

PALETTE = {
    "bg": "#f3ede3",
    "bg_alt": "#ece2d4",
    "panel": "#fffaf2",
    "panel_alt": "#f7ede0",
    "card": "#fffdf8",
    "card_alt": "#f0e6d8",
    "field": "#efe4d5",
    "text_light": "#26333c",
    "text_muted": "#6b7b86",
    "text_dark": "#162028",
    "line": "#d6c8b4",
    "line_soft": "#e4d9cb",
    "accent": "#c9853d",
    "accent_soft": "#9f6930",
    "accent_alt": "#4f9aa4",
    "success": "#7daa73",
    "success_soft": "#dbead6",
    "warning": "#d8a652",
    "warning_soft": "#f3e4c6",
    "danger": "#cf7960",
    "danger_soft": "#f1d8d1",
}


# -------------------------
# Audio Manager (resilient)
# -------------------------

class SoundManager:
    """
    Uses pygame.mixer if available. Missing files or mixer errors degrade silently.
    - music channel: intro/closing/slot loop/rating one-shots (mutually exclusive)
    - sound channel: timeup (attempts pygame.mixer.Sound; falls back if needed)
    - self.enabled flag + set_enabled() to globally mute/unmute audio.
    """
    def __init__(self):
        self._pygame = None
        self._mixer_ok = False
        self.enabled = True

        try:
            import pygame
            self._pygame = pygame
            pygame.mixer.init()
            self._mixer_ok = True
        except Exception:
            self._pygame = None
            self._mixer_ok = False

    def set_enabled(self, enabled: bool):
        """Enable/disable all audio. Disabling stops any currently playing audio."""
        self.enabled = bool(enabled)
        if not self.enabled and self._mixer_ok:
            try:
                self._pygame.mixer.stop()
            except Exception:
                pass

    def _file_exists(self, path: str) -> bool:
        return bool(path) and os.path.isfile(path)

    def stop_music(self):
        if not self._mixer_ok:
            return
        try:
            self._pygame.mixer.music.stop()
        except Exception:
            pass

    def play_music_once(self, path: str):
        """Plays on the music channel once; stops any existing music."""
        if not self.enabled or not self._mixer_ok or not self._file_exists(path):
            return
        try:
            self._pygame.mixer.music.stop()
            self._pygame.mixer.music.load(path)
            self._pygame.mixer.music.play()
        except Exception:
            pass

    def play_music_loop(self, path: str):
        """Plays on the music channel looping; stops any existing music."""
        if not self.enabled or not self._mixer_ok or not self._file_exists(path):
            return
        try:
            self._pygame.mixer.music.stop()
            self._pygame.mixer.music.load(path)
            self._pygame.mixer.music.play(-1)
        except Exception:
            pass

    def play_timeup(self, path: str):
        """Attempts to play as a Sound (can overlap music). Falls back silently."""
        if not self.enabled or not self._mixer_ok or not self._file_exists(path):
            return
        try:
            snd = self._pygame.mixer.Sound(path)
            snd.play()
        except Exception:
            try:
                self._pygame.mixer.music.stop()
                self._pygame.mixer.music.load(path)
                self._pygame.mixer.music.play()
            except Exception:
                pass


# -------------------------
# Data loading
# -------------------------

def load_students_by_class(path: str) -> dict[str, list[str]]:
    """
    Accepts a CSV with 2 columns: class_name, student_name.
    - If there is a header, it will be skipped automatically when detected.
    - UTF-8 is assumed.
    """
    classes: dict[str, list[str]] = {}
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Missing roster file: {path}")

    with open(path, "r", encoding="utf-8", newline="") as f:
        reader = csv.reader(f)
        first = True
        for row in reader:
            if not row or len(row) < 2:
                continue
            if first:
                first = False
                header = ",".join(row).strip().lower()
                if "class" in header and ("student" in header or "name" in header):
                    continue
            class_name = row[0].strip()
            student_name = row[1].strip()
            if not class_name or not student_name:
                continue
            classes.setdefault(class_name, []).append(student_name)

    for k in classes:
        classes[k] = list(dict.fromkeys(classes[k]))  # de-dup preserving order
    return classes


def load_messages_by_rating(path: str) -> dict[str, list[str]]:
    """
    Expects columns 'Rating' and 'Message' (legacy format).
    """
    messages: dict[str, list[str]] = {}
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Missing messages file: {path}")

    with open(path, "r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rating = (row.get("Rating") or "").strip()
            msg = (row.get("Message") or "").strip()
            if not rating or not msg:
                continue
            messages.setdefault(rating, []).append(msg)

    return messages


# -------------------------
# Application
# -------------------------

class InvisibleHandApp:
    def __init__(self, root: tk.Tk, style: Style):
        self.root = root
        self.style = style
        self.sound = SoundManager()
        self.palette = dict(PALETTE)

        # Typography / classroom projection
        self.FONT_FAMILY = self._pick_font_family("Aptos", "Segoe UI", "Arial")
        self.HEADING_FONT_FAMILY = self._pick_font_family("Bahnschrift SemiBold", "Bahnschrift", "Aptos Display", self.FONT_FAMILY)
        self.MONO_FONT_FAMILY = self._pick_font_family("Consolas", "Cascadia Mono", "Courier New", self.FONT_FAMILY)
        self.ui_scale = self._compute_classroom_ui_scale()
        self._apply_classroom_font_defaults()
        self._apply_visual_theme()

        # State
        self.classes = load_students_by_class(STUDENTS_CSV)
        self.messages = load_messages_by_rating(MESSAGES_CSV)

        self.selected_class = ttk.StringVar(value="Select a Class")
        self.time_preset_var = tk.IntVar(value=5)

        self.sound_enabled_var = tk.BooleanVar(value=True)
        self.sound.set_enabled(self.sound_enabled_var.get())
        self.slot_effect_enabled_var = tk.BooleanVar(value=True)

        self.session_students_by_class: dict[str, list[str]] = {}
        self.student_grades_by_class: dict[str, dict[str, str]] = {}
        self.active_class: str | None = None
        self.session_students: list[str] = []
        self.student_grades: dict[str, str] = {}
        self.exit_requested = False

        # Window placement and behavior
        self._configure_root_window()

        # UI
        self._build_main_screen()

    # ---------- Typography helpers ----------

    def _compute_classroom_ui_scale(self) -> float:
        """
        Scale fonts up based on screen height with a projection boost.
        """
        try:
            screen_h = max(1, int(self.root.winfo_screenheight()))
        except Exception:
            screen_h = 1080

        scale = screen_h / 1080.0
        scale = max(1.00, min(scale, 1.55))

        # projection boost
        scale *= 1.15
        return max(1.00, min(scale, 1.75))

    def fs(self, px: int) -> int:
        return max(10, int(round(px * self.ui_scale)))

    def _pick_font_family(self, *candidates: str) -> str:
        fallback = "Segoe UI"
        try:
            available = {name.casefold(): name for name in tkfont.families(self.root)}
        except Exception:
            available = {}

        for name in candidates:
            if not name:
                continue
            if not available or name.casefold() in available:
                return available.get(name.casefold(), name)

        return fallback

    def f(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.FONT_FAMILY, self.fs(px), weight)
        return (self.FONT_FAMILY, self.fs(px))

    def hf(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.HEADING_FONT_FAMILY, self.fs(px), weight)
        return (self.HEADING_FONT_FAMILY, self.fs(px))

    def mf(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.MONO_FONT_FAMILY, self.fs(px), weight)
        return (self.MONO_FONT_FAMILY, self.fs(px))

    def _fit_font_to_width(
        self,
        text: str,
        max_px: int,
        start_px: int,
        min_px: int,
        weight: str = "bold",
        family: str | None = None,
    ) -> tuple:
        """
        Prevent the title from clipping in the narrow dock window by shrinking
        the font until it fits (or hits min).
        """
        family = family or self.HEADING_FONT_FAMILY
        size = self.fs(start_px)
        min_size = self.fs(min_px)

        test_font = tkfont.Font(family=family, size=size, weight=weight)
        while size > min_size and test_font.measure(text) > max_px:
            size -= 1
            test_font.configure(size=size)

        return (family, size, weight)

    def _apply_classroom_font_defaults(self):
        base = self.f(18)
        self.root.option_add("*Font", base)
        self.root.option_add("*Listbox.Font", base)
        self.root.option_add("*TCombobox*Listbox.Font", self.f(20))
        self.root.option_add("*Text.Font", base)

        # TTK defaults
        self.style.configure(".", font=base)
        self.style.configure("TLabel", font=base)
        self.style.configure("secondary.TLabel", font=self.f(16))

        # Labelframe titles
        self.style.configure("TLabelframe.Label", font=self.hf(18, "bold"))
        for bs in ("primary", "secondary", "info", "success", "warning", "danger"):
            self.style.configure(f"{bs}.TLabelframe.Label", font=self.hf(18, "bold"))

        # Inputs
        self.style.configure("TCombobox", font=self.f(20))
        self.style.configure("TSpinbox", font=self.f(20))
        self.style.configure("TEntry", font=self.f(20))

        # Buttons
        self.style.configure("TButton", font=self.hf(16, "bold"))
        for bs in ("primary", "secondary", "success", "info", "warning", "danger"):
            self.style.configure(f"{bs}.TButton", font=self.hf(16, "bold"))

        # Checkbuttons/toggles
        self.style.configure("TCheckbutton", font=self.f(16))

        # Slot window progressbar: thicker for projection
        try:
            self.style.configure("Slot.Horizontal.TProgressbar", thickness=self.fs(12))
            self.style.configure(
                "Slot.Horizontal.TProgressbar",
                troughcolor=self.palette["line_soft"],
                background=self.palette["accent"],
                bordercolor=self.palette["line_soft"],
                lightcolor=self.palette["accent"],
                darkcolor=self.palette["accent"],
            )
        except Exception:
            pass

        # Small scaling nudge for projector clarity
        try:
            self.root.tk.call("tk", "scaling", 1.15)
        except Exception:
            pass

    def _shade(self, c1: str, c2: str, t: float) -> str:
        try:
            c1 = c1.lstrip("#")
            c2 = c2.lstrip("#")
            rgb1 = tuple(int(c1[i:i + 2], 16) for i in (0, 2, 4))
            rgb2 = tuple(int(c2[i:i + 2], 16) for i in (0, 2, 4))
            mixed = tuple(int(a + (b - a) * t) for a, b in zip(rgb1, rgb2))
            return "#{:02x}{:02x}{:02x}".format(*mixed)
        except Exception:
            return c1

    def _format_seconds(self, total_seconds: int | float) -> str:
        total = max(1, int(round(float(total_seconds))))
        minutes, seconds = divmod(total, 60)
        if minutes and seconds:
            return f"{minutes}m {seconds:02d}s"
        if minutes:
            unit = "min" if minutes == 1 else "mins"
            return f"{minutes} {unit}"
        return f"{seconds} sec"

    def _active_or_selected_class(self) -> str | None:
        selected = (self.selected_class.get() or "").strip()
        if selected in self.classes:
            return selected
        if self.active_class in self.classes:
            return self.active_class
        return None

    def _class_metrics(self, class_name: str | None = None) -> dict[str, int | str]:
        class_name = class_name or self._active_or_selected_class()
        roster = list(self.classes.get(class_name, [])) if class_name else []
        session_roster = self.session_students_by_class.get(class_name, roster)
        grades = self.student_grades_by_class.get(class_name, {})
        return {
            "class_name": class_name or "No class selected",
            "roster_total": len(roster),
            "remaining": len(session_roster),
            "graded": len(grades),
        }

    def _rating_meta(self, rating: str) -> dict[str, str]:
        meta = {
            "A*": {"label": "Excellent", "bg": self.palette["success"], "fg": "#1f301d"},
            "A": {"label": "Strong", "bg": self.palette["accent"], "fg": "#2d1c0d"},
            "B": {"label": "Secure", "bg": self.palette["warning"], "fg": "#33220a"},
            "C": {"label": "Needs support", "bg": self.palette["danger"], "fg": "#371814"},
        }
        return meta.get(rating, {"label": "Recorded", "bg": self.palette["panel_alt"], "fg": self.palette["text_light"]})

    def _apply_visual_theme(self):
        p = self.palette
        self.root.configure(bg=p["bg"])

        self.style.configure("App.TFrame", background=p["bg"])
        self.style.configure("Panel.TFrame", background=p["panel"])
        self.style.configure("PanelAlt.TFrame", background=p["panel_alt"])
        self.style.configure("Card.TFrame", background=p["card"])
        self.style.configure("DockHeroTitle.TLabel", background=p["panel"], foreground=p["text_light"], font=self.hf(28, "bold"))
        self.style.configure("DockHeroBody.TLabel", background=p["panel"], foreground=p["text_muted"], font=self.f(15))
        self.style.configure("SectionTitle.TLabel", background=p["card"], foreground=p["text_light"], font=self.hf(15, "bold"))
        self.style.configure("PanelAltTitle.TLabel", background=p["panel_alt"], foreground=p["text_light"], font=self.hf(15, "bold"))
        self.style.configure("PanelTitle.TLabel", background=p["panel"], foreground=p["text_light"], font=self.hf(15, "bold"))
        self.style.configure("Hint.TLabel", background=p["bg"], foreground=p["text_muted"], font=self.f(14))
        self.style.configure("PopupTitle.TLabel", background=p["bg"], foreground=p["text_light"], font=self.hf(34, "bold"))
        self.style.configure("PopupBody.TLabel", background=p["panel"], foreground=p["text_light"], font=self.f(26))
        self.style.configure("SummaryTitle.TLabel", background=p["bg"], foreground=p["text_light"], font=self.hf(26, "bold"))

        # Combobox popdown list uses classic Tk listbox colors, not ttk style colors.
        self.root.option_add("*TCombobox*Listbox.background", p["panel"])
        self.root.option_add("*TCombobox*Listbox.foreground", p["text_light"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", p["accent_alt"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        self.root.option_add("*TCombobox*Listbox.highlightThickness", 0)
        self.root.option_add("*TCombobox*Listbox.borderWidth", 0)
        self.root.option_add("*Listbox.background", p["panel"])
        self.root.option_add("*Listbox.foreground", p["text_light"])
        self.root.option_add("*Listbox.selectBackground", p["accent_alt"])
        self.root.option_add("*Listbox.selectForeground", "#ffffff")
        self.root.option_add("*Listbox.highlightThickness", 0)
        self.root.option_add("*Listbox.borderWidth", 0)

        self.style.configure(
            "PrimaryAction.TButton",
            font=self.hf(17, "bold"),
            foreground="#fffaf3",
            background=p["accent"],
            bordercolor=p["accent"],
            darkcolor=self._shade(p["accent"], "#000000", 0.14),
            lightcolor=self._shade(p["accent"], "#ffffff", 0.08),
            focusthickness=0,
            padding=self.fs(12),
        )
        self.style.map("PrimaryAction.TButton", background=[("active", self._shade(p["accent"], "#ffffff", 0.08))])

        self.style.configure(
            "SecondaryAction.TButton",
            font=self.hf(15, "bold"),
            foreground="#ffffff",
            background=p["accent_alt"],
            bordercolor=p["accent_alt"],
            darkcolor=self._shade(p["accent_alt"], "#000000", 0.16),
            lightcolor=self._shade(p["accent_alt"], "#ffffff", 0.08),
            focusthickness=0,
            padding=self.fs(10),
        )
        self.style.map("SecondaryAction.TButton", background=[("active", self._shade(p["accent_alt"], "#ffffff", 0.08))])

        self.style.configure(
            "Utility.TButton",
            font=self.hf(15, "bold"),
            foreground=p["text_dark"],
            background=p["bg_alt"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["bg_alt"], "#000000", 0.08),
            lightcolor=self._shade(p["bg_alt"], "#ffffff", 0.04),
            focusthickness=0,
            padding=self.fs(9),
        )
        self.style.map("Utility.TButton", background=[("active", self._shade(p["bg_alt"], "#000000", 0.03))])

        self.style.configure(
            "TimeOff.TButton",
            font=self.hf(14, "bold"),
            foreground=p["text_dark"],
            background=p["field"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["field"], "#000000", 0.08),
            lightcolor=self._shade(p["field"], "#ffffff", 0.04),
            focusthickness=0,
            padding=self.fs(8),
        )
        self.style.configure(
            "TimeOn.TButton",
            font=self.hf(14, "bold"),
            foreground="#f7fbff",
            background=p["accent_soft"],
            bordercolor=p["accent"],
            darkcolor=self._shade(p["accent"], "#000000", 0.12),
            lightcolor=self._shade(p["accent"], "#ffffff", 0.05),
            focusthickness=0,
            padding=self.fs(8),
        )
        self.style.map("TimeOff.TButton", background=[("active", self._shade(p["field"], "#ffffff", 0.05))])
        self.style.map("TimeOn.TButton", background=[("active", self._shade(p["accent"], "#ffffff", 0.04))])

        self.style.configure(
            "Dock.TCombobox",
            fieldbackground=p["field"],
            background=p["field"],
            foreground=p["text_dark"],
            bordercolor=p["line"],
            arrowcolor=p["accent_alt"],
            insertcolor=p["text_dark"],
            padding=self.fs(10),
        )
        self.style.map(
            "Dock.TCombobox",
            fieldbackground=[("readonly", p["field"])],
            foreground=[("readonly", p["text_dark"])],
            selectbackground=[("readonly", p["field"])],
            selectforeground=[("readonly", p["text_dark"])],
            bordercolor=[("readonly", p["line"])],
            lightcolor=[("readonly", p["field"])],
            darkcolor=[("readonly", p["field"])],
            arrowcolor=[("readonly", p["accent_alt"])],
        )

        self.style.configure(
            "GradeAStar.TButton",
            font=self.hf(16, "bold"),
            foreground="#122015",
            background=p["success"],
            bordercolor=p["success"],
            darkcolor=self._shade(p["success"], "#000000", 0.10),
            lightcolor=self._shade(p["success"], "#ffffff", 0.05),
            padding=self.fs(10),
        )
        self.style.configure(
            "GradeA.TButton",
            font=self.hf(16, "bold"),
            foreground="#221809",
            background=p["accent"],
            bordercolor=p["accent"],
            darkcolor=self._shade(p["accent"], "#000000", 0.10),
            lightcolor=self._shade(p["accent"], "#ffffff", 0.05),
            padding=self.fs(10),
        )
        self.style.configure(
            "GradeB.TButton",
            font=self.hf(16, "bold"),
            foreground="#241905",
            background=p["warning"],
            bordercolor=p["warning"],
            darkcolor=self._shade(p["warning"], "#000000", 0.10),
            lightcolor=self._shade(p["warning"], "#ffffff", 0.05),
            padding=self.fs(10),
        )
        self.style.configure(
            "GradeC.TButton",
            font=self.hf(16, "bold"),
            foreground="#2c140f",
            background=p["danger"],
            bordercolor=p["danger"],
            darkcolor=self._shade(p["danger"], "#000000", 0.10),
            lightcolor=self._shade(p["danger"], "#ffffff", 0.05),
            padding=self.fs(10),
        )

        self.style.configure(
            "Summary.Vertical.TScrollbar",
            troughcolor=p["line_soft"],
            background=p["bg_alt"],
            arrowcolor=p["text_dark"],
            bordercolor=p["line"],
        )

    # ---------- Window / Layout ----------

    def _desktop_work_area(self) -> tuple[int, int, int, int]:
        """
        Returns the usable desktop area (excluding the taskbar on Windows) as
        left, top, width, height. Falls back to Tk screen size elsewhere.
        """
        try:
            if os.name == "nt":
                class RECT(ctypes.Structure):
                    _fields_ = [
                        ("left", ctypes.c_long),
                        ("top", ctypes.c_long),
                        ("right", ctypes.c_long),
                        ("bottom", ctypes.c_long),
                    ]

                rect = RECT()
                ok = ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
                if ok:
                    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        return 0, 0, screen_w, screen_h

    def _root_window_size(self) -> tuple[int, int]:
        _, _, work_w, work_h = self._desktop_work_area()
        window_w = min(WINDOW_WIDTH, max(440, work_w - 24))
        window_h = max(540, work_h - max(12, TOP_PADDING_Y + 12))
        return window_w, window_h

    def _configure_root_window(self):
        self.root.title("Random Student Selector")
        self.root.attributes("-topmost", True)

        work_left, work_top, work_w, work_h = self._desktop_work_area()
        window_w, window_h = self._root_window_size()

        x = work_left + max(0, work_w - window_w - TOP_RIGHT_PADDING_X)
        y = work_top + TOP_PADDING_Y
        self.root.geometry(f"{window_w}x{window_h}+{x}+{y}")
        self.root.minsize(min(WINDOW_WIDTH, window_w), min(max(520, work_h - 80), window_h))
        self.root.maxsize(work_w, work_h)

    def _slot_window_height(self) -> int:
        _, work_top, _, work_h = self._desktop_work_area()
        available_h = max(480, work_h - max(SLOT_PANEL_Y - work_top, 0) - 24)
        return min(SLOT_PANEL_H, max(SLOT_PANEL_MIN_H, available_h))

    def _slot_window_size(self) -> tuple[int, int]:
        work_left, work_top, work_w, work_h = self._desktop_work_area()
        available_h = max(480, work_h - max(SLOT_PANEL_Y - work_top, 0) - 24)
        window_h = min(SLOT_PANEL_H, max(SLOT_PANEL_MIN_H, available_h))
        window_w = min(SLOT_PANEL_W, max(720, work_w - 40))
        return window_w, window_h

    def _slot_window_geometry(self, parent_w: int) -> str:
        work_left, work_top, work_w, work_h = self._desktop_work_area()
        window_w, window_h = self._slot_window_size()

        x = work_left + max(0, work_w - window_w - SLOT_PANEL_MARGIN_RIGHT)
        y = work_top + min(SLOT_PANEL_Y, max(0, work_h - window_h - 12))

        root_x = self.root.winfo_x() if self.root.winfo_ismapped() else x
        root_w = self.root.winfo_width() if self.root.winfo_ismapped() else WINDOW_WIDTH
        control_left = root_x
        if x < control_left + root_w:
            x = max(work_left, control_left - window_w - 15)

        return f"{window_w}x{window_h}+{x}+{y}"

    def _clear_root(self):
        for w in self.root.winfo_children():
            w.destroy()

    # ---------- Sound toggle ----------

    def _on_sound_toggle(self):
        self.sound.set_enabled(self.sound_enabled_var.get())

    def _on_slot_effect_toggle(self):
        # Intentionally minimal; toggle is read when the slot window starts.
        _ = self.slot_effect_enabled_var.get()

    # ---------- Main Screen ----------

    def _build_main_screen(self):
        self._clear_root()
        p = self.palette
        selected_class = self._active_or_selected_class()
        metrics = self._class_metrics(selected_class)
        has_class = selected_class in self.classes
        _, window_h = self._root_window_size()

        if window_h <= 900:
            dock_density = 0.64
        elif window_h <= 1080:
            dock_density = 0.74
        elif window_h <= 1280:
            dock_density = 0.84
        else:
            dock_density = 0.92

        def d(px: int, floor: int = 4) -> int:
            return max(floor, int(round(self.fs(px) * dock_density)))

        compact = dock_density < 1.0
        outer_pad = d(8)
        panel_pad = d(10)
        panel_gap = d(5)
        card_gap = d(4)
        button_gap = d(4)

        content = tk.Frame(self.root, bg=p["bg"], padx=outer_pad, pady=outer_pad)
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(0, weight=1)

        header = tk.Frame(content, bg=p["bg"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, panel_gap))
        tk.Label(
            header,
            text="Random Student Selector",
            font=self.hf(26 if compact else 28, "bold"),
            bg=p["bg"],
            fg=p["text_light"],
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            header,
            text=(
                f"{metrics['remaining']} left | {metrics['graded']} graded"
                if has_class else
                "Select a class to start"
            ),
            font=self.f(11 if compact else 12),
            bg=p["bg"],
            fg=p["text_muted"],
            anchor="w",
        ).pack(anchor="w", pady=(d(2), 0))

        panel = tk.Frame(content, bg=p["panel"], padx=panel_pad, pady=panel_pad)
        panel.grid(row=1, column=0, sticky="ew")
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(panel, text="Class", font=self.hf(13, "bold"), bg=p["panel"], fg=p["text_light"], anchor="w").grid(row=0, column=0, sticky="w")
        class_dropdown = ttk.Combobox(
            panel,
            textvariable=self.selected_class,
            state="readonly",
            font=self.f(20),
            style="Dock.TCombobox",
            values=["Select a Class"] + sorted(self.classes.keys()),
        )
        class_dropdown.grid(row=1, column=0, sticky="ew", pady=(d(4), d(4)))
        class_dropdown.bind("<<ComboboxSelected>>", lambda e: self._on_class_selected())

        roster_note = (
            f"{metrics['remaining']} left of {metrics['roster_total']}"
            if has_class else
            "No class selected"
        )
        tk.Label(panel, text=roster_note, font=self.f(11 if compact else 12), bg=p["panel"], fg=p["text_muted"], anchor="w").grid(row=2, column=0, sticky="w", pady=(0, d(8)))

        tk.Label(panel, text="Timer", font=self.hf(13, "bold"), bg=p["panel"], fg=p["text_light"], anchor="w").grid(row=3, column=0, sticky="w")
        tf = tk.Frame(panel, bg=p["panel"])
        tf.grid(row=4, column=0, sticky="ew", pady=(d(4), d(8)))
        for col in range(4):
            tf.grid_columnconfigure(col, weight=1, uniform="preset")

        presets = [("5 sec", 5), ("30 sec", 30), ("1 min", 60), ("2 min", 120)]

        def _set_preset(seconds: int):
            self.time_preset_var.set(seconds)
            self._build_main_screen()

        for i, (label, seconds) in enumerate(presets):
            btn = ttk.Button(
                tf,
                text=label,
                style="TimeOn.TButton" if seconds == self.time_preset_var.get() else "TimeOff.TButton",
                command=lambda s=seconds: _set_preset(s),
            )
            btn.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else button_gap, 0), ipady=d(4))

        toggles = tk.Frame(panel, bg=p["panel"])
        toggles.grid(row=5, column=0, sticky="ew", pady=(0, d(8)))
        toggles.grid_columnconfigure(0, weight=1, uniform="toggle")
        toggles.grid_columnconfigure(1, weight=1, uniform="toggle")

        def toggle_tile(col: int, title: str, variable, command, bootstyle: str):
            tile = tk.Frame(
                toggles,
                bg=p["bg_alt"],
                padx=d(8),
                pady=d(6),
            )
            tile.grid(row=0, column=col, sticky="ew", padx=(0, card_gap) if col == 0 else (card_gap, 0))
            text = tk.Frame(tile, bg=p["bg_alt"])
            text.pack(side="left", fill="both", expand=True)
            tk.Label(text, text=title, font=self.hf(12 if compact else 13, "bold"), bg=p["bg_alt"], fg=p["text_light"], anchor="w").pack(anchor="w")
            try:
                ttk.Checkbutton(tile, text="", variable=variable, command=command, bootstyle=bootstyle).pack(side="right", padx=(d(6), 0))
            except Exception:
                ttk.Checkbutton(tile, text="", variable=variable, command=command, bootstyle="round-toggle").pack(side="right", padx=(d(6), 0))

        toggle_tile(0, "Sound", self.sound_enabled_var, self._on_sound_toggle, "success-round-toggle")
        toggle_tile(1, "Slot Effect", self.slot_effect_enabled_var, self._on_slot_effect_toggle, "warning-round-toggle")

        controls = tk.Frame(panel, bg=p["panel"])
        controls.grid(row=6, column=0, sticky="ew", pady=(0, d(8)))
        for col in range(3):
            controls.grid_columnconfigure(col, weight=1, uniform="ctrl")
        ttk.Button(controls, text="Play Intro", style="Utility.TButton", command=self._play_intro).grid(
            row=0, column=0, sticky="ew", padx=(0, card_gap), ipady=d(4)
        )
        ttk.Button(controls, text="View Grades", style="SecondaryAction.TButton", command=self._show_grades_summary).grid(
            row=0, column=1, sticky="ew", padx=card_gap, ipady=d(4)
        )
        ttk.Button(controls, text="Play Closing", style="Utility.TButton", command=self._play_closing).grid(
            row=0, column=2, sticky="ew", padx=(card_gap, 0), ipady=d(4)
        )

        ttk.Button(
            panel,
            text="START SELECTION",
            style="PrimaryAction.TButton",
            command=self._start_session,
        ).grid(row=7, column=0, sticky="ew", ipady=d(6))

    def _on_class_selected(self):
        self.exit_requested = False
        self._build_main_screen()

    # ---------- Audio controls ----------

    def _play_intro(self):
        self.exit_requested = False
        self.sound.play_music_once(INTRO_MUSIC)

    def _play_closing(self):
        self.exit_requested = False
        self.sound.play_music_once(CLOSING_MUSIC)

    # ---------- Session logic ----------

    def _get_duration_seconds(self) -> float:
        try:
            total = int(self.time_preset_var.get())
        except Exception:
            total = 30
        total = max(1, total)
        return float(total)

    def _slot_sound_for_duration(self, duration: float) -> str:
        if duration <= SHORT_MAX:
            return SLOT_SOUND_SHORT
        if duration <= MEDIUM_MAX:
            return SLOT_SOUND_MEDIUM
        return SLOT_SOUND_LONG

    def _start_session(self):
        self.exit_requested = False

        class_name = (self.selected_class.get() or "").strip()
        if class_name == "Select a Class" or class_name not in self.classes:
            Messagebox.show_error(title="Error", message="Please select a valid class.")
            return

        if class_name not in self.session_students_by_class:
            roster = list(self.classes[class_name])
            if not roster:
                Messagebox.show_error(title="Error", message=f"No students found for {class_name}.")
                return
            self.session_students_by_class[class_name] = roster

        if class_name not in self.student_grades_by_class:
            self.student_grades_by_class[class_name] = {}

        self.active_class = class_name
        self.session_students = self.session_students_by_class[class_name]
        self.student_grades = self.student_grades_by_class[class_name]

        if not self.session_students:
            self._show_grades_summary()
            return

        self._next_student(class_name)

    def _next_student(self, class_name: str):
        if self.exit_requested:
            self._build_main_screen()
            return

        if not self.session_students:
            self._show_grades_summary()
            return

        final_student = random.choice(self.session_students)
        self._show_slot_window(class_name, final_student)

    # ---------- Slot window ----------

    def _show_slot_window(self, class_name: str, final_student: str):
        duration = self._get_duration_seconds()
        slot_sound = self._slot_sound_for_duration(duration)
        slot_effect_enabled = bool(self.slot_effect_enabled_var.get())
        window_w, window_h = self._slot_window_size()

        win = ttk.Toplevel(self.root)
        win.title(f"{class_name} - Select Student")
        win.attributes("-topmost", True)
        compact = window_h < 620 or window_w < 820
        win.geometry(self._slot_window_geometry(parent_w=WINDOW_WIDTH))
        _, _, work_w, work_h = self._desktop_work_area()
        win.minsize(min(window_w, 720), min(window_h, 500))
        win.maxsize(work_w, work_h)
        win.configure(bg=self.palette["bg"])

        p = self.palette
        bg = p["bg"]
        panel = p["panel"]
        panel_alt = p["panel_alt"]
        fg = p["text_light"]
        secondary = p["text_muted"]
        primary = p["accent"]
        success = p["success"]
        focus_fill = self._shade(primary, panel, 0.82)

        self.style.configure("SlotBg.TFrame", background=bg)
        self.style.configure("SlotPanel.TFrame", background=panel)
        self.style.configure("SlotInner.TFrame", background=panel_alt)

        alive = True

        def on_close():
            nonlocal alive
            alive = False
            self.exit_requested = True
            try:
                self.sound.stop_music()
            except Exception:
                pass
            try:
                win.destroy()
            except Exception:
                pass
            self._build_main_screen()

        def on_escape(event=None):
            on_close()

        win.bind("<Escape>", on_escape)
        win.protocol("WM_DELETE_WINDOW", on_close)

        main = tk.Frame(win, bg=bg, padx=self.fs(16), pady=self.fs(16))
        main.pack(fill="both", expand=True)

        topbar = tk.Frame(main, bg=panel_alt, padx=self.fs(12), pady=self.fs(10))
        topbar.pack(fill="x", pady=(0, self.fs(10)))
        remaining_value = tk.Label(topbar, text=f"{len(self.session_students)} left", font=self.hf(13, "bold"), bg=panel_alt, fg=p["accent_alt"])
        remaining_value.pack(side="left")
        timer_value = tk.Label(topbar, text=self._format_seconds(duration), font=self.hf(13, "bold"), bg=panel_alt, fg=primary)
        timer_value.pack(side="right")

        stage_body = tk.Frame(main, bg=panel, padx=self.fs(16), pady=self.fs(16))
        stage_body.pack(fill="both", expand=True)

        reel_wrap = tk.Frame(stage_body, bg=panel)
        reel_wrap.pack(fill="x")

        canvas_w = max(560, window_w - self.fs(100))
        reserved_h = self.fs(210 if compact else 230)
        usable_h = max(self.fs(170), window_h - reserved_h)
        row_min = self.fs(42 if compact else 48)
        row_max = self.fs(56 if compact else 62)
        row_h = max(row_min, min(row_max, int((usable_h - self.fs(54)) / 3)))
        canvas_h = row_h * 3 + self.fs(54)
        cx = canvas_w // 2
        cy = canvas_h // 2
        reel = tk.Canvas(
            reel_wrap,
            width=canvas_w,
            height=canvas_h,
            bg=panel,
            highlightthickness=0,
            bd=0,
        )
        reel.pack(pady=(0, self.fs(12)))

        pad = self.fs(10)

        def _round_rect(x1, y1, x2, y2, r, **kwargs):
            r = max(2, int(r))
            points = [
                x1 + r, y1,
                x2 - r, y1,
                x2, y1,
                x2, y1 + r,
                x2, y2 - r,
                x2, y2,
                x2 - r, y2,
                x1 + r, y2,
                x1, y2,
                x1, y2 - r,
                x1, y1 + r,
                x1, y1,
            ]
            return reel.create_polygon(points, smooth=True, splinesteps=20, **kwargs)

        outer_stage = _round_rect(
            pad,
            pad,
            canvas_w - pad,
            canvas_h - pad,
            self.fs(18),
            fill=self._shade(panel_alt, "#ffffff", 0.22),
            outline="",
            width=0,
        )

        band_h = int(row_h * 0.92)
        band = _round_rect(
            pad + self.fs(28),
            cy - band_h // 2,
            canvas_w - (pad + self.fs(28)),
            cy + band_h // 2,
            self.fs(12),
            fill=focus_fill,
            outline="",
            width=0,
        )

        def _is_cjk(ch: str) -> bool:
            o = ord(ch)
            return (
                0x4E00 <= o <= 0x9FFF or
                0x3400 <= o <= 0x4DBF or
                0x3040 <= o <= 0x30FF or
                0xAC00 <= o <= 0xD7AF
            )

        def _format_name(name: str) -> str:
            s = (name or "").strip()
            if not s:
                return s
            if " " in s:
                has_cjk = any(_is_cjk(ch) for ch in s)
                has_latin = any(("A" <= ch <= "Z") or ("a" <= ch <= "z") for ch in s)
                if has_cjk and has_latin:
                    left, right = s.split(None, 1)
                    return f"{left} · {right}"
            return s

        font_cache: dict[tuple[str, int, int, str], tuple] = {}
        max_text_w = canvas_w - (pad * 2) - self.fs(110)

        def _fit_font(text: str, base_px: int, min_px: int, weight="bold") -> tuple:
            key = (text, base_px, min_px, weight)
            if key in font_cache:
                return font_cache[key]

            size = self.fs(base_px)
            min_size = self.fs(min_px)
            fnt = tkfont.Font(family=self.HEADING_FONT_FAMILY, size=size, weight=weight)

            while size > min_size and fnt.measure(text) > max_text_w:
                size -= 1
                fnt.configure(size=size)

            font_cache[key] = (self.HEADING_FONT_FAMILY, size, weight)
            return font_cache[key]

        pool = self.session_students[:] if self.session_students else [final_student]
        prev_raw = random.choice(pool)
        cur_raw = random.choice(pool)
        next_raw = random.choice(pool)

        dim = secondary
        small_px = 19
        big_px = 32
        min_small_px = 15
        min_big_px = 22

        t_prev = reel.create_text(cx, cy - row_h, text=_format_name(prev_raw), fill=dim, anchor="center", justify="center")
        t_cur = reel.create_text(cx, cy, text=_format_name(cur_raw), fill=fg, anchor="center", justify="center")
        t_next = reel.create_text(cx, cy + row_h, text=_format_name(next_raw), fill=dim, anchor="center", justify="center")

        def _style_item(item, y_pos: float):
            d = abs(y_pos - cy) / max(1, row_h)
            d = min(1.0, d)
            scale = (1.0 - d) ** 2
            size_px = int(round(small_px + (big_px - small_px) * scale))
            min_px = int(round(min_small_px + (min_big_px - min_small_px) * scale))
            text = reel.itemcget(item, "text")
            reel.itemconfig(
                item,
                font=_fit_font(text, size_px, min_px),
                fill=fg if scale >= 0.92 else dim,
            )

        progress = ttk.Progressbar(
            stage_body,
            mode="determinate",
            maximum=100,
            bootstyle="primary",
            style="Slot.Horizontal.TProgressbar",
        )
        progress.pack(fill="x", pady=(self.fs(8), 0))

        buttons = ttk.Frame(main, style="SlotBg.TFrame", padding=(0, self.fs(12), 0, 0))
        buttons.pack(fill="x")

        if slot_effect_enabled:
            self.sound.play_music_loop(slot_sound)

        def _timeup_safe():
            if alive and win.winfo_exists():
                self.sound.play_timeup(TIMEUP_SOUND)

        win.after(max(0, int(duration * 1000) - 200), _timeup_safe)

        start_time = time.time()
        last_time = start_time
        phase = 0.0

        max_rows_per_sec = 7.5
        min_rows_per_sec = 1.0
        cur_rows_per_sec = max_rows_per_sec
        speed_smooth_tau = 0.18

        final_mode = False
        final_roll_time = 0.65
        final_start_time = None
        final_start_phase = 0.0

        def _update_timer_label(elapsed: float):
            remaining = max(0, duration - elapsed)
            timer_value.config(text=self._format_seconds(max(1, remaining)) if remaining > 0 else "Locking")

        def _rotate_once():
            nonlocal prev_raw, cur_raw, next_raw
            prev_raw, cur_raw = cur_raw, next_raw
            next_raw = random.choice(pool)

            reel.itemconfig(t_prev, text=_format_name(prev_raw))
            reel.itemconfig(t_cur, text=_format_name(cur_raw))
            reel.itemconfig(t_next, text=_format_name(next_raw))

        def _render(phase_px: float):
            y_prev = (cy - row_h) - phase_px
            y_cur = cy - phase_px
            y_next = (cy + row_h) - phase_px

            reel.coords(t_prev, cx, y_prev)
            reel.coords(t_cur, cx, y_cur)
            reel.coords(t_next, cx, y_next)

            _style_item(t_prev, y_prev)
            _style_item(t_cur, y_cur)
            _style_item(t_next, y_next)

        def _finalize(center_item=None):
            nonlocal cur_raw, t_prev, t_cur, t_next

            if center_item is t_next:
                t_cur, t_next = t_next, t_cur
            elif center_item is t_prev:
                t_cur, t_prev = t_prev, t_cur

            cur_raw = final_student
            reel.itemconfig(t_cur, text=_format_name(cur_raw))
            reel.itemconfig(t_cur, font=_fit_font(reel.itemcget(t_cur, "text"), big_px, min_big_px))
            reel.itemconfig(band, fill=self._shade(success, panel, 0.68), outline="")
            reel.itemconfig(t_cur, fill=success)
            reel.itemconfig(t_prev, state="hidden")
            reel.itemconfig(t_next, state="hidden")

            timer_value.config(text="Complete", fg=success)

            try:
                progress.pack_forget()
            except Exception:
                pass

            self.sound.stop_music()

            if final_student in self.session_students:
                self.session_students.remove(final_student)
            remaining_value.config(text=f"{len(self.session_students)} left")

            self._render_grading_controls(win, buttons, class_name, final_student)

        def _frame():
            nonlocal last_time, phase, final_mode, final_start_time, final_start_phase, next_raw, cur_rows_per_sec

            if (not alive) or (not win.winfo_exists()):
                return

            now = time.time()
            dt = max(0.001, now - last_time)
            last_time = now
            elapsed = now - start_time
            _update_timer_label(elapsed)

            if (not final_mode) and (elapsed >= duration):
                final_mode = True
                final_start_time = now
                final_start_phase = phase
                next_raw = final_student
                reel.itemconfig(t_next, text=_format_name(next_raw))

            if not final_mode:
                t = max(0.0, min(1.0, elapsed / max(0.001, duration)))
                target_rows = min_rows_per_sec + (max_rows_per_sec - min_rows_per_sec) * ((1.0 - t) ** 2.1)
                alpha = min(1.0, dt / max(0.001, speed_smooth_tau))
                cur_rows_per_sec += (target_rows - cur_rows_per_sec) * alpha
                phase += (cur_rows_per_sec * row_h) * dt

                while phase >= row_h - 1e-6:
                    phase -= row_h
                    _rotate_once()

                _render(phase)
            else:
                if final_start_time is None:
                    final_start_time = now
                    final_start_phase = phase

                progress_t = min(1.0, (now - final_start_time) / max(0.001, final_roll_time))
                eased = 1.0 - (1.0 - progress_t) ** 3
                phase = final_start_phase + (row_h - final_start_phase) * eased
                _render(phase)

                if progress_t >= 1.0:
                    _finalize(center_item=t_next)
                    return

            progress["value"] = min(100, int(min(elapsed, duration) / max(0.001, duration) * 100))
            win.after(16, _frame)

        def _frame_no_effect():
            if (not alive) or (not win.winfo_exists()):
                return

            now = time.time()
            elapsed = now - start_time
            _update_timer_label(elapsed)
            progress["value"] = min(100, int(min(elapsed, duration) / max(0.001, duration) * 100))

            if elapsed >= duration:
                _finalize(center_item=t_cur)
                return

            win.after(50, _frame_no_effect)

        _render(0.0)
        if slot_effect_enabled:
            _frame()
        else:
            reel.itemconfig(t_prev, state="hidden")
            reel.itemconfig(t_next, state="hidden")
            reel.itemconfig(t_cur, text="Get ready", fill=fg)
            reel.itemconfig(t_cur, font=_fit_font("Get ready", small_px, min_small_px))
            _frame_no_effect()

    # ---------- Grading controls ----------

    def _render_grading_controls(self, win, button_frame, class_name: str, student_name: str):
        for w in button_frame.winfo_children():
            w.destroy()

        p = self.palette
        button_frame.configure(style="SlotBg.TFrame", padding=(0, self.fs(8), 0, 0))
        card = tk.Frame(button_frame, bg=p["bg"])
        card.pack(fill="x")
        for col in range(4):
            card.columnconfigure(col, weight=1, uniform="rate")

        window_h = self._slot_window_height()
        compact = window_h < 620
        button_ipady = self.fs(10) if compact else self.fs(14)

        def apply_rating(rating: str):
            self.student_grades[student_name] = rating
            self.sound.play_music_once(RATING_SOUNDS.get(rating, ""))

            msg_list = self.messages.get(rating) or []
            msg = random.choice(msg_list) if msg_list else "Noted."

            try:
                win.destroy()
            except Exception:
                pass

            self._show_message_popup(title="Feedback", message=msg, class_name=class_name)

        rating_styles = {"A*": "GradeAStar.TButton", "A": "GradeA.TButton", "B": "GradeB.TButton", "C": "GradeC.TButton"}

        for i, rating in enumerate(("A*", "A", "B", "C")):
            meta = self._rating_meta(rating)
            ttk.Button(
                card,
                text=f"{rating}\n{meta['label']}",
                style=rating_styles[rating],
                command=lambda r=rating: apply_rating(r),
            ).grid(
                row=0,
                column=i,
                sticky="ew",
                padx=(0 if i == 0 else self.fs(10), 0),
                ipady=button_ipady,
            )

    # ---------- Feedback popup ----------

    def _show_message_popup(self, title: str, message: str, class_name: str):
        msg_win = ttk.Toplevel(self.root)
        msg_win.title(f"{class_name} - {title}")
        msg_win.attributes("-topmost", True)
        msg_win.geometry(self._slot_window_geometry(parent_w=WINDOW_WIDTH))
        msg_win.minsize(760, self._slot_window_height())
        msg_win.configure(bg=self.palette["bg"])
        p = self.palette
        metrics = self._class_metrics(class_name)

        def exit_to_main():
            self.exit_requested = True
            try:
                msg_win.destroy()
            except Exception:
                pass
            self._build_main_screen()

        def on_escape(event=None):
            exit_to_main()

        msg_win.bind("<Escape>", on_escape)
        msg_win.protocol("WM_DELETE_WINDOW", exit_to_main)

        outer = tk.Frame(msg_win, bg=p["bg"], padx=self.fs(24), pady=self.fs(24))
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        tk.Label(
            outer,
            text=f"{metrics['remaining']} left | {metrics['graded']} graded",
            font=self.f(12),
            bg=p["bg"],
            fg=p["text_muted"],
            anchor="center",
        ).grid(row=0, column=0, sticky="ew", pady=(0, self.fs(12)))

        body_card = tk.Frame(
            outer,
            bg=p["panel"],
            padx=self.fs(28),
            pady=self.fs(28),
        )
        body_card.grid(row=1, column=0, sticky="nsew", pady=(0, self.fs(16)))
        body_card.grid_rowconfigure(0, weight=1)
        body_card.grid_columnconfigure(0, weight=1)

        quote = tk.Frame(body_card, bg=p["panel"])
        quote.grid(row=0, column=0, sticky="nsew")
        tk.Label(quote, text="“", font=self.hf(48, "bold"), bg=p["panel"], fg=p["accent"]).pack(anchor="center")
        tk.Label(
            quote,
            text=message,
            font=self.f(24),
            bg=p["panel"],
            fg=p["text_light"],
            wraplength=SLOT_PANEL_W - 220,
            justify="center",
        ).pack(fill="both", expand=True, pady=(0, self.fs(10)))

        try:
            quote.destroy()
        except Exception:
            pass

        tk.Label(
            body_card,
            text=message,
            font=self.f(24),
            bg=p["panel"],
            fg=p["text_light"],
            wraplength=SLOT_PANEL_W - 220,
            justify="center",
        ).grid(row=0, column=0, sticky="nsew")

        btns = tk.Frame(outer, bg=p["bg"])
        btns.grid(row=2, column=0, sticky="ew")
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        def next_student():
            try:
                msg_win.destroy()
            except Exception:
                pass
            self._next_student(class_name)

        ttk.Button(
            btns,
            text="Next Student",
            style="PrimaryAction.TButton",
            command=next_student
        ).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(10)), ipady=self.fs(12))

        ttk.Button(
            btns,
            text="Return To Dock",
            style="Utility.TButton",
            command=exit_to_main
        ).grid(row=0, column=1, sticky="ew", padx=(self.fs(10), 0), ipady=self.fs(12))

    # ---------- Grades summary ----------

    def _show_grades_summary(self):
        self.exit_requested = False

        win = ttk.Toplevel(self.root)
        win.title("Grades Summary")
        win.attributes("-topmost", True)
        win.geometry("860x760+30+80")
        win.minsize(760, 680)
        win.configure(bg=self.palette["bg"])
        p = self.palette

        class_name = self.active_class or self._active_or_selected_class() or "Current Session"
        if self.active_class and self.active_class in self.student_grades_by_class:
            grades = self.student_grades_by_class.get(self.active_class, {})
            remaining = len(self.session_students_by_class.get(self.active_class, []))
            roster_total = len(self.classes.get(self.active_class, []))
        elif class_name in self.student_grades_by_class:
            grades = self.student_grades_by_class.get(class_name, {})
            remaining = len(self.session_students_by_class.get(class_name, self.classes.get(class_name, [])))
            roster_total = len(self.classes.get(class_name, []))
        else:
            grades = self.student_grades
            remaining = len(self.session_students)
            roster_total = len(grades) + remaining

        def on_escape(event=None):
            try:
                win.destroy()
            except Exception:
                pass

        win.bind("<Escape>", on_escape)

        counts = {rating: 0 for rating in ("A*", "A", "B", "C")}
        for rating in grades.values():
            counts[rating] = counts.get(rating, 0) + 1

        outer = tk.Frame(win, bg=p["bg"], padx=self.fs(24), pady=self.fs(24))
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        header = tk.Frame(outer, bg=p["bg"])
        header.grid(row=0, column=0, sticky="ew", pady=(0, self.fs(16)))
        header.grid_columnconfigure(0, weight=1)
        tk.Label(header, text=class_name, font=self.hf(24, "bold"), bg=p["bg"], fg=p["text_light"]).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text=f"{len(grades)} graded | {remaining} left" if roster_total else "No class session yet",
            font=self.f(13),
            bg=p["bg"],
            fg=p["text_muted"],
        ).grid(row=1, column=0, sticky="w", pady=(self.fs(4), 0))

        chip_row = tk.Frame(header, bg=p["bg"])
        chip_row.grid(row=0, column=1, rowspan=2, sticky="e", padx=(self.fs(18), 0))
        for idx, rating in enumerate(("A*", "A", "B", "C")):
            meta = self._rating_meta(rating)
            chip = tk.Frame(chip_row, bg=p["panel_alt"], padx=self.fs(12), pady=self.fs(10))
            chip.grid(row=0, column=idx, padx=(0 if idx == 0 else self.fs(8), 0))
            tk.Label(chip, text=rating, font=self.hf(16, "bold"), bg=p["panel_alt"], fg=meta["bg"]).pack(anchor="center")
            tk.Label(chip, text=str(counts.get(rating, 0)), font=self.f(13, "bold"), bg=p["panel_alt"], fg=p["text_light"]).pack(anchor="center", pady=(self.fs(4), 0))

        body = tk.Frame(outer, bg=p["panel"], padx=self.fs(18), pady=self.fs(18))
        body.grid(row=1, column=0, sticky="nsew", pady=(0, self.fs(16)))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        list_shell = tk.Frame(body, bg=p["bg_alt"])
        list_shell.grid(row=0, column=0, sticky="nsew")
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        if not grades:
            empty = tk.Frame(list_shell, bg=p["bg_alt"], padx=self.fs(24), pady=self.fs(24))
            empty.grid(row=0, column=0, sticky="nsew")
            tk.Label(empty, text="No ratings recorded yet.", font=self.hf(20, "bold"), bg=p["bg_alt"], fg=p["text_light"]).pack(anchor="center", pady=(self.fs(24), 0))
        else:
            canvas = tk.Canvas(list_shell, bg=p["bg_alt"], highlightthickness=0, bd=0)
            canvas.grid(row=0, column=0, sticky="nsew")
            scroll = ttk.Scrollbar(list_shell, orient="vertical", command=canvas.yview, style="Summary.Vertical.TScrollbar")
            scroll.grid(row=0, column=1, sticky="ns")
            canvas.configure(yscrollcommand=scroll.set)

            rows = tk.Frame(canvas, bg=p["bg_alt"])
            window_id = canvas.create_window((0, 0), window=rows, anchor="nw")

            def _sync_rows(_event=None):
                canvas.configure(scrollregion=canvas.bbox("all"))

            def _stretch_rows(event):
                canvas.itemconfigure(window_id, width=event.width)

            rows.bind("<Configure>", _sync_rows)
            canvas.bind("<Configure>", _stretch_rows)

            for idx, (student, rating) in enumerate(grades.items(), start=1):
                meta = self._rating_meta(rating)
                row_bg = p["panel"] if idx % 2 else p["panel_alt"]
                row = tk.Frame(rows, bg=row_bg, padx=self.fs(14), pady=self.fs(12))
                row.pack(fill="x", expand=True)
                row.grid_columnconfigure(1, weight=1)

                tk.Label(row, text=f"{idx:02d}", font=self.mf(12, "bold"), bg=row_bg, fg=p["text_muted"], width=4, anchor="w").grid(row=0, column=0, sticky="w")
                tk.Label(row, text=student, font=self.hf(16, "bold"), bg=row_bg, fg=p["text_light"], anchor="w").grid(row=0, column=1, sticky="w")

                pill = tk.Label(
                    row,
                    text=f"{rating}  {meta['label']}",
                    font=self.hf(12, "bold"),
                    bg=meta["bg"],
                    fg=meta["fg"],
                    padx=self.fs(10),
                    pady=self.fs(5),
                )
                pill.grid(row=0, column=2, sticky="e")

        btns = tk.Frame(outer, bg=p["bg"])
        btns.grid(row=2, column=0, sticky="ew")
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        active_ready = bool(self.active_class and self.session_students_by_class.get(self.active_class))

        if active_ready:
            ttk.Button(
                btns,
                text="Resume Session",
                style="SecondaryAction.TButton",
                command=lambda: (win.destroy(), self._next_student(self.active_class)),
            ).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(10)), ipady=self.fs(10))
            ttk.Button(
                btns,
                text="Close",
                style="Utility.TButton",
                command=win.destroy,
            ).grid(row=0, column=1, sticky="ew", padx=(self.fs(10), 0), ipady=self.fs(10))
        else:
            ttk.Button(
                btns,
                text="Close",
                style="Utility.TButton",
                command=win.destroy,
            ).grid(row=0, column=0, columnspan=2, sticky="ew", ipady=self.fs(10))


def main():
    style = Style(theme=THEME)
    root = style.master
    _ = InvisibleHandApp(root, style)
    root.mainloop()


if __name__ == "__main__":
    main()

