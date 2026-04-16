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


def enable_windows_high_dpi() -> None:
    """
    Prevent Windows from bitmap-scaling the whole app on high-DPI displays.
    """
    if os.name != "nt":
        return

    try:
        user32 = ctypes.windll.user32
        shcore = getattr(ctypes.windll, "shcore", None)

        set_context = getattr(user32, "SetProcessDpiAwarenessContext", None)
        if set_context:
            for context in (ctypes.c_void_p(-4), ctypes.c_void_p(-3)):
                try:
                    if set_context(context):
                        return
                except Exception:
                    pass

        if shcore:
            set_awareness = getattr(shcore, "SetProcessDpiAwareness", None)
            if set_awareness:
                try:
                    if set_awareness(2) == 0:
                        return
                except Exception:
                    pass

        set_aware = getattr(user32, "SetProcessDPIAware", None)
        if set_aware:
            try:
                set_aware()
            except Exception:
                pass
    except Exception:
        pass


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

FEEDBACK_LABELS = [
    "Deep Thinking",
    "Clear Explanation",
    "Well Supported",
    "Creative Idea",
    "Independent",
    "Confident",
    "Improving",
    "Good Use of Vocabulary",
    "Accurate Detail",
    "Strong Participation",
]

FEEDBACK_LABEL_PHRASES = {
    "Deep Thinking": "deep thinking about the question",
    "Clear Explanation": "a clear explanation",
    "Well Supported": "well-supported ideas",
    "Creative Idea": "a creative idea",
    "Independent": "independent thinking",
    "Confident": "confidence in your answer",
    "Improving": "real improvement",
    "Good Use of Vocabulary": "good use of vocabulary",
    "Accurate Detail": "accurate detail",
    "Strong Participation": "strong participation",
}

RATING_FEEDBACK_OPENERS = {
    "A*": [
        "Excellent response.",
        "That was a standout answer.",
        "That was top-level work.",
    ],
    "A": [
        "Strong response.",
        "That was a very solid answer.",
        "You gave a strong answer there.",
    ],
    "B": [
        "That was a solid attempt.",
        "You gave a useful answer there.",
        "That answer had some clear strengths.",
    ],
    "C": [
        "Thank you for contributing.",
        "You made a positive start there.",
        "Thank you for having a go.",
    ],
}

RATING_FEEDBACK_CLOSERS = {
    "A*": "Keep using those strengths consistently.",
    "A": "Keep pushing those strengths and you can reach the very top level.",
    "B": "Keep building on those strengths and aim for an even fuller answer next time.",
    "C": "Use those strengths as your starting point and build one clear idea at a time.",
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
        self.CJK_FONT_FAMILY = self._pick_font_family(
            "Microsoft YaHei UI",
            "Microsoft YaHei",
            "Microsoft JhengHei UI",
            "Yu Gothic UI",
            self.FONT_FAMILY,
        )
        self.MONO_FONT_FAMILY = self._pick_font_family("Consolas", "Cascadia Mono", "Courier New", self.FONT_FAMILY)
        self.tk_scaling = self._detect_tk_scaling()
        self.display_scale = self._detect_display_scale()
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
        self.student_ungraded_by_class: dict[str, list[str]] = {}
        self.absent_students_by_class: dict[str, list[str]] = {}
        self.active_class: str | None = None
        self.session_students: list[str] = []
        self.student_grades: dict[str, str] = {}
        self.student_ungraded: list[str] = []
        self.absent_students: list[str] = []
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
            screen_h = max(1, float(self.root.winfo_screenheight()))
        except Exception:
            screen_h = 1080.0

        logical_screen_h = screen_h / max(1.0, self.tk_scaling)

        scale = logical_screen_h / 1080.0
        scale = max(1.00, min(scale, 1.55))

        # projection boost
        scale *= 1.15
        return max(1.00, min(scale, 1.75))

    def fs(self, px: int) -> int:
        return max(1, int(round(px * self.display_scale)))

    def ft(self, pt: int) -> int:
        return max(10, int(round(pt * self.ui_scale)))

    def _detect_tk_scaling(self) -> float:
        try:
            return max(1.0, float(self.root.winfo_fpixels("1i")) / 72.0)
        except Exception:
            return 1.0

    def _detect_display_scale(self) -> float:
        baseline_tk_scaling = 96.0 / 72.0
        return max(1.0, self.tk_scaling / baseline_tk_scaling)

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

    def _contains_cjk(self, text: str) -> bool:
        for ch in text or "":
            o = ord(ch)
            if (
                0x4E00 <= o <= 0x9FFF or
                0x3400 <= o <= 0x4DBF or
                0x3040 <= o <= 0x30FF or
                0xAC00 <= o <= 0xD7AF
            ):
                return True
        return False

    def _heading_font_family_for_text(self, text: str) -> str:
        if self._contains_cjk(text):
            return self.CJK_FONT_FAMILY
        return self.HEADING_FONT_FAMILY

    def f(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.FONT_FAMILY, self.ft(px), weight)
        return (self.FONT_FAMILY, self.ft(px))

    def hf(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.HEADING_FONT_FAMILY, self.ft(px), weight)
        return (self.HEADING_FONT_FAMILY, self.ft(px))

    def mf(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.MONO_FONT_FAMILY, self.ft(px), weight)
        return (self.MONO_FONT_FAMILY, self.ft(px))

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
        size = self.ft(start_px)
        min_size = self.ft(min_px)

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

        # Match Tk's scaling to the current monitor DPI instead of forcing a
        # fixed value that can become soft on external displays.
        try:
            self.root.tk.call("tk", "scaling", self.tk_scaling)
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
        ungraded = self.student_ungraded_by_class.get(class_name, [])
        absent = self.absent_students_by_class.get(class_name, [])
        return {
            "class_name": class_name or "No class selected",
            "roster_total": len(roster),
            "remaining": len(session_roster),
            "graded": len(grades),
            "ungraded": len(ungraded),
            "absent": len(absent),
        }

    def _format_metrics_summary(self, metrics: dict[str, int | str]) -> str:
        parts = [f"{metrics['remaining']} left", f"{metrics['graded']} graded"]
        if metrics.get("ungraded"):
            parts.append(f"{metrics['ungraded']} no grade")
        if metrics.get("absent"):
            parts.append(f"{metrics['absent']} absent")
        return " | ".join(parts)

    def _rating_meta(self, rating: str) -> dict[str, str]:
        meta = {
            "A*": {"label": "Excellent", "bg": self.palette["success"], "fg": "#1f301d"},
            "A": {"label": "Strong", "bg": self.palette["accent"], "fg": "#2d1c0d"},
            "B": {"label": "Secure", "bg": self.palette["warning"], "fg": "#33220a"},
            "C": {"label": "Needs support", "bg": self.palette["danger"], "fg": "#371814"},
        }
        return meta.get(rating, {"label": "Recorded", "bg": self.palette["panel_alt"], "fg": self.palette["text_light"]})

    def _natural_join(self, items: list[str]) -> str:
        if not items:
            return ""
        if len(items) == 1:
            return items[0]
        if len(items) == 2:
            return f"{items[0]} and {items[1]}"
        return f"{', '.join(items[:-1])}, and {items[-1]}"

    def _build_feedback_message(self, student_name: str, rating: str, labels: list[str]) -> str:
        ordered_labels = [label for label in FEEDBACK_LABELS if label in labels]
        label_phrases = [FEEDBACK_LABEL_PHRASES.get(label, label.lower()) for label in ordered_labels]
        opener = random.choice(RATING_FEEDBACK_OPENERS.get(rating, ["Good work."]))
        strengths = self._natural_join(label_phrases)
        closer = RATING_FEEDBACK_CLOSERS.get(rating, "Keep going.")
        return f"{student_name}, {opener} You showed {strengths}. {closer}"

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
            "NoGrade.TButton",
            font=self.hf(15, "bold"),
            foreground=p["text_dark"],
            background=p["bg_alt"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["bg_alt"], "#000000", 0.08),
            lightcolor=self._shade(p["bg_alt"], "#ffffff", 0.04),
            padding=self.fs(10),
        )
        self.style.configure(
            "Absent.TButton",
            font=self.hf(15, "bold"),
            foreground=p["text_dark"],
            background=p["bg_alt"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["bg_alt"], "#000000", 0.08),
            lightcolor=self._shade(p["bg_alt"], "#ffffff", 0.04),
            padding=self.fs(10),
        )
        self.style.map("Absent.TButton", background=[("active", self._shade(p["bg_alt"], "#000000", 0.03))])
        self.style.configure(
            "LabelToggleOff.TButton",
            font=self.hf(13, "bold"),
            foreground=p["text_dark"],
            background=p["panel"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["panel"], "#000000", 0.08),
            lightcolor=self._shade(p["panel"], "#ffffff", 0.04),
            padding=self.fs(8),
        )
        self.style.map(
            "LabelToggleOff.TButton",
            background=[("active", self._shade(p["panel_alt"], "#ffffff", 0.04))],
        )
        self.style.configure(
            "LabelToggleOn.TButton",
            font=self.hf(13, "bold"),
            foreground="#f8fcff",
            background=p["accent_alt"],
            bordercolor=p["accent_alt"],
            darkcolor=self._shade(p["accent_alt"], "#000000", 0.12),
            lightcolor=self._shade(p["accent_alt"], "#ffffff", 0.05),
            padding=self.fs(8),
        )
        self.style.map(
            "LabelToggleOn.TButton",
            background=[("active", self._shade(p["accent_alt"], "#ffffff", 0.08))],
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
        target_w = self.fs(WINDOW_WIDTH)
        min_w = self.fs(440)
        edge_gap = self.fs(24)
        top_gap = max(self.fs(12), self.fs(TOP_PADDING_Y + 12))

        window_w = min(target_w, max(min_w, work_w - edge_gap))
        window_h = max(self.fs(540), work_h - top_gap)
        return window_w, window_h

    def _configure_root_window(self):
        self.root.title("Random Student Selector")
        self.root.attributes("-topmost", True)

        work_left, work_top, work_w, work_h = self._desktop_work_area()
        window_w, window_h = self._root_window_size()

        x = work_left + max(0, work_w - window_w - self.fs(TOP_RIGHT_PADDING_X))
        y = work_top + self.fs(TOP_PADDING_Y)
        self.root.geometry(f"{window_w}x{window_h}+{x}+{y}")
        self.root.minsize(
            min(self.fs(WINDOW_WIDTH), window_w),
            min(max(self.fs(520), work_h - self.fs(80)), window_h),
        )
        self.root.maxsize(work_w, work_h)

    def _monitor_work_area_for_window(self, win) -> tuple[int, int, int, int]:
        try:
            if win and win.winfo_exists():
                hwnd = int(win.winfo_id())
                monitor = ctypes.windll.user32.MonitorFromWindow(hwnd, 2)
                if monitor:
                    class RECT(ctypes.Structure):
                        _fields_ = [
                            ("left", ctypes.c_long),
                            ("top", ctypes.c_long),
                            ("right", ctypes.c_long),
                            ("bottom", ctypes.c_long),
                        ]

                    class MONITORINFO(ctypes.Structure):
                        _fields_ = [
                            ("cbSize", ctypes.c_uint),
                            ("rcMonitor", RECT),
                            ("rcWork", RECT),
                            ("dwFlags", ctypes.c_uint),
                        ]

                    info = MONITORINFO()
                    info.cbSize = ctypes.sizeof(MONITORINFO)
                    ok = ctypes.windll.user32.GetMonitorInfoW(monitor, ctypes.byref(info))
                    if ok:
                        work = info.rcWork
                        return work.left, work.top, work.right - work.left, work.bottom - work.top
        except Exception:
            pass

        return self._desktop_work_area()

    def _capture_window_rect(self, win) -> tuple[int, int, int, int] | None:
        try:
            if win and win.winfo_exists():
                win.update_idletasks()
                return win.winfo_x(), win.winfo_y(), win.winfo_width(), win.winfo_height()
        except Exception:
            pass
        return None

    def _slot_window_height(self, work_area: tuple[int, int, int, int] | None = None) -> int:
        _, work_top, _, work_h = work_area or self._desktop_work_area()
        available_h = max(self.fs(480), work_h - max(self.fs(SLOT_PANEL_Y) - work_top, 0) - self.fs(24))
        return min(self.fs(SLOT_PANEL_H), max(self.fs(SLOT_PANEL_MIN_H), available_h))

    def _slot_window_size(self, work_area: tuple[int, int, int, int] | None = None) -> tuple[int, int]:
        work_left, work_top, work_w, work_h = work_area or self._desktop_work_area()
        available_h = max(self.fs(480), work_h - max(self.fs(SLOT_PANEL_Y) - work_top, 0) - self.fs(24))
        window_h = min(self.fs(SLOT_PANEL_H), max(self.fs(SLOT_PANEL_MIN_H), available_h))
        window_w = min(self.fs(SLOT_PANEL_W), max(self.fs(720), work_w - self.fs(40)))
        return window_w, window_h

    def _slot_window_geometry(
        self,
        parent_w: int,
        anchor_rect: tuple[int, int, int, int] | None = None,
        anchor_work_area: tuple[int, int, int, int] | None = None,
    ) -> str:
        work_left, work_top, work_w, work_h = anchor_work_area or self._desktop_work_area()
        window_w, window_h = self._slot_window_size(work_area=(work_left, work_top, work_w, work_h))

        x = work_left + max(0, work_w - window_w - self.fs(SLOT_PANEL_MARGIN_RIGHT))
        y = work_top + min(self.fs(SLOT_PANEL_Y), max(0, work_h - window_h - self.fs(12)))

        if anchor_rect is not None:
            anchor_x, anchor_y, _, _ = anchor_rect
            max_x = work_left + max(0, work_w - window_w)
            max_y = work_top + max(0, work_h - window_h)
            x = max(work_left, min(anchor_x, max_x))
            y = max(work_top, min(anchor_y, max_y))
            return f"{window_w}x{window_h}+{x}+{y}"

        root_x = self.root.winfo_x() if self.root.winfo_ismapped() else x
        root_w = self.root.winfo_width() if self.root.winfo_ismapped() else self.fs(WINDOW_WIDTH)
        control_left = root_x
        if x < control_left + root_w:
            x = max(work_left, control_left - window_w - self.fs(15))

        return f"{window_w}x{window_h}+{x}+{y}"

    def _grow_window_to_fit(self, win, min_bottom_margin: int = 12):
        """
        Expand a toplevel vertically when newly-added controls exceed the initial
        geometry, while keeping the window inside the desktop work area.
        """
        try:
            win.update_idletasks()
        except Exception:
            return

        current_w = win.winfo_width()
        current_h = win.winfo_height()
        needed_h = win.winfo_reqheight()
        if needed_h <= current_h:
            return

        work_left, work_top, work_w, work_h = self._monitor_work_area_for_window(win)
        target_h = min(work_h, needed_h)
        if target_h <= current_h:
            return

        x = max(work_left, min(win.winfo_x(), work_left + work_w - current_w))
        y_max = work_top + max(0, work_h - target_h - min_bottom_margin)
        y = max(work_top, min(win.winfo_y(), y_max))
        win.geometry(f"{current_w}x{target_h}+{x}+{y}")

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
                self._format_metrics_summary(metrics)
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
        # Use two columns so labels can display fully: top row for attendance/summary,
        # bottom row for sound controls.
        for col in range(2):
            controls.grid_columnconfigure(col, weight=1, uniform="ctrl")
        # First row: attendance and summary
        ttk.Button(
            controls,
            text="Attendance",
            style="Utility.TButton",
            command=self._take_attendance_sequential,
        ).grid(row=0, column=0, sticky="ew", padx=(0, card_gap), ipady=d(4))
        ttk.Button(
            controls,
            text="View Summary",
            style="Utility.TButton",
            command=self._show_grades_summary,
        ).grid(row=0, column=1, sticky="ew", padx=(card_gap, 0), ipady=d(4))
        # Second row: sound controls (add vertical spacing above this row)
        ttk.Button(
            controls,
            text="Play Intro",
            style="Utility.TButton",
            command=self._play_intro,
        ).grid(row=1, column=0, sticky="ew", padx=(0, card_gap), ipady=d(4), pady=(d(6), 0))
        ttk.Button(
            controls,
            text="Play Closing",
            style="Utility.TButton",
            command=self._play_closing,
        ).grid(row=1, column=1, sticky="ew", padx=(card_gap, 0), ipady=d(4), pady=(d(6), 0))

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

    def _show_attendance_dialog(self):
        selected = (self.selected_class.get() or "").strip()
        if selected == "Select a Class" or selected not in self.classes:
            Messagebox.show_error(title="Error", message="Please select a valid class before taking attendance.")
            return

        class_name = selected
        roster = list(self.classes.get(class_name, []))

        win = ttk.Toplevel(self.root)
        win.title(f"{class_name} - Attendance")
        win.attributes("-topmost", True)
        win.geometry(f"{self.fs(560)}x{self.fs(520)}+{self.fs(40)}+{self.fs(80)}")
        win.configure(bg=self.palette["bg"])

        p = self.palette

        outer = tk.Frame(win, bg=p["bg"], padx=self.fs(12), pady=self.fs(12))
        outer.pack(fill="both", expand=True)
        outer.grid_columnconfigure(0, weight=1)

        tk.Label(outer, text=f"Mark absent students for {class_name}", font=self.hf(16, "bold"), bg=p["bg"], fg=p["text_light"]).pack(anchor="w", pady=(0, self.fs(8)))

        # Scrollable area for student checkboxes
        list_frame = tk.Frame(outer, bg=p["panel"], padx=self.fs(8), pady=self.fs(8))
        list_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_frame, bg=p["panel"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=p["panel"])

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        vars_by_student: dict[str, tk.IntVar] = {}
        pre_absent = set(self.absent_students_by_class.get(class_name, []))
        for name in roster:
            var = tk.IntVar(value=1 if name in pre_absent else 0)
            # Use a flat, indicatorless checkbutton so no white checkbox shows
            chk = tk.Checkbutton(
                scrollable,
                text=name,
                variable=var,
                bg=p["panel"],
                fg=p["text_light"],
                activebackground=p["panel"],
                selectcolor=p["panel"],
                anchor="w",
                bd=0,
                relief="flat",
                highlightthickness=0,
                indicatoron=False,
                activeforeground=p["text_light"],
            )
            chk.pack(fill="x", padx=self.fs(6), pady=self.fs(2))
            vars_by_student[name] = var

        btns = tk.Frame(outer, bg=p["bg"])
        btns.pack(fill="x", pady=(self.fs(8), 0))
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        def save():
            absent = [n for n, v in vars_by_student.items() if v.get()]
            self.absent_students_by_class.setdefault(class_name, [])
            self.absent_students_by_class[class_name] = absent

            # If a session roster exists, remove absent students immediately
            if class_name in self.session_students_by_class:
                current = [s for s in self.session_students_by_class[class_name] if s not in absent]
                self.session_students_by_class[class_name] = current

            win.destroy()
            self._build_main_screen()

        def cancel():
            win.destroy()

        ttk.Button(btns, text="Save", style="PrimaryAction.TButton", command=save).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(8)))
        ttk.Button(btns, text="Cancel", style="Utility.TButton", command=cancel).grid(row=0, column=1, sticky="ew")

    def _take_attendance_sequential(self):
        selected = (self.selected_class.get() or "").strip()
        if selected == "Select a Class" or selected not in self.classes:
            Messagebox.show_error(title="Error", message="Please select a valid class before taking attendance.")
            return

        class_name = selected
        roster = list(self.classes.get(class_name, []))
        if not roster:
            Messagebox.show_info(title="No Students", message="This class has no students in the roster.")
            return

        win = ttk.Toplevel(self.root)
        win.title(f"{class_name} - Roll Call")
        win.attributes("-topmost", True)
        win.geometry(f"{self.fs(720)}x{self.fs(420)}+{self.fs(40)}+{self.fs(80)}")
        win.configure(bg=self.palette["bg"])

        p = self.palette

        # State
        idx = 0
        absent: list[str] = []

        title = tk.Label(
            win,
            text=class_name,
            font=self.hf(20, "bold"),
            bg=p["bg"],
            fg=p["text_light"],
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        title.pack(anchor="n", pady=(self.fs(8), 0))

        name_frame = tk.Frame(win, bg=p["panel"], padx=self.fs(16), pady=self.fs(16))
        name_frame.pack(fill="both", expand=True, padx=self.fs(12), pady=self.fs(12))

        name_label = tk.Label(
            name_frame,
            text="",
            font=self.hf(48, "bold"),
            bg=p["panel"],
            fg=p["text_dark"],
            wraplength=self.fs(680),
            justify="center",
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        name_label.pack(expand=True)

        info_label = tk.Label(
            win,
            text="Press Present or Absent for each student",
            font=self.f(12),
            bg=p["bg"],
            fg=p["text_muted"],
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        info_label.pack()

        btn_frame = tk.Frame(win, bg=p["bg"]) 
        btn_frame.pack(fill="x", pady=(self.fs(8), self.fs(12)))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        def _update_display():
            nonlocal idx
            if idx >= len(roster):
                _finish()
                return
            name = roster[idx]
            name_label.config(text=name)

        def _mark_present():
            nonlocal idx
            idx += 1
            _update_display()

        def _mark_absent():
            nonlocal idx
            absent.append(roster[idx])
            idx += 1
            _update_display()

        def _finish():
            # Save absent list
            self.absent_students_by_class[class_name] = absent

            # Rebuild the session roster from the master class list minus absentees
            master = list(self.classes.get(class_name, []))
            self.session_students_by_class[class_name] = [s for s in master if s not in absent]

            # Close the attendance (roll-call) window now that we're finished
            try:
                win.destroy()
            except Exception:
                pass

            # Immediately refresh the main dock so counters/summary update
            try:
                # keep the combobox selection consistent
                self.selected_class.set(class_name)
                self._build_main_screen()
                self.root.update_idletasks()
            except Exception:
                pass

            # Show results in a copyable text area
            res_win = ttk.Toplevel(self.root)
            res_win.title(f"{class_name} - Absent Students")
            res_win.attributes("-topmost", True)
            res_win.geometry(f"{self.fs(640)}x{self.fs(420)}+{self.fs(60)}+{self.fs(100)}")
            res_win.configure(bg=self.palette["bg"])

            p2 = self.palette
            outer = tk.Frame(res_win, bg=p2["bg"], padx=self.fs(12), pady=self.fs(12))
            outer.pack(fill="both", expand=True)
            tk.Label(outer, text=f"Absent students ({len(absent)})", font=self.hf(16, "bold"), bg=p2["bg"], fg=p2["text_light"]).pack(anchor="w")

            # Style the text area to match the dialog (avoid default white background)
            text = tk.Text(
                outer,
                height=12,
                wrap="word",
                bg=p2["panel"],
                fg=p2["text_light"],
                bd=0,
                relief="flat",
                highlightthickness=0,
                insertbackground=p2["text_light"],
                selectbackground=self._shade(p2["accent_alt"], p2["panel"], 0.3),
            )
            text.pack(fill="both", expand=True, pady=(self.fs(8), self.fs(8)))
            absent_text = "\n".join(absent) if absent else "(none)"
            text.insert("1.0", absent_text)
            try:
                text.configure(state="disabled")
            except Exception:
                pass

            def _copy():
                try:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(absent_text)
                    Messagebox.show_info(title="Copied", message="Absent list copied to clipboard.")
                except Exception:
                    Messagebox.show_error(title="Error", message="Unable to copy to clipboard.")

            def _close_all():
                try:
                    res_win.destroy()
                except Exception:
                    pass
                try:
                    win.destroy()
                except Exception:
                    pass
                self._build_main_screen()

            btns = tk.Frame(outer, bg=p2["bg"]) 
            btns.pack(fill="x")
            ttk.Button(btns, text="Copy Absent List", style="PrimaryAction.TButton", command=_copy).pack(side="left", expand=True, fill="x", padx=(0, self.fs(8)))
            ttk.Button(btns, text="Close", style="Utility.TButton", command=_close_all).pack(side="left", expand=True, fill="x")

        ttk.Button(btn_frame, text="Present", style="Utility.TButton", command=_mark_present).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(8)), ipady=self.fs(8))
        ttk.Button(btn_frame, text="Absent", style="PrimaryAction.TButton", command=_mark_absent).grid(row=0, column=1, sticky="ew", ipady=self.fs(8))

        # start
        _update_display()

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

        # Initialize (or reinitialize) the session roster from the master class list
        # and always filter out students previously marked absent.
        roster = list(self.classes.get(class_name, []))
        if not roster:
            Messagebox.show_error(title="Error", message=f"No students found for {class_name}.")
            return

        absent = set(self.absent_students_by_class.get(class_name, []))
        if absent:
            roster = [s for s in roster if s not in absent]

        self.session_students_by_class[class_name] = roster

        if class_name not in self.student_grades_by_class:
            self.student_grades_by_class[class_name] = {}
        if class_name not in self.student_ungraded_by_class:
            self.student_ungraded_by_class[class_name] = []
        if class_name not in self.absent_students_by_class:
            self.absent_students_by_class[class_name] = []

        self.active_class = class_name
        self.session_students = self.session_students_by_class[class_name]
        self.student_grades = self.student_grades_by_class[class_name]
        self.student_ungraded = self.student_ungraded_by_class[class_name]
        self.absent_students = self.absent_students_by_class[class_name]

        if not self.session_students:
            self._show_grades_summary()
            return

        self._next_student(class_name)

    def _next_student(
        self,
        class_name: str,
        anchor_rect: tuple[int, int, int, int] | None = None,
        anchor_work_area: tuple[int, int, int, int] | None = None,
    ):
        if self.exit_requested:
            self._build_main_screen()
            return

        if not self.session_students:
            self._show_grades_summary()
            return

        final_student = random.choice(self.session_students)
        self._show_slot_window(
            class_name,
            final_student,
            anchor_rect=anchor_rect,
            anchor_work_area=anchor_work_area,
        )

    # ---------- Slot window ----------

    def _show_slot_window(
        self,
        class_name: str,
        final_student: str,
        anchor_rect: tuple[int, int, int, int] | None = None,
        anchor_work_area: tuple[int, int, int, int] | None = None,
    ):
        duration = self._get_duration_seconds()
        slot_sound = self._slot_sound_for_duration(duration)
        slot_effect_enabled = bool(self.slot_effect_enabled_var.get())
        target_work_area = anchor_work_area or self._desktop_work_area()
        window_w, window_h = self._slot_window_size(work_area=target_work_area)

        win = ttk.Toplevel(self.root)
        win.title(f"{class_name} - Select Student")
        win.attributes("-topmost", True)
        compact = window_h < 700 or window_w < 900
        win.geometry(
            self._slot_window_geometry(
                parent_w=WINDOW_WIDTH,
                anchor_rect=anchor_rect,
                anchor_work_area=anchor_work_area,
            )
        )
        _, _, work_w, work_h = target_work_area
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
        reserved_h = self.fs(184 if compact else 230)
        usable_h = max(self.fs(170), window_h - reserved_h)
        row_min = self.fs(42 if compact else 48)
        row_max = self.fs(52 if compact else 62)
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


        def _format_name(name: str) -> str:
            s = (name or "").strip()
            if not s:
                return s
            if " " in s:
                has_cjk = self._contains_cjk(s)
                has_latin = any(("A" <= ch <= "Z") or ("a" <= ch <= "z") for ch in s)
                if has_cjk and has_latin:
                    left, right = s.split(None, 1)
                    return f"{left} · {right}"
            return s

        font_cache: dict[tuple[str, str, int, int, str], tuple] = {}
        max_text_w = canvas_w - (pad * 2) - self.fs(110)

        def _fit_font(text: str, base_px: int, min_px: int, weight="bold") -> tuple:
            family = self._heading_font_family_for_text(text)
            key = (family, text, base_px, min_px, weight)
            if key in font_cache:
                return font_cache[key]

            size = self.ft(base_px)
            min_size = self.ft(min_px)
            fnt = tkfont.Font(family=family, size=size, weight=weight)

            while size > min_size and fnt.measure(text) > max_text_w:
                size -= 1
                fnt.configure(size=size)

            font_cache[key] = (family, size, weight)
            return font_cache[key]

        pool = self.session_students[:] if self.session_students else [final_student]
        prev_raw = random.choice(pool)
        cur_raw = random.choice(pool)
        next_raw = random.choice(pool)

        dim = secondary
        small_px = 21
        big_px = 38
        min_small_px = 16
        min_big_px = 26

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
        button_frame.configure(style="SlotBg.TFrame", padding=(0, self.fs(14), 0, 0))
        shell = tk.Frame(button_frame, bg=p["bg"])
        shell.pack(fill="x")

        card = tk.Frame(
            shell,
            bg=p["panel_alt"],
            padx=self.fs(18),
            pady=self.fs(16),
            highlightthickness=1,
            highlightbackground=p["line_soft"],
        )
        card.pack(fill="x")
        for col in range(4):
            card.columnconfigure(col, weight=1, uniform="rate")

        work_area = self._monitor_work_area_for_window(win)
        window_h = self._slot_window_height(work_area=work_area)
        compact = window_h < 700
        button_ipady = self.fs(10) if compact else self.fs(14)
        utility_ipady = self.fs(8) if compact else self.fs(10)
        button_gap = self.fs(10 if compact else 12)
        absent_students = self.absent_students_by_class.setdefault(class_name, [])
        ungraded_students = self.student_ungraded_by_class.setdefault(class_name, [])

        def _clear_non_grade_marks():
            if student_name in absent_students:
                absent_students.remove(student_name)
            if student_name in ungraded_students:
                ungraded_students.remove(student_name)

        def apply_rating(rating: str):
            _clear_non_grade_marks()
            self.student_grades[student_name] = rating
            self.sound.play_music_once(RATING_SOUNDS.get(rating, ""))

            msg_list = self.messages.get(rating) or []
            msg = random.choice(msg_list) if msg_list else "Noted."
            anchor_rect = self._capture_window_rect(win)
            anchor_work_area = self._monitor_work_area_for_window(win)

            try:
                win.destroy()
            except Exception:
                pass

            self._show_message_popup(
                title="Feedback",
                message=msg,
                class_name=class_name,
                anchor_rect=anchor_rect,
                anchor_work_area=anchor_work_area,
            )

        def mark_no_grade():
            self.student_grades.pop(student_name, None)
            if student_name in absent_students:
                absent_students.remove(student_name)
            if student_name not in ungraded_students:
                ungraded_students.append(student_name)
            anchor_rect = self._capture_window_rect(win)
            anchor_work_area = self._monitor_work_area_for_window(win)

            try:
                win.destroy()
            except Exception:
                pass

            self._show_message_popup(
                title="No Grade",
                message=f"No grade recorded for {student_name} this round.",
                class_name=class_name,
                anchor_rect=anchor_rect,
                anchor_work_area=anchor_work_area,
            )

        def mark_absent():
            self.student_grades.pop(student_name, None)
            if student_name in ungraded_students:
                ungraded_students.remove(student_name)
            if student_name not in absent_students:
                absent_students.append(student_name)
            anchor_rect = self._capture_window_rect(win)
            anchor_work_area = self._monitor_work_area_for_window(win)

            try:
                win.destroy()
            except Exception:
                pass

            self._show_message_popup(
                title="Absent",
                message=f"{student_name} was marked absent and removed from today's list.",
                class_name=class_name,
                anchor_rect=anchor_rect,
                anchor_work_area=anchor_work_area,
            )

        actions = [
            ("A*", self._rating_meta("A*")["label"], "GradeAStar.TButton", lambda: apply_rating("A*")),
            ("A", self._rating_meta("A")["label"], "GradeA.TButton", lambda: apply_rating("A")),
            ("B", self._rating_meta("B")["label"], "GradeB.TButton", lambda: apply_rating("B")),
            ("C", self._rating_meta("C")["label"], "GradeC.TButton", lambda: apply_rating("C")),
            ("No Grade", "Skip grading", "NoGrade.TButton", mark_no_grade),
            ("Absent", "Remove for today", "Absent.TButton", mark_absent),
        ]

        title_wrap = max(self.fs(300), self._slot_window_size(work_area=work_area)[0] - self.fs(180))
        tk.Label(
            card,
            text=f"Choose an outcome for {student_name}",
            font=self.hf(18 if compact else 20, "bold"),
            bg=p["panel_alt"],
            fg=p["text_light"],
            anchor="w",
            justify="left",
            wraplength=title_wrap,
        ).grid(row=0, column=0, columnspan=4, sticky="w")

        tk.Label(
            card,
            text="Grades are the primary action. Attendance shortcuts sit below.",
            font=self.f(11 if compact else 12),
            bg=p["panel_alt"],
            fg=p["text_muted"],
            anchor="w",
            justify="left",
            wraplength=title_wrap,
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(self.fs(4), self.fs(14)))

        grade_actions = actions[:4]
        utility_actions = actions[4:]

        for col, (title, subtitle, style, command) in enumerate(grade_actions):
            ttk.Button(
                card,
                text=f"{title}\n{subtitle}",
                style=style,
                command=command,
            ).grid(
                row=2,
                column=col,
                sticky="nsew",
                padx=(0 if col == 0 else button_gap, 0),
                pady=(0, self.fs(10)),
                ipady=button_ipady,
            )

        for i, (title, subtitle, style, command) in enumerate(utility_actions):
            start_col = i * 2
            ttk.Button(
                card,
                text=f"{title}\n{subtitle}",
                style=style,
                command=command,
            ).grid(
                row=3,
                column=start_col,
                columnspan=2,
                sticky="ew",
                padx=(0, button_gap // 2) if i == 0 else (button_gap // 2, 0),
                ipady=utility_ipady,
            )

        self._grow_window_to_fit(win)

    # ---------- Feedback popup ----------

    def _show_message_popup(
        self,
        title: str,
        message: str,
        class_name: str,
        anchor_rect: tuple[int, int, int, int] | None = None,
        anchor_work_area: tuple[int, int, int, int] | None = None,
    ):
        target_work_area = anchor_work_area or self._desktop_work_area()
        msg_win = ttk.Toplevel(self.root)
        msg_win.title(f"{class_name} - {title}")
        msg_win.attributes("-topmost", True)
        msg_win.geometry(
            self._slot_window_geometry(
                parent_w=WINDOW_WIDTH,
                anchor_rect=anchor_rect,
                anchor_work_area=anchor_work_area,
            )
        )
        msg_win.minsize(self.fs(760), self._slot_window_height(work_area=target_work_area))
        msg_win.configure(bg=self.palette["bg"])
        p = self.palette
        metrics = self._class_metrics(class_name)
        message_wrap = self._slot_window_size(work_area=target_work_area)[0] - self.fs(160)

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
            text=self._format_metrics_summary(metrics),
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
        tk.Label(quote, text="“", font=self.hf(56, "bold"), bg=p["panel"], fg=p["accent"]).pack(anchor="center")
        tk.Label(
            quote,
            text=message,
            font=self.f(30),
            bg=p["panel"],
            fg=p["text_light"],
            wraplength=message_wrap,
            justify="center",
        ).pack(fill="both", expand=True, pady=(0, self.fs(10)))

        try:
            quote.destroy()
        except Exception:
            pass

        tk.Label(
            body_card,
            text=message,
            font=self.f(30),
            bg=p["panel"],
            fg=p["text_light"],
            wraplength=message_wrap,
            justify="center",
        ).grid(row=0, column=0, sticky="nsew")

        btns = tk.Frame(outer, bg=p["bg"])
        btns.grid(row=2, column=0, sticky="ew")
        btns.grid_columnconfigure(0, weight=1)
        btns.grid_columnconfigure(1, weight=1)

        def next_student():
            next_anchor_rect = self._capture_window_rect(msg_win)
            next_anchor_work_area = self._monitor_work_area_for_window(msg_win)
            try:
                msg_win.destroy()
            except Exception:
                pass
            self._next_student(
                class_name,
                anchor_rect=next_anchor_rect,
                anchor_work_area=next_anchor_work_area,
            )

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

    # ---------- Session summary ----------

    def _show_grades_summary(self):
        self.exit_requested = False

        win = ttk.Toplevel(self.root)
        win.title("Session Summary")
        win.attributes("-topmost", True)
        win.geometry(f"{self.fs(860)}x{self.fs(760)}+{self.fs(30)}+{self.fs(80)}")
        win.minsize(self.fs(760), self.fs(680))
        win.configure(bg=self.palette["bg"])
        p = self.palette

        class_name = self.active_class or self._active_or_selected_class() or "Current Session"
        if self.active_class and self.active_class in self.student_grades_by_class:
            grades = self.student_grades_by_class.get(self.active_class, {})
            ungraded = self.student_ungraded_by_class.get(self.active_class, [])
            absent = self.absent_students_by_class.get(self.active_class, [])
            remaining = len(self.session_students_by_class.get(self.active_class, []))
            roster_total = len(self.classes.get(self.active_class, []))
        elif class_name in self.student_grades_by_class:
            grades = self.student_grades_by_class.get(class_name, {})
            ungraded = self.student_ungraded_by_class.get(class_name, [])
            absent = self.absent_students_by_class.get(class_name, [])
            remaining = len(self.session_students_by_class.get(class_name, self.classes.get(class_name, [])))
            roster_total = len(self.classes.get(class_name, []))
        else:
            grades = self.student_grades
            ungraded = self.student_ungraded
            absent = self.absent_students
            remaining = len(self.session_students)
            roster_total = len(grades) + len(ungraded) + len(absent) + remaining

        def on_escape(event=None):
            try:
                win.destroy()
            except Exception:
                pass

        win.bind("<Escape>", on_escape)

        counts = {rating: 0 for rating in ("A*", "A", "B", "C")}
        for rating in grades.values():
            counts[rating] = counts.get(rating, 0) + 1
        metrics = {
            "remaining": remaining,
            "graded": len(grades),
            "ungraded": len(ungraded),
            "absent": len(absent),
        }

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
            text=self._format_metrics_summary(metrics) if roster_total else "No class session yet",
            font=self.f(13),
            bg=p["bg"],
            fg=p["text_muted"],
        ).grid(row=1, column=0, sticky="w", pady=(self.fs(4), 0))

        chip_row = tk.Frame(header, bg=p["bg"])
        chip_row.grid(row=0, column=1, rowspan=2, sticky="e", padx=(self.fs(18), 0))
        chip_data = [
            ("A*", counts.get("A*", 0), self._rating_meta("A*")["bg"]),
            ("A", counts.get("A", 0), self._rating_meta("A")["bg"]),
            ("B", counts.get("B", 0), self._rating_meta("B")["bg"]),
            ("C", counts.get("C", 0), self._rating_meta("C")["bg"]),
            ("No Grade", len(ungraded), p["text_dark"]),
            ("Absent", len(absent), p["danger"]),
        ]
        for idx, (label, value, accent) in enumerate(chip_data):
            chip = tk.Frame(chip_row, bg=p["panel_alt"], padx=self.fs(12), pady=self.fs(10))
            chip.grid(
                row=idx // 3,
                column=idx % 3,
                padx=(0 if idx % 3 == 0 else self.fs(8), 0),
                pady=(0 if idx < 3 else self.fs(8), 0),
            )
            tk.Label(chip, text=label, font=self.hf(16, "bold"), bg=p["panel_alt"], fg=accent).pack(anchor="center")
            tk.Label(chip, text=str(value), font=self.f(13, "bold"), bg=p["panel_alt"], fg=p["text_light"]).pack(anchor="center", pady=(self.fs(4), 0))

        body = tk.Frame(outer, bg=p["panel"], padx=self.fs(18), pady=self.fs(18))
        body.grid(row=1, column=0, sticky="nsew", pady=(0, self.fs(16)))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        list_shell = tk.Frame(body, bg=p["bg_alt"])
        list_shell.grid(row=0, column=0, sticky="nsew")
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        if not (grades or ungraded or absent):
            empty = tk.Frame(list_shell, bg=p["bg_alt"], padx=self.fs(24), pady=self.fs(24))
            empty.grid(row=0, column=0, sticky="nsew")
            tk.Label(empty, text="No session records yet.", font=self.hf(20, "bold"), bg=p["bg_alt"], fg=p["text_light"]).pack(anchor="center", pady=(self.fs(24), 0))
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

            row_number = 1

            def add_section(title: str):
                section = tk.Frame(rows, bg=p["bg_alt"], padx=self.fs(14), pady=self.fs(12))
                section.pack(fill="x", expand=True)
                tk.Label(section, text=title, font=self.hf(14, "bold"), bg=p["bg_alt"], fg=p["accent_alt"], anchor="w").pack(anchor="w")

            def add_row(student: str, pill_text: str, pill_bg: str, pill_fg: str):
                nonlocal row_number
                idx = row_number
                row_bg = p["panel"] if idx % 2 else p["panel_alt"]
                row = tk.Frame(rows, bg=row_bg, padx=self.fs(14), pady=self.fs(12))
                row.pack(fill="x", expand=True)
                row.grid_columnconfigure(1, weight=1)

                tk.Label(row, text=f"{idx:02d}", font=self.mf(12, "bold"), bg=row_bg, fg=p["text_muted"], width=4, anchor="w").grid(row=0, column=0, sticky="w")
                tk.Label(row, text=student, font=self.hf(16, "bold"), bg=row_bg, fg=p["text_light"], anchor="w").grid(row=0, column=1, sticky="w")

                pill = tk.Label(
                    row,
                    text=pill_text,
                    font=self.hf(12, "bold"),
                    bg=pill_bg,
                    fg=pill_fg,
                    padx=self.fs(10),
                    pady=self.fs(5),
                )
                pill.grid(row=0, column=2, sticky="e")
                row_number += 1

            if grades:
                add_section("Grades")
                for student, rating in grades.items():
                    meta = self._rating_meta(rating)
                    add_row(student, f"{rating}  {meta['label']}", meta["bg"], meta["fg"])

            if ungraded:
                add_section("No Grade")
                for student in ungraded:
                    add_row(student, "No Grade", p["bg_alt"], p["text_dark"])

            if absent:
                add_section("Absent")
                for student in absent:
                    add_row(student, "Absent", p["danger"], "#fff6ef")

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
    enable_windows_high_dpi()
    style = Style(theme=THEME)
    root = style.master
    _ = InvisibleHandApp(root, style)
    root.mainloop()


if __name__ == "__main__":
    main()

