import os
import csv
import random
import time
import tkinter as tk
import tkinter.font as tkfont
import math
from collections import deque

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

WINDOW_WIDTH = 500
TOP_RIGHT_PADDING_X = 20
TOP_PADDING_Y = 0

SLOT_PANEL_W = 780
SLOT_PANEL_H = 620
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

        # Typography / classroom projection
        self.FONT_FAMILY = "Segoe UI"
        self.ui_scale = self._compute_classroom_ui_scale()
        self._apply_classroom_font_defaults()

        # State
        self.classes = load_students_by_class(STUDENTS_CSV)
        self.messages = load_messages_by_rating(MESSAGES_CSV)

        self.selected_class = ttk.StringVar(value="Select a Class")
        self.minutes_var = ttk.IntVar(value=0)
        self.seconds_var = ttk.IntVar(value=10)

        self.sound_enabled_var = tk.BooleanVar(value=True)
        self.sound.set_enabled(self.sound_enabled_var.get())

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

    def f(self, px: int, weight: str | None = None) -> tuple:
        if weight:
            return (self.FONT_FAMILY, self.fs(px), weight)
        return (self.FONT_FAMILY, self.fs(px))

    def _fit_font_to_width(self, text: str, max_px: int, start_px: int, min_px: int, weight: str = "bold") -> tuple:
        """
        Prevent the title from clipping in the narrow dock window by shrinking
        the font until it fits (or hits min).
        """
        size = self.fs(start_px)
        min_size = self.fs(min_px)

        test_font = tkfont.Font(family=self.FONT_FAMILY, size=size, weight=weight)
        while size > min_size and test_font.measure(text) > max_px:
            size -= 1
            test_font.configure(size=size)

        return (self.FONT_FAMILY, size, weight)

    def _apply_classroom_font_defaults(self):
        base = self.f(20)
        self.root.option_add("*Font", base)
        self.root.option_add("*Listbox.Font", base)
        self.root.option_add("*TCombobox*Listbox.Font", self.f(22))
        self.root.option_add("*Text.Font", base)

        # TTK defaults
        self.style.configure(".", font=base)
        self.style.configure("TLabel", font=base)
        self.style.configure("secondary.TLabel", font=self.f(18))

        # Labelframe titles
        self.style.configure("TLabelframe.Label", font=self.f(20, "bold"))
        for bs in ("primary", "secondary", "info", "success", "warning", "danger"):
            self.style.configure(f"{bs}.TLabelframe.Label", font=self.f(20, "bold"))

        # Inputs
        self.style.configure("TCombobox", font=self.f(24))
        self.style.configure("TSpinbox", font=self.f(24))
        self.style.configure("TEntry", font=self.f(24))

        # Buttons
        self.style.configure("TButton", font=self.f(20, "bold"))
        for bs in ("primary", "secondary", "success", "info", "warning", "danger"):
            self.style.configure(f"{bs}.TButton", font=self.f(20, "bold"))

        # Checkbuttons/toggles
        self.style.configure("TCheckbutton", font=self.f(20))

        # Slot window progressbar: thicker for projection
        try:
            self.style.configure("Slot.Horizontal.TProgressbar", thickness=self.fs(16))
        except Exception:
            pass

        # Small scaling nudge for projector clarity
        try:
            self.root.tk.call("tk", "scaling", 1.15)
        except Exception:
            pass

    # ---------- Window / Layout ----------

    def _configure_root_window(self):
        self.root.title("Random Student Selector")
        self.root.attributes("-topmost", True)

        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        window_h = screen_h - 50

        x = screen_w - WINDOW_WIDTH - TOP_RIGHT_PADDING_X
        y = TOP_PADDING_Y
        self.root.geometry(f"{WINDOW_WIDTH}x{window_h}+{x}+{y}")

    def _slot_window_geometry(self, parent_w: int) -> str:
        screen_w = self.root.winfo_screenwidth()

        x = screen_w - SLOT_PANEL_W - SLOT_PANEL_MARGIN_RIGHT
        y = SLOT_PANEL_Y

        control_left = screen_w - WINDOW_WIDTH - TOP_RIGHT_PADDING_X
        if x < control_left + WINDOW_WIDTH:
            x = max(0, control_left - SLOT_PANEL_W - 15)

        return f"{SLOT_PANEL_W}x{SLOT_PANEL_H}+{x}+{y}"

    def _clear_root(self):
        for w in self.root.winfo_children():
            w.destroy()

    # ---------- Sound toggle ----------

    def _on_sound_toggle(self):
        self.sound.set_enabled(self.sound_enabled_var.get())

    # ---------- Main Screen ----------

    def _build_main_screen(self):
        self._clear_root()

        outer = ttk.Frame(self.root, padding=(16, 16, 16, 16))
        outer.pack(fill="both", expand=True)

        # Content + footer separation so footer doesn't get pushed off-screen
        content = ttk.Frame(outer)
        content.pack(fill="both", expand=True)

        footer = ttk.Frame(outer)
        footer.pack(fill="x", side="bottom", pady=(10, 0))

        # Header
        header = ttk.Frame(content, padding=(10, 10))
        header.pack(fill="x")

        max_title_px = WINDOW_WIDTH - 70  # account for outer padding + safe margin
        title_text = "Random Student Selector"
        title_font = self._fit_font_to_width(
            text=title_text,
            max_px=max_title_px,
            start_px=44,
            min_px=30,
            weight="bold",
        )

        ttk.Label(
            header,
            text=title_text,
            font=title_font,
            bootstyle="primary",
            wraplength=max_title_px,
            justify="left",
        ).pack(anchor="w")

        ttk.Separator(content).pack(fill="x", pady=(8, 14))

        # Sound toggle row
        sound_row = ttk.Frame(content, padding=(10, 0, 10, 0))
        sound_row.pack(fill="x")

        ttk.Label(
            sound_row,
            text="Sound",
            font=self.f(22),
            bootstyle="secondary"
        ).pack(side="left")

        try:
            ttk.Checkbutton(
                sound_row,
                text="",
                variable=self.sound_enabled_var,
                command=self._on_sound_toggle,
                bootstyle="success-round-toggle",
            ).pack(side="right")
        except Exception:
            ttk.Checkbutton(
                sound_row,
                text="",
                variable=self.sound_enabled_var,
                command=self._on_sound_toggle,
                bootstyle="round-toggle",
            ).pack(side="right")

        # Class card
        class_card = ttk.Labelframe(content, text="Class", padding=(14, 14), bootstyle="info")
        class_card.pack(fill="x", padx=10, pady=(14, 14))

        class_dropdown = ttk.Combobox(
            class_card,
            textvariable=self.selected_class,
            state="readonly",
            font=self.f(24),
            bootstyle="info",
            values=["Select a Class"] + sorted(self.classes.keys()),
        )
        class_dropdown.pack(fill="x")
        class_dropdown.bind("<<ComboboxSelected>>", lambda e: self._on_class_selected())

        # Timer card
        timer_card = ttk.Labelframe(content, text="Timer", padding=(14, 14), bootstyle="info")
        timer_card.pack(fill="x", padx=10, pady=(0, 16))

        tf = ttk.Frame(timer_card)
        tf.pack(fill="x")

        ttk.Label(tf, text="Min", font=self.f(20)).grid(row=0, column=0, sticky="w", padx=(0, 18))
        ttk.Label(tf, text="Sec", font=self.f(20)).grid(row=0, column=1, sticky="w")

        ttk.Spinbox(tf, from_=0, to=60, textvariable=self.minutes_var, width=6, font=self.f(24)).grid(
            row=1, column=0, sticky="w", padx=(0, 18), pady=(4, 0)
        )
        ttk.Spinbox(tf, from_=0, to=59, textvariable=self.seconds_var, width=6, font=self.f(24)).grid(
            row=1, column=1, sticky="w", pady=(4, 0)
        )

        # Main action
        ttk.Button(
            content,
            text="START RANDOM PICK • 随机抽选",
            bootstyle="success",
            command=self._start_session
        ).pack(fill="x", padx=10, ipady=18, pady=(0, 16))

        # Secondary controls (grid for consistent sizing)
        controls = ttk.Frame(content)
        controls.pack(fill="x", padx=10, pady=(0, 10))

        controls.columnconfigure(0, weight=1, uniform="ctrl")
        controls.columnconfigure(1, weight=1, uniform="ctrl")
        controls.columnconfigure(2, weight=1, uniform="ctrl")

        ttk.Button(controls, text="Intro", bootstyle="secondary", command=self._play_intro).grid(
            row=0, column=0, sticky="ew", padx=(0, 8), ipady=14
        )
        ttk.Button(controls, text="Grades", bootstyle="secondary", command=self._show_grades_summary).grid(
            row=0, column=1, sticky="ew", padx=8, ipady=14
        )
        ttk.Button(controls, text="Closing", bootstyle="secondary", command=self._play_closing).grid(
            row=0, column=2, sticky="ew", padx=(8, 0), ipady=14
        )

        # Footer (always visible)
        ttk.Separator(footer).pack(fill="x", pady=(0, 10))

        ttk.Label(
            footer,
            text="Tip: Esc closes slot/message windows",
            font=self.f(18),
            bootstyle="secondary",
            wraplength=max_title_px,
            justify="left",
        ).pack(anchor="w")

    def _on_class_selected(self):
        self.exit_requested = False

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
            m = int(self.minutes_var.get())
            s = int(self.seconds_var.get())
        except Exception:
            m, s = 0, 10
        total = max(1, (m * 60) + s)
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

        self.session_students = list(self.classes[class_name])
        self.student_grades = {}

        if not self.session_students:
            Messagebox.show_error(title="Error", message=f"No students found for {class_name}.")
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

        win = ttk.Toplevel(self.root)
        win.title("Lucky Student")
        win.attributes("-topmost", True)
        win.geometry(self._slot_window_geometry(parent_w=WINDOW_WIDTH))
        win.minsize(760, 560)

        # --- theme-safe colors ---
        colors = getattr(self.style, "colors", None)
        bg = getattr(colors, "bg", "#f8f9fa") if colors else "#f8f9fa"
        fg = getattr(colors, "fg", "#212529") if colors else "#212529"
        primary = getattr(colors, "primary", "#0d6efd") if colors else "#0d6efd"
        secondary = getattr(colors, "secondary", "#6c757d") if colors else "#6c757d"
        success = getattr(colors, "success", "#198754") if colors else "#198754"

        alive = True  # stop scheduled animation safely

        def _hex_to_rgb(h: str):
            h = (h or "").strip().lstrip("#")
            if len(h) != 6:
                raise ValueError("not hex")
            return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

        def _rgb_to_hex(rgb):
            return "#{:02x}{:02x}{:02x}".format(*rgb)

        def _mix(c1: str, c2: str, t: float) -> str:
            # mix c1 -> c2 by t (0..1); if colors aren't hex, just return c1
            try:
                r1, g1, b1 = _hex_to_rgb(c1)
                r2, g2, b2 = _hex_to_rgb(c2)
                r = int(r1 + (r2 - r1) * t)
                g = int(g1 + (g2 - g1) * t)
                b = int(b1 + (b2 - b1) * t)
                return _rgb_to_hex((r, g, b))
            except Exception:
                return c1

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

        # --- Header ---
        header = ttk.Frame(win, padding=(18, 16))
        header.pack(fill="x")

        ttk.Label(
            header,
            text=f"Class: {class_name}",
            font=self.f(26, "bold"),
            bootstyle="secondary"
        ).pack(side="left")

        remaining_label = ttk.Label(
            header,
            text=f"Students Left: {len(self.session_students)}",
            font=self.f(26, "bold"),
            bootstyle="secondary"
        )
        remaining_label.pack(side="right")

        status_label = ttk.Label(
            win,
            text="Now Choosing 现在抽选",
            font=self.f(46, "bold"),
            bootstyle="primary"
        )
        status_label.pack(pady=(10, 10))

        # --- Borderless reel canvas ---
        reel_wrap = ttk.Frame(win, padding=(18, 6))
        reel_wrap.pack(fill="x", padx=10)

        canvas_w = SLOT_PANEL_W - 80
        row_h = max(self.fs(92), 82)
        canvas_h = row_h * 3 + self.fs(34)
        cx = canvas_w // 2
        cy = canvas_h // 2

        reel = tk.Canvas(
            reel_wrap,
            width=canvas_w,
            height=canvas_h,
            bg=bg,
            highlightthickness=0,
            bd=0
        )
        reel.pack()

        pad = self.fs(14)
        radius = self.fs(18)

        # Rounded-rect helper (Canvas has no native rounded rectangle)
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
                x1, y1
            ]
            return reel.create_polygon(points, smooth=True, splinesteps=20, **kwargs)

        # Borderless fills + optional soft shadow
        card_fill = _mix(bg, fg, 0.02)
        band_fill = _mix(bg, primary, 0.06)

        shadow = _mix(bg, "#000000", 0.10)
        _round_rect(
            pad + self.fs(6), pad + self.fs(6),
            canvas_w - pad + self.fs(6), canvas_h - pad + self.fs(6),
            radius,
            fill=shadow, outline=""
        )

        _round_rect(
            pad, pad, canvas_w - pad, canvas_h - pad,
            radius,
            fill=card_fill, outline=""
        )

        band_h = int(row_h * 0.90)
        band = _round_rect(
            pad + self.fs(16), cy - band_h // 2,
            canvas_w - (pad + self.fs(16)), cy + band_h // 2,
            self.fs(16),
            fill=band_fill, outline=""
        )

        # --- Name formatting: single-line only ---
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
                    left, right = s.split(None, 1)  # safe two-part split
                    return f"{left} · {right}"
            return s

        font_cache: dict[tuple[str, int, int, str], tuple] = {}
        max_text_w = canvas_w - (pad * 2) - self.fs(90)

        def _fit_font(text: str, base_px: int, min_px: int, weight="bold") -> tuple:
            key = (text, base_px, min_px, weight)
            if key in font_cache:
                return font_cache[key]

            size = self.fs(base_px)
            min_size = self.fs(min_px)
            fnt = tkfont.Font(family=self.FONT_FAMILY, size=size, weight=weight)

            while size > min_size and fnt.measure(text) > max_text_w:
                size -= 1
                fnt.configure(size=size)

            font_cache[key] = (self.FONT_FAMILY, size, weight)
            return font_cache[key]

        # Reel state
        pool = self.session_students[:] if self.session_students else [final_student]
        prev_raw = random.choice(pool)
        cur_raw = random.choice(pool)
        next_raw = random.choice(pool)

        dim = _mix(fg, card_fill, 0.58)

        t_prev = reel.create_text(cx, cy - row_h, text=_format_name(prev_raw), fill=dim,
                                  anchor="center", justify="center")
        t_cur = reel.create_text(cx, cy, text=_format_name(cur_raw), fill=primary,
                                 anchor="center", justify="center")
        t_next = reel.create_text(cx, cy + row_h, text=_format_name(next_raw), fill=dim,
                                  anchor="center", justify="center")

        def _style_items():
            reel.itemconfig(t_prev, font=_fit_font(reel.itemcget(t_prev, "text"), 40, 22))
            reel.itemconfig(t_next, font=_fit_font(reel.itemcget(t_next, "text"), 40, 22))
            reel.itemconfig(t_cur, font=_fit_font(reel.itemcget(t_cur, "text"), 62, 32))

        _style_items()

        # --- Progress + bottom panel ---
        progress = ttk.Progressbar(
            win,
            mode="determinate",
            maximum=100,
            bootstyle="primary",
            style="Slot.Horizontal.TProgressbar",
        )
        progress.pack(fill="x", padx=25, pady=(8, 14))

        buttons = ttk.Frame(win, padding=(18, 10))
        buttons.pack(pady=(2, 10), fill="x")

        # --- Audio ---
        self.sound.play_music_loop(slot_sound)

        if duration >= 10:
            def _timeup_safe():
                if alive and win.winfo_exists():
                    self.sound.play_timeup(TIMEUP_SOUND)
            win.after(max(0, int(duration * 1000) - 200), _timeup_safe)

        # --- Animation ---
        start_time = time.time()
        last_time = start_time
        phase = 0.0

        max_rows_per_sec = 9.5
        min_rows_per_sec = 1.5

        final_mode = False
        final_roll_time = 0.55
        final_mode_deadline = None

        def _rotate_once():
            nonlocal prev_raw, cur_raw, next_raw
            prev_raw, cur_raw = cur_raw, next_raw
            next_raw = random.choice(pool)

            reel.itemconfig(t_prev, text=_format_name(prev_raw))
            reel.itemconfig(t_cur, text=_format_name(cur_raw))
            reel.itemconfig(t_next, text=_format_name(next_raw))
            _style_items()

        def _render(phase_px: float):
            reel.coords(t_prev, cx, (cy - row_h) - phase_px)
            reel.coords(t_cur, cx, cy - phase_px)
            reel.coords(t_next, cx, (cy + row_h) - phase_px)

        def _finalize():
            nonlocal cur_raw

            cur_raw = final_student
            reel.itemconfig(t_cur, text=_format_name(cur_raw))
            _style_items()

            reel.itemconfig(band, fill=_mix(bg, success, 0.08), outline="")
            reel.itemconfig(t_cur, fill=success)

            reel.itemconfig(t_prev, state="hidden")
            reel.itemconfig(t_next, state="hidden")

            status_label.config(text="Selected 选中", bootstyle="success")
            try:
                progress.pack_forget()
            except Exception:
                pass

            self.sound.stop_music()

            if final_student in self.session_students:
                self.session_students.remove(final_student)
            remaining_label.config(text=f"Students Left: {len(self.session_students)}")

            self._render_grading_controls(win, buttons, class_name, final_student)

        def _frame():
            nonlocal last_time, phase, final_mode, final_mode_deadline, next_raw

            if (not alive) or (not win.winfo_exists()):
                return

            now = time.time()
            dt = max(0.001, now - last_time)
            last_time = now
            elapsed = now - start_time

            # Enter final mode: make NEXT be final_student so the next rotation lands it in center
            if (not final_mode) and (elapsed >= duration):
                final_mode = True
                final_mode_deadline = now + final_roll_time
                next_raw = final_student
                reel.itemconfig(t_next, text=_format_name(next_raw))
                _style_items()

            if not final_mode:
                t = max(0.0, min(1.0, elapsed / max(0.001, duration)))
                rows_per_sec = min_rows_per_sec + (max_rows_per_sec - min_rows_per_sec) * ((1.0 - t) ** 2.1)
                phase += (rows_per_sec * row_h) * dt
            else:
                remaining = max(0.0, row_h - phase)
                remaining_time = max(0.05, (final_mode_deadline or now) - now)
                phase += min(remaining, (remaining / remaining_time) * dt)

            rotated = False
            while phase >= row_h - 1e-6:
                phase -= row_h
                _rotate_once()
                rotated = True

            _render(phase)

            progress["value"] = min(100, int(min(elapsed, duration) / max(0.001, duration) * 100))

            if final_mode and rotated and abs(phase) < 1e-3 and cur_raw == final_student:
                _finalize()
                return

            win.after(16, _frame)

        _render(0.0)
        _frame()

    # ---------- Grading controls ----------

    def _render_grading_controls(self, win, button_frame, class_name: str, student_name: str):
        """
        Render the grading prompt + buttons INSIDE button_frame using grid.
        Fixes clipping caused by packing additional widgets onto `win`.
        """
        for w in button_frame.winfo_children():
            w.destroy()

        button_frame.columnconfigure(0, weight=1, uniform="rate")
        button_frame.columnconfigure(1, weight=1, uniform="rate")
        button_frame.columnconfigure(2, weight=1, uniform="rate")
        button_frame.columnconfigure(3, weight=1, uniform="rate")

        prompt = ttk.Label(
            button_frame,
            text="Rate the response • 评价",
            font=self.f(30, "bold"),
            bootstyle="secondary",
            anchor="center",
            justify="center",
        )
        prompt.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, self.fs(10)))

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

        rating_styles = {"A*": "success", "A": "primary", "B": "warning", "C": "danger"}
        ratings = ("A*", "A", "B", "C")

        for i, rating in enumerate(ratings):
            ttk.Button(
                button_frame,
                text=rating,
                bootstyle=rating_styles[rating],
                command=lambda r=rating: apply_rating(r),
            ).grid(
                row=1,
                column=i,
                sticky="ew",
                padx=(0 if i == 0 else self.fs(10), 0),
                ipady=self.fs(16),
            )

    # ---------- Feedback popup ----------

    def _show_message_popup(self, title: str, message: str, class_name: str):
        msg_win = ttk.Toplevel(self.root)
        msg_win.title(title)
        msg_win.attributes("-topmost", True)
        msg_win.geometry(self._slot_window_geometry(parent_w=WINDOW_WIDTH))
        msg_win.minsize(720, 560)

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

        # Structured layout: header, message (expands), bottom buttons (always visible)
        outer = ttk.Frame(msg_win, padding=(24, 22, 24, 22))
        outer.pack(fill="both", expand=True)

        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        header = ttk.Label(
            outer,
            text="Feedback • 反馈",
            font=self.f(46, "bold"),
            bootstyle="primary",
            anchor="center",
            justify="center",
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, self.fs(14)))

        body = ttk.Label(
            outer,
            text=message,
            font=self.f(38),
            wraplength=SLOT_PANEL_W - 80,
            justify="center",
            anchor="center",
        )
        body.grid(row=1, column=0, sticky="nsew", pady=(0, self.fs(18)))

        btns = ttk.Frame(outer)
        btns.grid(row=2, column=0, sticky="ew")
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)

        def next_student():
            try:
                msg_win.destroy()
            except Exception:
                pass
            self._next_student(class_name)

        ttk.Button(
            btns,
            text="Next Student • 下一个",
            bootstyle="success",
            command=next_student
        ).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(12)), ipady=self.fs(16))

        ttk.Button(
            btns,
            text="Exit • 返回选班",
            bootstyle="secondary",
            command=exit_to_main
        ).grid(row=0, column=1, sticky="ew", padx=(self.fs(12), 0), ipady=self.fs(16))

    # ---------- Grades summary ----------

    def _show_grades_summary(self):
        self.exit_requested = False

        win = ttk.Toplevel(self.root)
        win.title("Grades Summary")
        win.attributes("-topmost", True)
        win.geometry("680x680+30+80")

        def on_escape(event=None):
            try:
                win.destroy()
            except Exception:
                pass

        win.bind("<Escape>", on_escape)

        ttk.Label(win, text="Grades Summary", font=self.f(34, "bold")).pack(pady=18)

        if not self.student_grades:
            summary_text = "No grades recorded in this session."
        else:
            lines = [f"{student}: {grade}" for student, grade in self.student_grades.items()]
            summary_text = "\n".join(lines)

        box = ttk.Text(win, height=18, wrap="word", font=self.f(20))
        box.pack(fill="both", expand=True, padx=15, pady=10)
        box.insert("1.0", summary_text)
        box.configure(state="disabled")

        ttk.Button(win, text="Close", bootstyle="secondary", command=win.destroy).pack(pady=(0, 18), ipady=14)


def main():
    style = Style(theme=THEME)
    root = style.master
    _ = InvisibleHandApp(root, style)
    root.mainloop()


if __name__ == "__main__":
    main()
