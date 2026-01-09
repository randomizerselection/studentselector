import os
import csv
import random
import time
import tkinter as tk
import tkinter.font as tkfont

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

INTRO_MUSIC   = resource_path(r"assets\welcome.mp3")
CLOSING_MUSIC = resource_path(r"assets\closing.mp3")

SLOT_SOUND_SHORT  = resource_path(r"assets\select_student.mp3")
SLOT_SOUND_MEDIUM = resource_path(r"assets\medium_slot.mp3")
SLOT_SOUND_LONG   = resource_path(r"assets\long_slot.mp3")
TIMEUP_SOUND      = resource_path(r"assets\timeup.mp3")

RATING_SOUNDS = {
    "A*": resource_path(r"assets\sound_a_star.mp3"),
    "A":  resource_path(r"assets\sound_a.mp3"),
    "B":  resource_path(r"assets\sound_b.mp3"),
    "C":  resource_path(r"assets\sound_c.mp3"),
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
        """
        Visual improvements based on the screenshot:
        - Prevent header clipping (handled via _fit_font_to_width in build)
        - Ensure labelframe titles, subtitle, small helper text are readable
        - Improve button readability and consistent sizing
        - Keep footer visible (layout changes in build)
        """
        base = self.f(20)
        self.root.option_add("*Font", base)
        self.root.option_add("*Listbox.Font", base)
        self.root.option_add("*TCombobox*Listbox.Font", self.f(22))
        self.root.option_add("*Text.Font", base)

        # TTK defaults
        self.style.configure(".", font=base)
        self.style.configure("TLabel", font=base)
        self.style.configure("secondary.TLabel", font=self.f(18))

        # Labelframe titles (were small originally)
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
        """
        Visual improvements based on screenshot:
        - Title no longer clips (auto-fit + wraplength)
        - Cleaner spacing rhythm
        - Buttons aligned consistently (grid with uniform columns)
        - Footer is always visible (dedicated footer frame)
        """
        self._clear_root()

        outer = ttk.Frame(self.root, padding=(16, 16, 16, 16))
        outer.pack(fill="both", expand=True)

        # Content + footer separation so copyright doesn't get pushed off-screen
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

        def on_escape(event=None):
            on_close()
        win.bind("<Escape>", on_escape)

        def on_close():
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

        win.protocol("WM_DELETE_WINDOW", on_close)

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

        ttk.Label(
            win,
            text="Now Choosing 现在抽选",
            font=self.f(46, "bold"),
            bootstyle="primary"
        ).pack(pady=(10, 10))

        slot_label = ttk.Label(
            win,
            text="",
            font=self.f(92, "bold"),
            anchor="center",
            justify="center",
        )
        slot_label.pack(pady=(8, 12), padx=20)

        progress = ttk.Progressbar(win, mode="determinate", maximum=100, bootstyle="info-striped")
        progress.pack(fill="x", padx=25, pady=(0, 14))

        buttons = ttk.Frame(win, padding=(18, 10))
        buttons.pack(pady=(2, 10))

        self.sound.play_music_loop(slot_sound)

        if duration >= 10:
            win.after(max(0, int(duration * 1000) - 200), lambda: self.sound.play_timeup(TIMEUP_SOUND))

        start_time = time.time()

        def tick():
            elapsed = time.time() - start_time
            if elapsed < duration:
                slot_label.config(text=random.choice(self.session_students))
                progress["value"] = min(100, int((elapsed / duration) * 100))

                t = elapsed / duration
                delay = int(TICK_MIN_MS + (TICK_MAX_MS - TICK_MIN_MS) * (t ** 2))
                win.after(delay, tick)
                return

            self.sound.stop_music()
            progress["value"] = 100
            slot_label.config(text=final_student)

            if final_student in self.session_students:
                self.session_students.remove(final_student)
            remaining_label.config(text=f"Students Left: {len(self.session_students)}")

            self._render_grading_controls(win, buttons, class_name, final_student)

        tick()

    # ---------- Grading controls ----------

    def _render_grading_controls(self, win, button_frame, class_name: str, student_name: str):
        for w in button_frame.winfo_children():
            w.destroy()

        ttk.Label(
            win,
            text="Rate the response",
            font=self.f(32, "bold"),
            bootstyle="secondary"
        ).pack(pady=(10, 10))

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

        for rating in ("A*", "A", "B", "C"):
            ttk.Button(
                button_frame,
                text=rating,
                bootstyle=rating_styles[rating],
                width=6,
                command=lambda r=rating: apply_rating(r),
            ).pack(side="left", padx=12, ipady=20)

    # ---------- Feedback popup ----------

    def _show_message_popup(self, title: str, message: str, class_name: str):
        msg_win = ttk.Toplevel(self.root)
        msg_win.title(title)
        msg_win.attributes("-topmost", True)
        msg_win.geometry(self._slot_window_geometry(parent_w=WINDOW_WIDTH))

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

        ttk.Label(
            msg_win,
            text="Feedback",
            font=self.f(46, "bold"),
            bootstyle="primary"
        ).pack(pady=(18, 12))

        ttk.Label(
            msg_win,
            text=message,
            font=self.f(38),
            wraplength=SLOT_PANEL_W - 60,
            justify="center",
            anchor="center",
        ).pack(padx=25, pady=18, expand=True)

        def next_student():
            try:
                msg_win.destroy()
            except Exception:
                pass
            self._next_student(class_name)

        ttk.Button(
            msg_win,
            text="Next Student",
            bootstyle="success",
            command=next_student
        ).pack(fill="x", padx=25, pady=(0, 12), ipady=18)

        ttk.Button(
            msg_win,
            text="Exit to Class Selection",
            bootstyle="secondary",
            command=exit_to_main
        ).pack(fill="x", padx=25, pady=(0, 18), ipady=16)

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
