import ctypes
import os
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


