import tkinter as tk

import ttkbootstrap as ttk
from ttkbootstrap import Style

from studentselector_app_attendance import AppAttendanceMixin
from studentselector_app_main import AppMainMixin
from studentselector_app_session import AppSessionMixin
from studentselector_app_style import AppStyleMixin
from studentselector_app_window import AppWindowMixin
from studentselector_config import MESSAGES_CSV, PALETTE, STUDENTS_CSV
from studentselector_services import SoundManager, load_messages_by_rating, load_students_by_class


class InvisibleHandApp(
    AppStyleMixin,
    AppWindowMixin,
    AppMainMixin,
    AppAttendanceMixin,
    AppSessionMixin,
):
    def __init__(self, root: tk.Tk, style: Style):
        self.root = root
        self.style = style
        self.sound = SoundManager()
        self.palette = dict(PALETTE)

        # Typography / classroom projection
        self.FONT_FAMILY = self._pick_font_family("Aptos", "Segoe UI", "Arial")
        self.HEADING_FONT_FAMILY = self._pick_font_family(
            "Bahnschrift SemiBold",
            "Bahnschrift",
            "Aptos Display",
            self.FONT_FAMILY,
        )
        self.CJK_FONT_FAMILY = self._pick_font_family(
            "Microsoft YaHei UI",
            "Microsoft YaHei",
            "Microsoft JhengHei UI",
            "Yu Gothic UI",
            self.FONT_FAMILY,
        )
        self.MONO_FONT_FAMILY = self._pick_font_family(
            "Consolas",
            "Cascadia Mono",
            "Courier New",
            self.FONT_FAMILY,
        )
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
        self.selected_students_by_class: dict[str, list[str]] = {}
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

    def _remember_selected_student(self, class_name: str, student_name: str) -> None:
        if class_name not in self.classes or student_name not in self.classes[class_name]:
            return

        selected_students = self.selected_students_by_class.setdefault(class_name, [])
        if student_name not in selected_students:
            selected_students.append(student_name)
