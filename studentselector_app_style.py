import tkinter.font as tkfont


class AppStyleMixin:
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
                0x4E00 <= o <= 0x9FFF
                or 0x3400 <= o <= 0x4DBF
                or 0x3040 <= o <= 0x30FF
                or 0xAC00 <= o <= 0xD7AF
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

    def _effective_session_roster(self, class_name: str | None) -> list[str]:
        if not class_name:
            return []
        blocked = set(self._chosen_students_for_class(class_name))
        blocked.update(self.absent_students_by_class.get(class_name, []))
        if class_name in self.session_students_by_class:
            return [student for student in self.session_students_by_class[class_name] if student not in blocked]
        roster = list(self.classes.get(class_name, []))
        return [student for student in roster if student not in blocked]

    def _chosen_students_for_class(self, class_name: str | None) -> list[str]:
        if not class_name:
            return []

        chosen: list[str] = []

        def add(student: str) -> None:
            if student and student not in chosen:
                chosen.append(student)

        for student in getattr(self, "selected_students_by_class", {}).get(class_name, []):
            add(student)
        for student in self.student_grades_by_class.get(class_name, {}):
            add(student)
        for student in self.student_ungraded_by_class.get(class_name, []):
            add(student)

        return chosen

    def _sync_live_class_state(self, class_name: str | None = None) -> None:
        class_name = class_name or self._active_or_selected_class()
        if not class_name or class_name not in self.classes:
            return

        self.student_grades = self.student_grades_by_class.setdefault(class_name, {})
        self.student_ungraded = self.student_ungraded_by_class.setdefault(class_name, [])
        self.absent_students = self.absent_students_by_class.setdefault(class_name, [])
        effective_roster = self._effective_session_roster(class_name)
        if class_name in self.session_students_by_class:
            self.session_students_by_class[class_name] = effective_roster
        self.session_students = self.session_students_by_class.get(class_name, effective_roster)

    def _class_metrics(self, class_name: str | None = None) -> dict[str, int | str]:
        class_name = class_name or self._active_or_selected_class()
        roster = list(self.classes.get(class_name, [])) if class_name else []
        session_roster = self._effective_session_roster(class_name)
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
        return meta.get(
            rating,
            {"label": "Recorded", "bg": self.palette["panel_alt"], "fg": self.palette["text_light"]},
        )

    def _apply_visual_theme(self):
        p = self.palette
        self.root.configure(bg=p["bg"])

        self.style.configure("App.TFrame", background=p["bg"])
        self.style.configure("Panel.TFrame", background=p["panel"])
        self.style.configure("PanelAlt.TFrame", background=p["panel_alt"])
        self.style.configure("Card.TFrame", background=p["card"])
        self.style.configure(
            "DockHeroTitle.TLabel",
            background=p["panel"],
            foreground=p["text_light"],
            font=self.hf(28, "bold"),
        )
        self.style.configure(
            "DockHeroBody.TLabel",
            background=p["panel"],
            foreground=p["text_muted"],
            font=self.f(15),
        )
        self.style.configure(
            "SectionTitle.TLabel",
            background=p["card"],
            foreground=p["text_light"],
            font=self.hf(15, "bold"),
        )
        self.style.configure(
            "PanelAltTitle.TLabel",
            background=p["panel_alt"],
            foreground=p["text_light"],
            font=self.hf(15, "bold"),
        )
        self.style.configure(
            "PanelTitle.TLabel",
            background=p["panel"],
            foreground=p["text_light"],
            font=self.hf(15, "bold"),
        )
        self.style.configure(
            "Hint.TLabel",
            background=p["bg"],
            foreground=p["text_muted"],
            font=self.f(14),
        )
        self.style.configure(
            "PopupTitle.TLabel",
            background=p["bg"],
            foreground=p["text_light"],
            font=self.hf(34, "bold"),
        )
        self.style.configure(
            "PopupBody.TLabel",
            background=p["panel"],
            foreground=p["text_light"],
            font=self.f(26),
        )
        self.style.configure(
            "SummaryTitle.TLabel",
            background=p["bg"],
            foreground=p["text_light"],
            font=self.hf(26, "bold"),
        )

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
        self.style.map(
            "PrimaryAction.TButton",
            background=[("active", self._shade(p["accent"], "#ffffff", 0.08))],
        )

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
        self.style.map(
            "SecondaryAction.TButton",
            background=[("active", self._shade(p["accent_alt"], "#ffffff", 0.08))],
        )

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
        self.style.map(
            "Utility.TButton",
            background=[("active", self._shade(p["bg_alt"], "#000000", 0.03))],
        )

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
        self.style.map(
            "TimeOff.TButton",
            background=[("active", self._shade(p["field"], "#ffffff", 0.05))],
        )
        self.style.map(
            "TimeOn.TButton",
            background=[("active", self._shade(p["accent"], "#ffffff", 0.04))],
        )

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
        self.style.map(
            "Absent.TButton",
            background=[("active", self._shade(p["bg_alt"], "#000000", 0.03))],
        )
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
