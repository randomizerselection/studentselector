import random
import time
import tkinter as tk
import tkinter.font as tkfont

import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox

from studentselector_config import (
    MEDIUM_MAX,
    RATING_SOUNDS,
    SHORT_MAX,
    SLOT_SOUND_LONG,
    SLOT_SOUND_MEDIUM,
    SLOT_SOUND_SHORT,
    TIMEUP_SOUND,
    WINDOW_WIDTH,
)


class AppSessionMixin:
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

        self.selected_students_by_class.setdefault(class_name, [])
        if class_name not in self.student_grades_by_class:
            self.student_grades_by_class[class_name] = {}
        if class_name not in self.student_ungraded_by_class:
            self.student_ungraded_by_class[class_name] = []
        if class_name not in self.absent_students_by_class:
            self.absent_students_by_class[class_name] = []

        master_roster = list(self.classes.get(class_name, []))
        if not master_roster:
            Messagebox.show_error(title="Error", message=f"No students found for {class_name}.")
            return

        # Initialize (or reinitialize) from the master class list, excluding
        # students already chosen during this app run and students marked absent.
        roster = self._effective_session_roster(class_name)
        self.session_students_by_class[class_name] = roster

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
        remaining_value = tk.Label(
            topbar,
            text=f"{len(self.session_students)} left",
            font=self.hf(13, "bold"),
            bg=panel_alt,
            fg=p["accent_alt"],
        )
        remaining_value.pack(side="left")
        timer_value = tk.Label(
            topbar,
            text=self._format_seconds(duration),
            font=self.hf(13, "bold"),
            bg=panel_alt,
            fg=primary,
        )
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

        _round_rect(
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
                    return f"{left} {right}"
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

        t_prev = reel.create_text(
            cx,
            cy - row_h,
            text=_format_name(prev_raw),
            fill=dim,
            anchor="center",
            justify="center",
        )
        t_cur = reel.create_text(cx, cy, text=_format_name(cur_raw), fill=fg, anchor="center", justify="center")
        t_next = reel.create_text(
            cx,
            cy + row_h,
            text=_format_name(next_raw),
            fill=dim,
            anchor="center",
            justify="center",
        )

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

            self._remember_selected_student(class_name, final_student)

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
            self._remember_selected_student(class_name, student_name)
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
            self._remember_selected_student(class_name, student_name)
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
            self._remember_selected_student(class_name, student_name)
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
            command=next_student,
        ).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(10)), ipady=self.fs(12))

        ttk.Button(
            btns,
            text="Return To Dock",
            style="Utility.TButton",
            command=exit_to_main,
        ).grid(row=0, column=1, sticky="ew", padx=(self.fs(10), 0), ipady=self.fs(12))

    def _show_grades_summary(self):
        self.exit_requested = False

        win = ttk.Toplevel(self.root)
        win.title("Session Summary")
        win.attributes("-topmost", True)
        win.geometry(f"{self.fs(860)}x{self.fs(760)}+{self.fs(30)}+{self.fs(80)}")
        win.minsize(self.fs(760), self.fs(680))
        win.configure(bg=self.palette["bg"])
        p = self.palette

        class_name = self._active_or_selected_class() or self.active_class or "Current Session"
        if class_name in self.classes:
            grades = self.student_grades_by_class.get(class_name, {})
            ungraded = self.student_ungraded_by_class.get(class_name, [])
            absent = self.absent_students_by_class.get(class_name, [])
            remaining = len(self._effective_session_roster(class_name))
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
        tk.Label(
            header,
            text=class_name,
            font=self.hf(24, "bold"),
            bg=p["bg"],
            fg=p["text_light"],
        ).grid(row=0, column=0, sticky="w")
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
            tk.Label(
                chip,
                text=label,
                font=self.hf(16, "bold"),
                bg=p["panel_alt"],
                fg=accent,
            ).pack(anchor="center")
            tk.Label(
                chip,
                text=str(value),
                font=self.f(13, "bold"),
                bg=p["panel_alt"],
                fg=p["text_light"],
            ).pack(anchor="center", pady=(self.fs(4), 0))

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
            tk.Label(
                empty,
                text="No session records yet.",
                font=self.hf(20, "bold"),
                bg=p["bg_alt"],
                fg=p["text_light"],
            ).pack(anchor="center", pady=(self.fs(24), 0))
        else:
            canvas = tk.Canvas(list_shell, bg=p["bg_alt"], highlightthickness=0, bd=0)
            canvas.grid(row=0, column=0, sticky="nsew")
            scroll = ttk.Scrollbar(
                list_shell,
                orient="vertical",
                command=canvas.yview,
                style="Summary.Vertical.TScrollbar",
            )
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
                tk.Label(
                    section,
                    text=title,
                    font=self.hf(14, "bold"),
                    bg=p["bg_alt"],
                    fg=p["accent_alt"],
                    anchor="w",
                ).pack(anchor="w")

            def add_row(student: str, pill_text: str, pill_bg: str, pill_fg: str):
                nonlocal row_number
                idx = row_number
                row_bg = p["panel"] if idx % 2 else p["panel_alt"]
                row = tk.Frame(rows, bg=row_bg, padx=self.fs(14), pady=self.fs(12))
                row.pack(fill="x", expand=True)
                row.grid_columnconfigure(1, weight=1)

                tk.Label(
                    row,
                    text=f"{idx:02d}",
                    font=self.mf(12, "bold"),
                    bg=row_bg,
                    fg=p["text_muted"],
                    width=4,
                    anchor="w",
                ).grid(row=0, column=0, sticky="w")
                tk.Label(
                    row,
                    text=student,
                    font=self.hf(16, "bold"),
                    bg=row_bg,
                    fg=p["text_light"],
                    anchor="w",
                ).grid(row=0, column=1, sticky="w")

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
