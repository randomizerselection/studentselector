import os
import tkinter as tk

import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox


class AppAttendanceMixin:
    def _show_attendance_dialog(self):
        selected = (self.selected_class.get() or "").strip()
        if selected == "Select a Class" or selected not in self.classes:
            Messagebox.show_error(
                title="Error",
                message="Please select a valid class before taking attendance.",
            )
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

        tk.Label(
            outer,
            text=f"Mark absent students for {class_name}",
            font=self.hf(16, "bold"),
            bg=p["bg"],
            fg=p["text_light"],
        ).pack(anchor="w", pady=(0, self.fs(8)))

        # Scrollable area for student checkboxes
        list_frame = tk.Frame(outer, bg=p["panel"], padx=self.fs(8), pady=self.fs(8))
        list_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_frame, bg=p["panel"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=p["panel"])

        scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
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

            self._sync_live_class_state(class_name)
            win.destroy()
            self._build_main_screen()

        def cancel():
            win.destroy()

        ttk.Button(btns, text="Save", style="PrimaryAction.TButton", command=save).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, self.fs(8)),
        )
        ttk.Button(btns, text="Cancel", style="Utility.TButton", command=cancel).grid(
            row=0,
            column=1,
            sticky="ew",
        )

    def _take_attendance_sequential(self):
        selected = (self.selected_class.get() or "").strip()
        if selected == "Select a Class" or selected not in self.classes:
            Messagebox.show_error(
                title="Error",
                message="Please select a valid class before taking attendance.",
            )
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
        roster_total = len(roster)
        target_work_left, target_work_top, target_work_w, target_work_h = self._secondary_monitor_work_area()
        normal_w = min(target_work_w, max(self.fs(720), int(target_work_w * 0.72)))
        normal_h = min(target_work_h, max(self.fs(420), int(target_work_h * 0.62)))
        normal_x = target_work_left + max(0, (target_work_w - normal_w) // 2)
        normal_y = target_work_top + max(0, (target_work_h - normal_h) // 2)
        normal_geometry = f"{normal_w}x{normal_h}+{normal_x}+{normal_y}"
        win.geometry(normal_geometry)

        self.style.configure(
            "AttendancePresent.TButton",
            font=self.hf(24, "bold"),
            foreground=p["text_dark"],
            background=p["bg_alt"],
            bordercolor=p["line"],
            darkcolor=self._shade(p["bg_alt"], "#000000", 0.08),
            lightcolor=self._shade(p["bg_alt"], "#ffffff", 0.04),
            focusthickness=0,
            padding=self.fs(18),
        )
        self.style.map(
            "AttendancePresent.TButton",
            background=[("active", self._shade(p["bg_alt"], "#000000", 0.03))],
        )
        self.style.configure(
            "AttendanceAbsent.TButton",
            font=self.hf(24, "bold"),
            foreground="#fffaf3",
            background=p["accent"],
            bordercolor=p["accent"],
            darkcolor=self._shade(p["accent"], "#000000", 0.14),
            lightcolor=self._shade(p["accent"], "#ffffff", 0.08),
            focusthickness=0,
            padding=self.fs(18),
        )
        self.style.map(
            "AttendanceAbsent.TButton",
            background=[("active", self._shade(p["accent"], "#ffffff", 0.08))],
        )

        # State
        idx = 0
        absent: list[str] = []
        is_fullscreen = False

        def _set_fullscreen(enabled: bool):
            nonlocal is_fullscreen
            enabled = bool(enabled)
            if os.name == "nt":
                try:
                    win.state("normal")
                except Exception:
                    pass
                try:
                    if enabled:
                        win.geometry(f"{target_work_w}x{target_work_h}+{target_work_left}+{target_work_top}")
                        win.update_idletasks()
                        win.state("zoomed")
                    else:
                        win.geometry(normal_geometry)
                    is_fullscreen = enabled
                    return
                except Exception:
                    pass

            try:
                win.attributes("-fullscreen", enabled)
                is_fullscreen = enabled
                return
            except Exception:
                pass

            try:
                win.state("zoomed" if enabled else "normal")
                is_fullscreen = enabled
            except Exception:
                is_fullscreen = False

        def _toggle_fullscreen(event=None):
            _set_fullscreen(not is_fullscreen)
            return "break"

        def _close_window(event=None):
            try:
                win.destroy()
            except Exception:
                pass
            return "break"

        def _handle_escape(event=None):
            if is_fullscreen:
                _set_fullscreen(False)
                return "break"
            return _close_window()

        title = tk.Label(
            win,
            text=class_name,
            font=self.hf(28, "bold"),
            bg=p["bg"],
            fg=p["text_light"],
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        title.pack(anchor="n", pady=(self.fs(20), 0))

        progress_label = tk.Label(
            win,
            text="",
            font=self.hf(18, "bold"),
            bg=p["bg"],
            fg=p["accent_alt"],
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        progress_label.pack(anchor="n", pady=(self.fs(8), 0))

        name_frame = tk.Frame(win, bg=p["panel"], padx=self.fs(28), pady=self.fs(28))
        name_frame.pack(fill="both", expand=True, padx=self.fs(24), pady=self.fs(20))

        name_label = tk.Label(
            name_frame,
            text="",
            font=self.hf(72, "bold"),
            bg=p["panel"],
            fg=p["text_dark"],
            wraplength=max(self.fs(680), target_work_w - self.fs(220)),
            justify="center",
            bd=0,
            relief="flat",
            highlightthickness=0,
        )
        name_label.pack(expand=True)

        info_label = tk.Label(
            win,
            text="Present: Enter, Space, P, Right Arrow    Absent: A, Backspace, Left Arrow    F11: Full Screen",
            font=self.f(16),
            bg=p["bg"],
            fg=p["text_muted"],
            bd=0,
            relief="flat",
            highlightthickness=0,
            justify="center",
            wraplength=max(self.fs(760), target_work_w - self.fs(180)),
        )
        info_label.pack(padx=self.fs(20))

        btn_frame = tk.Frame(win, bg=p["bg"])
        btn_frame.pack(fill="x", padx=self.fs(24), pady=(self.fs(16), self.fs(24)))
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        def _update_display():
            nonlocal idx
            if idx >= len(roster):
                _finish()
                return
            name = roster[idx]
            progress_label.config(text=f"Student {idx + 1} of {roster_total}")
            name_label.config(text=name)

        def _mark_present(event=None):
            nonlocal idx
            if idx >= len(roster):
                return "break"
            idx += 1
            _update_display()
            return "break"

        def _mark_absent(event=None):
            nonlocal idx
            if idx >= len(roster):
                return "break"
            absent.append(roster[idx])
            idx += 1
            _update_display()
            return "break"

        def _finish():
            # Save absent list
            self.absent_students_by_class[class_name] = absent

            # Rebuild the session roster from the master class list minus absentees
            master = list(self.classes.get(class_name, []))
            self.session_students_by_class[class_name] = [s for s in master if s not in absent]
            self._sync_live_class_state(class_name)

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
            result_w = min(target_work_w, max(self.fs(640), int(target_work_w * 0.5)))
            result_h = min(target_work_h, max(self.fs(420), int(target_work_h * 0.5)))
            result_x = target_work_left + max(0, (target_work_w - result_w) // 2)
            result_y = target_work_top + max(0, (target_work_h - result_h) // 2)
            res_win.geometry(f"{result_w}x{result_h}+{result_x}+{result_y}")
            res_win.configure(bg=self.palette["bg"])

            p2 = self.palette
            outer = tk.Frame(res_win, bg=p2["bg"], padx=self.fs(12), pady=self.fs(12))
            outer.pack(fill="both", expand=True)
            tk.Label(
                outer,
                text=f"Absent students ({len(absent)})",
                font=self.hf(16, "bold"),
                bg=p2["bg"],
                fg=p2["text_light"],
            ).pack(anchor="w")

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
            ttk.Button(
                btns,
                text="Copy Absent List",
                style="PrimaryAction.TButton",
                command=_copy,
            ).pack(side="left", expand=True, fill="x", padx=(0, self.fs(8)))
            ttk.Button(btns, text="Close", style="Utility.TButton", command=_close_all).pack(
                side="left",
                expand=True,
                fill="x",
            )
            self._grow_window_to_fit(res_win, min_bottom_margin=self.fs(24))

        win.bind("<Escape>", _handle_escape)
        win.bind("<F11>", _toggle_fullscreen)
        for seq in ("<Return>", "<KP_Enter>", "<space>", "<Right>", "<p>", "<P>"):
            win.bind(seq, _mark_present)
        for seq in ("<BackSpace>", "<Left>", "<a>", "<A>"):
            win.bind(seq, _mark_absent)
        win.protocol("WM_DELETE_WINDOW", _close_window)

        ttk.Button(
            btn_frame,
            text="Present",
            style="AttendancePresent.TButton",
            command=_mark_present,
        ).grid(row=0, column=0, sticky="ew", padx=(0, self.fs(10)), ipady=self.fs(12))
        ttk.Button(
            btn_frame,
            text="Absent",
            style="AttendanceAbsent.TButton",
            command=_mark_absent,
        ).grid(row=0, column=1, sticky="ew", ipady=self.fs(12))

        # start
        _set_fullscreen(True)
        try:
            win.focus_force()
        except Exception:
            pass
        _update_display()
