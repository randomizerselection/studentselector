import tkinter as tk

import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox

from studentselector_config import CLOSING_MUSIC, INTRO_MUSIC


class AppMainMixin:
    def _on_sound_toggle(self):
        self.sound.set_enabled(self.sound_enabled_var.get())

    def _on_slot_effect_toggle(self):
        # Intentionally minimal; toggle is read when the slot window starts.
        _ = self.slot_effect_enabled_var.get()

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
                if has_class
                else "Select a class to start"
            ),
            font=self.f(11 if compact else 12),
            bg=p["bg"],
            fg=p["text_muted"],
            anchor="w",
        ).pack(anchor="w", pady=(d(2), 0))

        panel = tk.Frame(content, bg=p["panel"], padx=panel_pad, pady=panel_pad)
        panel.grid(row=1, column=0, sticky="ew")
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            panel,
            text="Class",
            font=self.hf(13, "bold"),
            bg=p["panel"],
            fg=p["text_light"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")
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
            if has_class
            else "No class selected"
        )
        tk.Label(
            panel,
            text=roster_note,
            font=self.f(11 if compact else 12),
            bg=p["panel"],
            fg=p["text_muted"],
            anchor="w",
        ).grid(row=2, column=0, sticky="w", pady=(0, d(8)))

        tk.Label(
            panel,
            text="Timer",
            font=self.hf(13, "bold"),
            bg=p["panel"],
            fg=p["text_light"],
            anchor="w",
        ).grid(row=3, column=0, sticky="w")
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

        def toggle_tile(col: int, title: str, variable, command):
            tile = tk.Frame(
                toggles,
                bg=p["bg_alt"],
                padx=d(8),
                pady=d(6),
            )
            tile.grid(
                row=0,
                column=col,
                sticky="ew",
                padx=(0, card_gap) if col == 0 else (card_gap, 0),
            )
            text = tk.Frame(tile, bg=p["bg_alt"])
            text.pack(side="left", fill="both", expand=True)
            tk.Label(
                text,
                text=title,
                font=self.hf(12 if compact else 13, "bold"),
                bg=p["bg_alt"],
                fg=p["text_light"],
                anchor="w",
            ).pack(anchor="w")
            ttk.Checkbutton(
                tile,
                text="",
                variable=variable,
                command=command,
                bootstyle="round-toggle",
            ).pack(side="right", padx=(d(6), 0))

        toggle_tile(0, "Sound", self.sound_enabled_var, self._on_sound_toggle)
        toggle_tile(1, "Slot Effect", self.slot_effect_enabled_var, self._on_slot_effect_toggle)

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

    def _play_intro(self):
        self.exit_requested = False
        self.sound.play_music_once(INTRO_MUSIC)

    def _play_closing(self):
        self.exit_requested = False
        self.sound.play_music_once(CLOSING_MUSIC)
