import ctypes
import os

from studentselector_config import (
    SLOT_PANEL_H,
    SLOT_PANEL_MARGIN_RIGHT,
    SLOT_PANEL_MIN_H,
    SLOT_PANEL_W,
    SLOT_PANEL_Y,
    TOP_PADDING_Y,
    TOP_RIGHT_PADDING_X,
    WINDOW_WIDTH,
)


class AppWindowMixin:
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

    def _monitor_work_areas(self) -> list[tuple[int, int, int, int, bool]]:
        """
        Returns one entry per monitor as left, top, width, height, is_primary.
        Falls back to the primary desktop work area if monitor enumeration fails.
        """
        monitors: list[tuple[int, int, int, int, bool]] = []
        try:
            if os.name == "nt":
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

                enum_proc = ctypes.WINFUNCTYPE(
                    ctypes.c_int,
                    ctypes.c_void_p,
                    ctypes.c_void_p,
                    ctypes.POINTER(RECT),
                    ctypes.c_longlong,
                )

                def _collect_monitor(monitor, _hdc, _rect, _lparam):
                    info = MONITORINFO()
                    info.cbSize = ctypes.sizeof(MONITORINFO)
                    ok = ctypes.windll.user32.GetMonitorInfoW(monitor, ctypes.byref(info))
                    if ok:
                        work = info.rcWork
                        monitors.append(
                            (
                                work.left,
                                work.top,
                                work.right - work.left,
                                work.bottom - work.top,
                                bool(info.dwFlags & 1),
                            )
                        )
                    return 1

                ctypes.windll.user32.EnumDisplayMonitors(
                    0,
                    0,
                    enum_proc(_collect_monitor),
                    0,
                )
        except Exception:
            monitors = []

        if monitors:
            return monitors

        left, top, width, height = self._desktop_work_area()
        return [(left, top, width, height, True)]

    def _secondary_monitor_work_area(self) -> tuple[int, int, int, int]:
        """
        Prefer the first non-primary monitor so attendance can be projected on
        screen 2 while the control dock stays on screen 1.
        """
        monitors = self._monitor_work_areas()
        for left, top, width, height, is_primary in monitors:
            if not is_primary:
                return left, top, width, height
        left, top, width, height, _ = monitors[0]
        return left, top, width, height

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

    def _slot_window_height(
        self,
        work_area: tuple[int, int, int, int] | None = None,
    ) -> int:
        _, work_top, _, work_h = work_area or self._desktop_work_area()
        available_h = max(
            self.fs(480),
            work_h - max(self.fs(SLOT_PANEL_Y) - work_top, 0) - self.fs(24),
        )
        return min(self.fs(SLOT_PANEL_H), max(self.fs(SLOT_PANEL_MIN_H), available_h))

    def _slot_window_size(
        self,
        work_area: tuple[int, int, int, int] | None = None,
    ) -> tuple[int, int]:
        work_left, work_top, work_w, work_h = work_area or self._desktop_work_area()
        available_h = max(
            self.fs(480),
            work_h - max(self.fs(SLOT_PANEL_Y) - work_top, 0) - self.fs(24),
        )
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
