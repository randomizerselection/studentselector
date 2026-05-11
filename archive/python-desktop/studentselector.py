from ttkbootstrap import Style

from studentselector_app import InvisibleHandApp
from studentselector_config import THEME, enable_windows_high_dpi


def main():
    enable_windows_high_dpi()
    style = Style(theme=THEME)
    root = style.master
    _ = InvisibleHandApp(root, style)
    root.mainloop()


if __name__ == "__main__":
    main()
