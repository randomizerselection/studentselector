# Archived Python Desktop Version

This folder contains the previous Python/Tkinter implementation of Random Student Selector.

It is preserved for historical reference only. The active app is now the static web version at the repository root:

- `index.html`
- `selector.css`
- `selector.js`
- `assets/`

Future feature work should happen in the web app. Do not treat this archive as the source of truth unless you are deliberately recovering old desktop behavior for comparison.

## Contents

- `studentselector.py`: old desktop entry point.
- `studentselector_app*.py`: Tkinter application modules.
- `studentselector_config.py`: desktop constants and paths.
- `studentselector_services.py`: CSV/audio helper services.
- `studentselector.spec`: old PyInstaller spec.
- `build.bat` and `Run (debug).bat`: old local Windows helpers, ignored by Git.

## Notes

The files were moved out of the repository root so GitHub Pages and future maintenance are centered on the browser app.

The archived code may need path adjustments if someone tries to run it from this folder, because it originally lived at the repository root beside `assets/`.
