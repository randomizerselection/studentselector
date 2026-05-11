# Random Student Selector

`Random Student Selector` is a static web app for classroom questioning. It runs on GitHub Pages with HTML, CSS, and browser JavaScript. Future development should happen in the web app files at the repository root.

The old Python/Tkinter desktop version is archived under `archive/python-desktop/` for reference only.

## Live App

- App: https://randomizerselection.github.io/studentselector/
- About page: https://randomizerselection.github.io/studentselector/about.html
- Repository: https://github.com/randomizerselection/studentselector

## What The Web App Does

- Loads classes from `assets/students.csv`.
- Loads feedback messages from `assets/messages.csv`.
- Runs timer presets: `5 sec`, `30 sec`, `1 min`, and `2 min`.
- Supports sound and slot-effect toggles.
- Runs attendance and excludes absent students for the current browser tab.
- Randomly selects students from the remaining active roster.
- Records outcomes: `A*`, `A`, `B`, `C`, `No Grade`, and `Absent`.
- Shows a session summary with counts and per-student outcomes.
- Exposes `window.StudentSelector` so lesson pages can open or mount the selector on demand.

Session state is stored in `sessionStorage`, so it persists while the browser tab is open and resets in a new tab.

## Repository Structure

```text
.
├── index.html                 # Primary GitHub Pages app entry
├── selector.css               # Web app styles
├── selector.js                # Web app logic and public browser API
├── about.html                 # Public explanation / teaching rationale
├── landing.css                # About page styles
├── landing.js                 # About page behavior
├── assets/
│   ├── students.csv           # Public roster file loaded by the browser
│   ├── messages.csv           # Feedback messages
│   ├── *.mp3                  # Optional classroom audio cues
│   └── icon.png               # App icon
├── docs/
│   └── integration.md         # How lesson pages embed the selector
├── tests/
│   └── selector.spec.js       # Playwright smoke tests
└── archive/
    └── python-desktop/        # Archived Tkinter version, not active development
```

## Active Development Files

Most future changes should touch these files:

- `index.html` for the standalone app shell.
- `selector.css` for the selector UI.
- `selector.js` for app behavior and the `window.StudentSelector` API.
- `assets/students.csv` for roster data.
- `assets/messages.csv` for feedback text.
- `tests/selector.spec.js` for Playwright coverage.

Do not add new behavior to the archived Python app unless you are deliberately reviving the desktop version.

## Data Files

### `assets/students.csv`

Required. This file is published with the GitHub Pages site so the browser app can load the roster.

```csv
class,student
IC 1.1,Anna Chen
IC 1.1,Ben Carter
IC 1.2,Jordan Lee
```

Rules:

- Use two columns: `class` and `student`.
- A header row is allowed and skipped automatically.
- Duplicate student names inside the same class are de-duplicated while preserving order.
- This file is public on GitHub Pages.

### `assets/messages.csv`

Required. This file supplies the feedback popup text after grading.

```csv
Rating,Message
A*,"Outstanding contribution today."
A,"Strong answer. Keep pushing."
B,"Good effort. Tighten the detail."
C,"Have another go and build the idea."
```

Rules:

- Use `Rating` and `Message` columns.
- Multiple messages per rating are supported.
- Messages are chosen at random from the matching rating bucket.

## Public Browser API

`selector.js` exposes:

```js
window.StudentSelector.mount(container, options)
window.StudentSelector.open(options)
```

Use `mount` when another page provides its own panel or container. Use `open` for the selector's built-in full-page overlay.

Common options:

- `basePath`: base URL for `assets/` and `selector.css`.
- `skipStyles`: set `true` when the host page supplies its own scoped selector styles.
- `onClose`: callback called when the selector overlay closes.

See [docs/integration.md](docs/integration.md) for examples.

## Running Locally

Install test dependencies:

```powershell
npm install
```

Serve the repository root with any static server:

```powershell
python -m http.server 8766 --bind 127.0.0.1
```

Open:

```text
http://127.0.0.1:8766/
```

Run Playwright tests:

```powershell
npm test
```

## GitHub Pages Deployment

The `gh-pages` branch is the published site. GitHub Pages serves the repository root, so these files must stay at the top level:

- `index.html`
- `selector.css`
- `selector.js`
- `about.html`
- `assets/`

The app has no build step. Commit and push changes to `gh-pages` to publish.

## Archived Python Version

The previous Python/Tkinter desktop implementation lives in `archive/python-desktop/`.

It is retained as implementation history and as a reference for behavior that has already been ported to the web app. It is not the active runtime for GitHub Pages, and future feature work should target the web app.

## Teaching Rationale

The app is designed for structured classroom questioning:

1. Plan the question.
2. Give visible thinking time.
3. Select from the present class.
4. Probe the student's reasoning.
5. Record a quick formative outcome.
6. Use the pattern to adapt the next teaching move.

It supports active learning, assessment for learning, inclusive participation, and teacher reflection. It does not replace lesson planning, teacher judgement, written work, or richer feedback.
