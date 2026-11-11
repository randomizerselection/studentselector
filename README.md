# Random Student Selector

`Random Student Selector` is a static web app for classroom questioning. It runs on GitHub Pages with HTML, CSS, and browser JavaScript. Future development should happen in the web app files at the repository root.

The old Python/Tkinter desktop version is archived under `archive/python-desktop/` for reference only.

## Live App

- App: https://randomizerselection.github.io/studentselector/
- About page: https://randomizerselection.github.io/studentselector/about.html
- Repository: https://github.com/randomizerselection/studentselector

## What The Web App Does

- Loads a roster CSV chosen from your computer, entirely in the browser.
- Loads feedback messages from `assets/messages.csv`.
- Runs timer presets: `5 sec`, `30 sec`, `1 min`, and `2 min`.
- Starts muted; when sound is enabled, audio cues play at 15% volume.
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
- Private roster CSV files stay outside the repository and are selected in the app.
- `assets/messages.csv` for feedback text.
- `tests/selector.spec.js` for Playwright coverage.

Do not add new behavior to the archived Python app unless you are deliberately reviving the desktop version.

## Data Files

### Private roster CSV

Choose **Load roster CSV** in the app, then select your local CSV. The browser reads the file directly; the app does not upload it or request a roster from GitHub Pages.

```csv
class,student
Demo Class,Student One
Demo Class,Student Two
Demo Class,Student Three
```

- Use two columns: `class` and `student`.
- A header row is allowed and skipped automatically.
- Duplicate names inside the same class are de-duplicated while preserving order.
- Files must be 2 MB or less.
- The roster and session state stay in this tab's `sessionStorage` when available, including after a page refresh. Close the tab when finished, especially on a shared computer. Browser session restoration may restore tab data.
- Loading a different roster resets the attendance, selections, and grades for the session.
- Keep real rosters outside this public repository. `.gitignore` excludes all CSV files except `assets/messages.csv`, plus the `private-rosters/` directory, but these rules can be bypassed by force-adding files or uploading them through GitHub's website.
- Before pulling the commit that removes `assets/students.csv`, copy your local CSV outside the repository. A pull can remove a previously tracked local file.
- Removing the current file does not erase previous commits, forks, or downloaded copies. History cleanup is a separate step.

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

