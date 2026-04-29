# Random Student Selector

`Random Student Selector` is a Windows classroom app for running attendance-aware questioning, timed thinking, random student selection, quick formative outcomes, and live lesson summaries.

It was developed by Samuel Oehler-Huang at Suzhou Foreign Language School for practical Cambridge classroom use. The app is not intended to make questioning random for its own sake. It gives teachers a simple routine for involving more learners, making thinking time visible, collecting informal evidence, and reflecting on who may need follow-up.

The app is built with Tkinter + `ttkbootstrap` and is optimized for Windows classroom setups, including high-DPI displays and multi-monitor use.

## Screenshot

![Random Student Selector main screen](assets/main-screen.png)

## Pedagogical position

This software aligns best with Cambridge teaching when it is used as part of a planned questioning and feedback routine.

Cambridge's guidance on teaching programmes emphasizes effective questioning, effective use of assessment, checking understanding repeatedly, scaffolding new learning, modelling good thinking, and adjusting teaching in response to evidence from learners. The related Cambridge professional development materials frame active learning as students thinking hard rather than passively receiving information, assessment for learning as feedback used to improve performance, and metacognition as learners planning, monitoring, evaluating, and changing their learning behaviours.

Random Student Selector supports those aims in a narrow but useful way:

- **Active learning:** every present learner remains part of the lesson conversation, not only the quickest volunteers.
- **Assessment for learning:** quick outcomes and notes from questioning give the teacher informal evidence for reteaching, extension, or targeted support.
- **Metacognition:** the summary can feed a short reflection routine: What did I understand? What strategy helped? What should I do next?
- **Differentiation:** attendance, selection history, and outcome patterns help the teacher notice who has not contributed, who needs support, and who may be ready for challenge.
- **Reflective practice:** the session record gives the teacher a small evidence base for reviewing what worked and planning the next lesson.

The app does not replace strong lesson planning, success criteria, teacher judgement, peer discussion, written work, or rich feedback. It is a classroom control tool for one recurring part of effective teaching: structured whole-class questioning with enough evidence to respond intelligently.

## A Cambridge-aligned classroom routine

A useful pattern is:

1. **Plan the question.** Choose a question that fits the learning objective and success criteria.
2. **Give thinking time.** Use the timer before taking answers so more learners can prepare.
3. **Select fairly.** Randomly select from the present roster to widen participation.
4. **Probe understanding.** Ask follow-up questions such as "Why?", "How do you know?", or "Can you build on that?"
5. **Record the response.** Mark a quick outcome without turning discussion into heavy data entry.
6. **Adapt the lesson.** Use the emerging pattern to decide whether to reteach, extend, pair students, or move on.
7. **Reflect.** Use the summary to prompt teacher reflection or a brief learner reflection.

This routine is especially useful for cold calling, retrieval practice, checking understanding after teacher explanation, reviewing previous learning, and making whole-class discussion more inclusive.

## Current functionality

- Select a class from a roster loaded from `assets/students.csv`.
- Choose a built-in timer preset: `5 sec`, `30 sec`, `1 min`, or `2 min`.
- Run a slot-style student selection window with an optional animated reel effect.
- Toggle `Sound` and `Slot Effect` from the main dock.
- Play optional intro and closing audio cues from the main screen.
- Take attendance before or during a session.
- Mark each selected student as `A*`, `A`, `B`, `C`, `No Grade`, or `Absent`.
- Remove absent students from the active class roster for the current session.
- Keep chosen and absent students excluded while the app remains open.
- Show a session summary with counts and per-student outcomes.
- Resume an active session from the summary window.

## Typical classroom flow

1. Start the app.
2. Select a class.
3. Optionally click `Attendance` to run roll call.
4. Choose a timer preset.
5. Ask the class a question and give silent thinking time.
6. Click `START SELECTION`.
7. When the selected student appears, listen to the response and probe as needed.
8. Choose one outcome:
   - `A*`, `A`, `B`, `C`
   - `No Grade`
   - `Absent`
9. Continue with `Next Student` or return to the main dock.
10. Open `View Summary` to review participation, absences, and outcome patterns.

## Attendance mode

The app includes a sequential roll-call mode.

- Attendance prefers a secondary monitor when one is available.
- The roll-call window opens full screen by default.
- Keyboard shortcuts:
  - `Enter`, `Space`, `P`, `Right Arrow`: mark present
  - `A`, `Backspace`, `Left Arrow`: mark absent
  - `F11`: toggle full screen
  - `Esc`: exit full screen, then close
- When roll call finishes, the app shows a copyable absent list.
- Absent students are removed from the day's active session roster.

## Session outcomes

Each selected student can be recorded as one of these outcomes:

- `A*`
- `A`
- `B`
- `C`
- `No Grade`
- `Absent`

After a graded outcome, the app shows a feedback message chosen from `assets/messages.csv`.

For stronger assessment-for-learning use, treat these outcomes as quick teacher evidence rather than final attainment data. The most important follow-up is the instructional decision: clarify a misconception, ask another learner to build on the response, revisit success criteria, or note a student for later support.

The summary window tracks:

- remaining students
- graded students
- no-grade students
- absent students
- counts by grade band

## Required data files

The app expects these CSV files in `assets/`.

### `assets/students.csv`

Required. This file is intentionally ignored by Git so each user can keep a local class list.

Format:

```csv
class,student
IC 1.1,Anna Chen
IC 1.1,Ben Carter
IC 1.2,Jordan Lee
```

Notes:

- Two columns are required: class name, student name.
- A header row is allowed and will be skipped automatically.
- Duplicate student names within the same class are de-duplicated while preserving order.

### `assets/messages.csv`

Required. This file supplies the feedback popup text after grading.

Format:

```csv
Rating,Message
A*,"Outstanding contribution today."
A,"Strong answer. Keep pushing."
B,"Good effort. Tighten the detail."
C,"Have another go and build the idea."
```

Notes:

- The file uses `Rating` and `Message` columns.
- Multiple messages per rating are supported.
- Messages are chosen at random from the matching rating bucket.
- Messages work best when they are task-focused and improvement-oriented.

## Audio behavior

Audio is optional.

- If `pygame` is available and the asset files exist, the app can play:
  - intro music
  - closing music
  - slot-loop audio
  - time-up sound
  - rating sounds
- If audio fails to initialize or files are missing, the app continues running without crashing.

## Running locally

Install the Python dependencies first:

```powershell
pip install ttkbootstrap pygame
```

Then run:

```powershell
python studentselector.py
```

## Build an executable

The repo includes `studentselector.spec` for PyInstaller packaging.

Example:

```powershell
pip install pyinstaller
pyinstaller --noconfirm studentselector.spec
```

The included `build.bat` also builds the app using `.venv\Scripts\python.exe`.

## Source framing

The README and landing page were rewritten around these Cambridge materials:

- Cambridge International Education, *Teaching Cambridge Programmes* section from `Teaching Cambridge Programmes.pdf`.
- [Getting started with Active Learning](https://www.cambridge-community.org.uk/professional-development/gswal/index.html)
- [Getting started with Assessment for Learning](https://cambridge-community.org.uk/professional-development/gswafl/index.html)
- [Getting started with Metacognition](https://cambridge-community.org.uk/professional-development/gswmeta/index.html)

## Notes

- The UI is designed as a small always-on-top control dock plus larger session windows.
- The app is Windows-aware for DPI scaling and taskbar-safe window placement.
- The selection and attendance windows are also kept on top for classroom use.
- Closing and relaunching the whole app starts a fresh selection and attendance history.
