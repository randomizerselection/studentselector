# Random Student Selector

## What this app does

This app helps a teacher pick students at random from a class list. You choose a class and set a timer, then press Start. The app spins through names, stops on one student, and shows the result clearly on screen. After a student is picked, you can give a quick grade (A*, A, B, or C). The app then shows a short feedback message and moves on to the next student until everyone has been chosen.

In short: it is a classroom-friendly way to call on students fairly and keep simple participation notes.

## Pedagogical rationale

This tool is designed to increase engagement through structured cold calling. Randomized selection has been shown to be effective for participation and reduces selection bias, promoting equitable participation. The countdown "slot" effect prompts students to prepare and discuss while time runs, increasing readiness and reducing off-task slack. The rating-and-feedback step provides immediate, low-stakes formative assessment.

## How to use it

1) Open the app.
2) Pick a class from the dropdown.
3) Set a timer (minutes and seconds).
4) Click Start Random Pick.
5) When a student is selected, choose a rating (A*, A, B, or C).
6) Read the feedback message, then choose Next Student or Exit.

Tip: Press Esc to close the slot or feedback windows.

## CSV files you need

The app reads two CSV files from the `assets` folder. You can edit them in Excel, Google Sheets, or any text editor and save as CSV.

1) `assets/students.csv`
- Purpose: the class list.
- Format: two columns per row: class name, student name.
- A header row is optional. If you include one, use something like `class,student` or `class_name,student_name`.
- Example:
```csv
class,student
IC 1.1,Anna Chen
IC 1.1,张伟
IC 1.2,Jordan Lee
```
Privacy note: this file is ignored by Git, so each user keeps their own local list.

2) `assets/messages.csv`
- Purpose: feedback messages shown after you give a rating.
- Format: two columns with a header row: `Rating` and `Message`.
- Ratings used in the app: `A*`, `A`, `B`, `C`.
- Example:
```csv
Rating,Message
A*,"Outstanding work today!"
A,"Great job, keep it up."
B,"Good effort, aim a little higher."
C,"Let's focus and try again next time."
```

## What happens during a pick

- The app cycles through names for the amount of time you set, then lands on a final student.
- The selected student is removed from that class list for the rest of the session.
- At the end, you can view a simple Grades Summary of who received which rating.

## Sounds (optional)

The app plays short sounds and music during selection and after rating. If audio files are missing, the app still runs. You can toggle sound on/off from the main screen.
