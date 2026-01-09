# Random Student Selector

## What this app does

This app helps a teacher pick students at random from a class list. You choose a class and set a timer, then press Start. The app spins through names, stops on one student, and shows the result clearly on screen. After a student is picked, you can give a quick grade (A*, A, B, or C). The app then shows a short feedback message and moves on to the next student until everyone has been chosen.

In short: it is a classroom-friendly way to call on students fairly and keep simple participation notes.

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

2) `assets/messages.csv`
- Purpose: feedback messages shown after you give a rating.
- Format: two columns with a header row: `Rating` and `Message`.
- Ratings used in the app: `A*`, `A`, `B`, `C`.
- Example:
```csv
Rating,Message
A*,Outstanding work today!
A,Great job, keep it up.
B,Good effort, aim a little higher.
C,Let's focus and try again next time.
```
