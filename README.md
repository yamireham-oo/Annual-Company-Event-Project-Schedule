# Project-Schedule
An Excel Gantt planner for a company event. Start and end dates are calculated automatically around the weekend, and a timeline shows each task as a bar.

<img width="1165" height="409" alt="Screenshot 2026-11-22 104252" src="https://github.com/user-attachments/assets/0228a96f-c163-4407-b945-b82aebc2d41e" />

<img width="1319" height="394" alt="Screenshot 2026-11-22 104321" src="https://github.com/user-attachments/assets/d45ab808-11d4-42bf-9a21-abf69fa45036" />


[Download the workbook](Project Schedule.xlsx)


What it does

- Plans 8 event tasks with owner, duration, status and % complete.
- Calculates every start and end date from one Project Start cell, skipping Friday and Saturday (the Saudi work week).
- Draws each task as a bar on a date timeline using conditional formatting, with the completed part of each bar shaded in a second colour. The first date on the timeline is linked to Project Start, so the chart moves when the start date or a duration changes.
- Summarises the plan: total tasks, tasks by status, overall % complete weighted by duration, and the project end date.
- Looks up any task from a dropdown to show its owner, status and working days left.

Formulas

Start date of the first task begins on the Project Start date, or the next working day if it falls on a weekend.

```excel
=WORKDAY.INTL($B$3-1,1,7)
```

Start date of every other task is the next working day after the previous task ends.

```excel
=WORKDAY.INTL(F13,1,7)
```

End date moves forward by the duration, counting the start day as day one.

```excel
=WORKDAY.INTL($D13,$E13-1,7)
```

Timeline: the first date is linked to Project Start, and every next date adds one day.

```excel
=B3
=K12+1
```

Timeline bars colour a cell when its date falls between a task's start and end.

```excel
=AND(K$12>=$D13,K$12<=$F13)
```

Offset shows how many days after the project start each task begins.

```excel
=D13-$B$3
```

Working days left in a task

```excel
=ROUNDUP(E13*(1-H13),0)
```

Summary block

```excel
=COUNT(A13:A20)
=COUNTIF($G$13:$G$20,"Not started")
=COUNTIF($G$13:$G$20,"In progress")
=COUNTIF($G$13:$G$20,"Completed")
=SUMPRODUCT(E13:E20,H13:H20)/SUM(E13:E20)
=MAX(F13:F20)
```

Task lookup (owner, status, working days left)

```excel
=XLOOKUP(D5,B13:B20,C13:C20,,0)
=XLOOKUP(D5,B13:B20,G13:G20,,0)
=XLOOKUP(D5,B13:B20,J13:J20,,0)
```

Skills: WORKDAY.INTL, XLOOKUP, COUNTIF, SUMPRODUCT, MAX, ROUNDUP, conditional formatting, data validation, project scheduling
