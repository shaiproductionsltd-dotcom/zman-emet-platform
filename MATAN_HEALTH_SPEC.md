# Matan Health Care Script Spec

## Goal

Build a filter-driven reporting toolset for `Matan Health Care` from exported attendance files.

The client wants a small reporting machine where they choose filters and receive focused management reports.

## Source Files Identified

### 1. Monthly detailed attendance export

File type:
- XLS

Observed structure:
- one sheet per employee
- employee name appears inside the sheet
- employee ID appears inside the sheet
- entry / exit times may include `*`

Main value for this file:
- count manual edits by detecting `*` near entry / exit values

### 2. Consolidated missing-vs-standard report

File type:
- XLS

Observed structure:
- one flat report sheet
- includes employee number, employee name, standard hours, missing hours, attendance hours and leave categories

Main value for this file:
- filter employees by missing-hours thresholds
- build management reports from missing-hours data

### 3. Organizational structure file

File type:
- CSV

Observed structure:
- employee name
- employee number (`שכר`)
- ID
- direct manager
- manager indicator
- department
- email

Main value for this file:
- build a visual organization chart
- likely export to PDF

## First Recommended Deliverables

### A. Missing Hours Filter Tool

Input:
- consolidated missing-vs-standard XLS

Output:
- filtered Excel report
- summary area at top

Suggested filters:
- minimum missing hours
- maximum missing hours
- department
- direct manager
- specific employee number / ID / name
- only employees with missing hours

Suggested output columns:
- employee number
- employee name
- ID number
- department
- direct manager
- standard hours
- attendance hours
- missing hours
- leave columns from source report

Suggested summary metrics:
- total employees in result
- total missing hours in result
- departments represented
- managers represented

### B. Manual Edits Report

Input:
- monthly detailed attendance XLS

Output:
- detailed sheet by employee with total edit count
- summary sheet ranking employees by edit count

Suggested metrics:
- total manual edits in file
- employees with highest manual edit count
- employees with zero manual edits

Detection logic:
- count `*` in entry / exit cells

### C. Organization Chart PDF

Input:
- organizational structure CSV

Output:
- styled visual hierarchy diagram
- PDF export

Main challenge:
- determine the exact parent-child rule from `מנהל ישיר` and manager markers

## What Is Already Confirmed From Inspection

### Monthly detailed report
- employee name appears in row 5
- employee ID appears in row 7
- daily section begins later in the sheet
- `*` appears next to edited entry / exit values

### Missing-vs-standard report
- employee number exists in column 1
- employee name exists in column 3
- missing hours exist in the `חוסר` column
- source is already flat and easy to filter

### Organization CSV
- includes department and direct manager data
- likely enough to build hierarchy after one matching rule is chosen

## Recommended Build Order

1. Missing Hours Filter Tool
2. Manual Edits Report
3. Organization Chart PDF

This order gives the fastest customer value with the lowest implementation risk.

## Open Questions Still Needing Decision

1. Should the missing-hours tool support only one report type at first, or multiple presets?
2. Which filters are mandatory in version 1?
3. Should department and manager data be merged into the missing-hours report by joining with the organization CSV?
4. For manual edits, do we count each `*` separately, or count a day once even if both entry and exit were edited?
5. For the organization chart, do you want:
   - full company chart
   - one department chart
   - both

