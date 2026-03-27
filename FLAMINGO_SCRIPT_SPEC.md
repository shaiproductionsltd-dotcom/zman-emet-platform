# Flamingo Payroll Script Spec

## Goal

Build a new platform script for the client `Flamingo` that calculates each worker's salary from the monthly attendance export.

The attendance export contains:
- worker sheets with daily attendance rows
- a summary sheet that appears after each worker sheet
- a value in the `הערות` field that should contain the worker's hourly rate

The script should produce a management-ready payroll summary file and clearly flag workers who cannot be calculated.

## Core Business Rule

For each worker:

`payable_hours * hourly_rate = calculated_salary`

Where:
- `hourly_rate` comes from the `הערות` field
- `payable_hours` comes from the worker summary section in the exported report

## Version 1 Scope

Version 1 should be reliable and explainable, not overly ambitious.

### Required output

1. A new summary sheet at the start of the workbook:
   - sheet name: `Payroll Summary`
   - one row per worker

2. A separate exceptions sheet:
   - sheet name: `Requires Attention`
   - only workers with missing or invalid payroll data

3. Keep the original worker/summarized sheets in the workbook unless later decided otherwise.

## Payroll Summary Sheet

Each row should contain:
- worker name
- department, if available
- source sheet name
- matched summary sheet name
- hourly rate
- payable hours
- calculated salary
- status
- notes

### Status values

- `OK`
- `Missing hourly rate`
- `Invalid hourly rate`
- `Missing payable hours`
- `Could not match summary sheet`
- `Calculation skipped`

## Summary Metrics Area

At the top of `Payroll Summary`, include:
- total workers detected
- workers calculated successfully
- workers missing hourly rate
- workers with matching problems
- total payable hours
- total payroll amount
- average hourly rate for calculated workers

## Requires Attention Sheet

Include one row for every worker that needs manual action:
- worker name
- issue type
- details
- recommended action

### Recommended action examples

- update the hourly rate in `הערות` and export again
- verify the summary tab that belongs to this worker
- verify that payable hours exist in the exported report

## Optional Nice-to-Have Outputs

These are not mandatory for the first working version, but they add clear value:
- department summary totals
- top 10 highest payroll workers
- highlighted warning box when unresolved workers exist

## Main Technical Risk

The report structure is ambiguous:
- the tab after a worker sheet appears to belong to that same worker
- the summary tab name may not clearly match the worker name

Because of that, the script must verify how the export is structured using a real sample file before implementation.

## Required Discovery Before Coding

We need one real sample export and must answer:

1. Where exactly is the worker name stored on the worker sheet?
2. Where exactly is the `הערות` field located?
3. Where exactly are the payable hours located on the summary sheet?
4. How do we reliably match a worker sheet to its summary sheet?
5. Are the payable hours trustworthy enough to use directly, or do they need recalculation?

## Implementation Strategy

### Phase 1

- inspect one real sample workbook
- map the layout
- define matching rules between worker sheet and summary sheet

### Phase 2

- add a new script entry to the platform registry
- implement workbook parsing
- calculate payroll rows
- generate the summary and exception sheets

### Phase 3

- polish the Excel output
- add warning highlights
- test on multiple real files

## Definition of Done

The script is considered ready for customer testing when:
- it processes one real Flamingo workbook end-to-end
- it calculates salary for workers with valid hourly rate data
- it flags workers with missing rate or matching problems
- it creates a clear payroll summary sheet
- it creates a clear exceptions sheet
