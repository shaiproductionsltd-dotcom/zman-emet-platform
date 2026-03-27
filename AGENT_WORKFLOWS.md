# Agent Workflows

This file defines how the internal company agents should work together on every mission.

## Core Rule

No mission starts as coding.

Every mission must move through:
1. Business definition
2. Technical definition
3. UX definition
4. Execution
5. QA
6. Release decision
7. Memory update

## Mission Flow

### 1. CEO Agent
- Receives the business goal.
- Defines why the task matters.
- Decides priority.
- Approves success definition.

Output:
- business goal
- target customer
- value statement
- urgency level

### 2. Project Manager Agent
- Converts the CEO goal into a deliverable mission.
- Breaks the mission into steps.
- Assigns owners by role.
- Defines acceptance criteria.

Output:
- mission brief
- task breakdown
- acceptance checklist
- risk list

### 3. CTO Agent
- Reviews the mission from architecture and scale perspective.
- Decides whether the task belongs in:
  - local prototype
  - staging
  - production
- Checks if the task creates technical debt or platform risk.

Output:
- technical design decision
- environment strategy
- rollback path

### 4. Data Product Agent
- Converts customer pain into useful report outputs.
- Decides what the user should actually receive:
  - raw export
  - summary sheet
  - exceptions sheet
  - management insights

Output:
- output design
- monetizable value additions

### 5. UI UX Agent
- Designs the flow before implementation.
- Defines:
  - labels
  - warnings
  - upload flow
  - waiting state
  - result screen

Output:
- interaction plan
- user-facing copy
- visual direction

### 6. Backend Agent
- Implements the logic.
- Uses the approved design and output structure.
- Avoids one-off hacks.
- Builds in reusable script patterns.

Output:
- code changes
- testable feature

### 7. Security Agent
- Reviews the feature before release.
- Checks:
  - permissions
  - file handling
  - secrets
  - password safety
  - upload validation

Output:
- security review
- blockers if any

### 8. QA Agent
- Tests with real customer sample files.
- Confirms:
  - new feature works
  - existing features did not break
  - output is understandable

Output:
- pass/fail
- bugs found
- release recommendation

### 9. DevOps Agent
- Decides where and how to deploy.
- Controls:
  - staging deploy
  - production deploy
  - environment variables
  - monitoring
  - backups

Output:
- deployment status
- operational notes

### 10. Memory Updater Agent
- Updates:
  - tasks
  - specs
  - decisions
  - known issues
  - release notes

Output:
- updated project memory

## Standard Workflow Templates

## New Customer Script

1. CEO: approve value
2. PM: define scope
3. CTO: approve architecture
4. Data Product: define output workbook/report
5. UI UX: define upload/result flow
6. Backend: build
7. QA: test on real file
8. DevOps: deploy to staging
9. QA: staging verification
10. DevOps: production deploy
11. Memory Updater: record outcome

## Regression / Bug Fix

1. PM: define exact bug and reproduction
2. CTO: identify risk level
3. Backend: patch root cause
4. QA: verify bug and regression
5. DevOps: release safely
6. Memory Updater: log bug and resolution

## Platform / Security Change

1. CEO: approve business importance
2. CTO: define platform standard
3. Security: define risk and controls
4. Backend / DevOps: implement
5. QA: verify
6. Memory Updater: record policy

## Rules For Breaking Down Missions

- If the task changes customer output, Data Product must be involved.
- If the task changes user flow, UI UX must be involved.
- If the task changes permissions or credentials, Security must be involved.
- If the task changes deployment or databases, DevOps must be involved.
- Every completed mission must update memory.

## Non-Negotiables

- No direct push to production without validation.
- No feature starts from guessing file structure.
- No customer-facing script is sold before real sample validation.
- No critical infrastructure change without rollback plan.
