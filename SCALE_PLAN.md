# Scale Plan To 1000 Users

## Target
- Reach 1000 users by the end of the year on a stable multi-script HR and accounting platform.

## Business Model Direction
- Base subscription: 150 NIS per month.
- Premium scripts: additional monthly charge.
- Product focus:
  - HR control reports
  - payroll preparation
  - management summaries
  - accounting helper automations

## Product Principles
- Each script must solve a real repetitive pain.
- Outputs must look professional enough to justify monthly payment.
- Management summaries are more valuable than raw cleaned files.

## Platform Priorities

### Phase 1. Stabilise Foundation
- staging environment
- production hygiene
- backups
- text/encoding cleanup
- release process
- basic observability

### Phase 2. Productise Existing Work
- attendance cleanup
- Flamingo payroll
- Matan missing-hours
- Matan manual edits
- organisation chart export

### Phase 3. Operational Scale
- login security hardening
- audit logging
- usage tracking by customer and script
- billing readiness
- customer management workflow

### Phase 4. Technical Scale
- queue long-running jobs instead of holding web requests open
- background workers
- persistent job status
- richer admin analytics

## Metrics To Track
- active customers
- active users
- scripts run per month
- most-used scripts
- failed runs
- average processing time
- churn risk by customer

## Current Strategic Risks
- single-file app architecture is still too monolithic
- user-facing text encoding is unstable
- no staging environment yet
- no automated regression suite yet
- no proper job queue for long-running processing
