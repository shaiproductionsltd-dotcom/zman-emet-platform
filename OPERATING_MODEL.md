# Operating Model

## Environments

### 1. Local
- Used for fast iteration.
- Never trusted as the only validation point.
- Useful for layout and logic checks.

### 2. Staging
- Mandatory before production.
- Separate Render service.
- Separate database.
- Used to test new scripts with real customer files.

### 3. Production
- Stable customer-facing environment.
- No unfinished features.
- Only validated code should reach production.

## Release Flow

1. Define feature scope in a spec.
2. Test with real sample files.
3. Validate in local if possible.
4. Deploy to staging.
5. Test staging with real files again.
6. Only then deploy to production.

## Branch Strategy

- `main`: production-ready code only.
- `staging`: integration branch for features being validated.
- feature branches:
  - `feature/matan-missing-hours`
  - `feature/matan-manual-edits`
  - `feature/org-chart`

## Backup Rules

### Code
- Commit before every risky refactor.
- Tag stable release points.
- Keep specs and task memory in repo.

### Data
- Production database must be backed up regularly.
- Before schema changes, take a manual backup.
- Never test destructive migrations directly on production.

### Customer Files
- Uploaded files should be treated as temporary processing artifacts.
- Do not keep them longer than needed unless there is a clear retention policy.

## Quality Gates

Before pushing:
- syntax check passes
- previous scripts still run
- new script tested on real sample file
- user-facing text readable

Before production deploy:
- staging test passes
- acceptance criteria reviewed
- rollback path known

## Immediate Structural Priorities

1. Create a staging Render service.
2. Move all active development testing to staging.
3. Add automated backup plan for production database.
4. Add basic regression test coverage for existing scripts.
5. Separate user-facing strings from code logic.
