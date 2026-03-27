# Project Tasks

## Done

- [x] Fix Excel upload processing for `.xls` and `.xlsx`
- [x] Add safer upload validation and clearer error handling
- [x] Increase request timeout for long Excel processing on Render
- [x] Connect GitHub access from this computer and enable push workflow
- [x] Add PostgreSQL support with SQLite fallback
- [x] Configure the Render service to use `DATABASE_URL`
- [x] Fix PostgreSQL permissions save flow
- [x] Verify persistence after redeploy
- [x] Add a loading message on the upload page so users know processing can take a few minutes
- [x] Add a progress-style loading bar or animated waiting state during report processing
- [x] Refactor the app so each script can define its own processor and labels

## Next

- [x] Improve the wording in the waiting UI so it clearly explains that the report is being prepared
- [x] Clean the mixed English/Hebrew text on the upload flow and success screen
- [x] Inspect one real Flamingo export and map worker sheet / summary sheet matching rules
- [x] Add the Flamingo payroll script using the new script registry
- [x] Generate payroll summary and exceptions sheets for Flamingo
- [ ] Fix Matan missing-hours report regression in production
- [ ] Clean all remaining broken UI text into readable English first
- [ ] Create a staging Render service separate from production
- [ ] Define release checklist and backup workflow
- [ ] Build Matan missing-hours filter tool
- [ ] Build Matan manual-edits report from monthly detailed XLS
- [ ] Build Matan organization-chart PDF from organizational CSV
- [ ] Create the final Render production service with the final public name

## Later

- [ ] Change the default admin password to a strong password
- [ ] Add protection against repeated failed login attempts
- [ ] Optimize processing speed for larger reports
- [ ] Decide whether to move long-running processing to a background job
- [ ] Split the monolithic `app.py` into cleaner modules
- [ ] Add automated regression tests for existing scripts
- [ ] Add monitoring and error tracking
- [ ] Add customer usage analytics
- [ ] Add custom domain instead of relying on the default `onrender.com` address
- [ ] Add richer payroll analytics such as top earners and payroll distribution for Flamingo
- [ ] Add more preset management reports for Matan Health Care
