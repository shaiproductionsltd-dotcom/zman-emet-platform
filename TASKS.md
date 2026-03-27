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

- [ ] Improve the wording in the waiting UI so it clearly explains that the report is being prepared
- [ ] Clean the mixed English/Hebrew text on the upload flow and success screen
- [x] Inspect one real Flamingo export and map worker sheet / summary sheet matching rules
- [x] Add the Flamingo payroll script using the new script registry
- [x] Generate payroll summary and exceptions sheets for Flamingo
- [ ] Create the final Render production service with the final public name

## Later

- [ ] Change the default admin password to a strong password
- [ ] Add protection against repeated failed login attempts
- [ ] Optimize processing speed for larger reports
- [ ] Decide whether to move long-running processing to a background job
- [ ] Add custom domain instead of relying on the default `onrender.com` address
- [ ] Add richer payroll analytics such as top earners and payroll distribution for Flamingo
