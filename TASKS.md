# Project Tasks

## Done

- [x] Fix Excel upload processing for `.xls` and `.xlsx`
- [x] Add safer upload validation and clearer error handling
- [x] Increase request timeout for long Excel processing on Render
- [x] Connect GitHub access from this computer and enable push workflow
- [x] Add PostgreSQL support with SQLite fallback
- [x] Configure the Render service to use `DATABASE_URL`
- [x] Fix PostgreSQL permissions save flow

## Next

- [ ] Verify persistence after redeploy
- [ ] Add a loading message on the upload page so users know processing can take a few minutes
- [ ] Add a progress-style loading bar or animated waiting state during report processing
- [ ] Improve the wording in the waiting UI so it clearly explains that the report is being prepared

## Later

- [ ] Change the default admin password to a strong password
- [ ] Add protection against repeated failed login attempts
- [ ] Optimize processing speed for larger reports
- [ ] Decide whether to move long-running processing to a background job
- [ ] Support multiple customer-specific scripts in the platform
