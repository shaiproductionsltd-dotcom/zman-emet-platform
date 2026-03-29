# Project Tasks

## Current Product Priorities

- [ ] Stabilize UI language handling and add Hebrew/English switch for the main user flow
- [ ] Formalize and document the working environment strategy: local for development, `zman-emet-platform` for cloud test, `scriptly-platform` for production
- [ ] Work with temporary shared-database mode safely until a separate production DB is financially justified
- [ ] Enforce restore points before every critical shared-environment change
- [ ] Define the platform data model for users, customer accounts, script assignments, favourites, trials, and tickets
- [ ] Define the Tranzila billing integration approach and license-state transitions
- [ ] Build three additional customer-ready scripts before the full UI/admin expansion
- [ ] Build the new homepage, lead form, login, and customer ticket intake flow
- [ ] Build the admin panel for tickets, user management, script assignment, and trial/license visibility
- [ ] Build user-facing script catalogue with favourites and per-user access controls
- [ ] Add 30-day trial enforcement and payment reminder flow
- [ ] Add email and admin notifications for all incoming tickets
- [ ] Create durable memory/reporting flow so approved decisions are always recorded

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
- [x] Define initial multi-agent operating model documents for internal delivery
- [x] Establish `zman-emet-platform` as cloud test/staging and `scriptly-platform` as the active production environment
- [x] Deliver Matan manual corrections report
- [x] Deliver Rimon home office summary report
- [x] Deliver organizational hierarchy report with Excel, PowerPoint, and ZIP output modes
- [x] Test these completed reports successfully on production at `https://script-ly.com/dashboard`

## Next

- [ ] Replace broken/missing Hebrew in the main app without causing another frontend regression
- [ ] Add a safe translation layer for user-facing UI text
- [ ] Add a Hebrew/English language switcher for login, dashboard, and script screens
- [x] Improve the wording in the waiting UI so it clearly explains that the report is being prepared
- [x] Clean the mixed English/Hebrew text on the upload flow and success screen
- [x] Inspect one real Flamingo export and map worker sheet / summary sheet matching rules
- [x] Add the Flamingo payroll script using the new script registry
- [x] Generate payroll summary and exceptions sheets for Flamingo
- [ ] Fix Matan missing-hours report regression in production
- [ ] Clean all remaining broken UI text into readable English first
- [ ] Define release checklist and backup workflow
- [ ] Build Matan missing-hours filter tool
- [ ] Preserve and extend org-report V1 with smart field mapping plus user confirmation
- [ ] Add admin Hebrew localization / language switch

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
- [ ] Add advanced billing analytics and churn alerts
- [ ] Add role-based admin permissions
