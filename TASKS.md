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

- [ ] **Messaging V1 — Phase 1 (Contacts + Lists CRUD)** live on staging behind `ENABLE_MESSAGING=1` (2026-04-26). Routes under `/messaging`, scoped per `user_id`. Production unaffected. Brief: `MISSIONS/messaging_and_forms_v1.md`. Packet: `MISSIONS/messaging_v1_packet_01.md`.
- [ ] **Messaging V1 — Phase 2 (Templates CRUD + preview)** live on staging behind `ENABLE_MESSAGING=1` (2026-04-26). Variables locked: `{{first_name}}`, `{{last_name}}`, `{{employee_number}}`, `{{phone}}`. Save blocked on unknown placeholders. Packet: `MISSIONS/messaging_v1_packet_02.md`.
- [ ] **Messaging V1 — Phase 3 (Provider abstraction, backend-only)** live on staging behind `ENABLE_MESSAGING=1` (2026-04-26). `MessagingProvider` interface + `MockProvider` + 4 NotImplementedError stubs (Inforu/019/Twilio/Meta) + `dispatch_send_message()` + `msg_logs` table. ZERO routes, ZERO UI, no real SMS, no API keys. Packet: `MISSIONS/messaging_v1_packet_03.md`.
- [ ] **Messaging V1 — Phase 4a (Send Safety, mock-only)** live on staging behind `ENABLE_MESSAGING=1` (2026-04-26). Quota (`msg_quota` 50/500 placeholders, manual top-up via helper), suppression (`msg_suppression`, manual add/remove), two-step confirmation (`msg_send_confirmations`, single-use token, 5-min expiry), atomic quota increment with decrement-on-failure, history view. All sends route through `MockProvider`. Packet: `MISSIONS/messaging_v1_packet_04.md`.
- [ ] **Messaging V1 — Phase 4b (019 test endpoint adapter)** implemented locally and ready for staging QA (2026-04-26). `OneNineProvider` real implementation against `https://019sms.co.il/api/test` ONLY (per docs: "During development you may send requests to..."). NO real-mode branch, NO `ONENINE_MODE`, NO production URL anywhere in the code. Default provider stays `mock`; 4a UI unchanged. 019 reachable only via explicit `dispatch_send_message(... provider_name="019")` from Python. Phase 5 (separate packet) will introduce the production URL with separate approval. Packet: `MISSIONS/messaging_v1_packet_04b.md`.
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
