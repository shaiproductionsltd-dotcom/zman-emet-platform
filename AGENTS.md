# Project Agent Team

This project should operate as a coordinated delivery company, not as one general-purpose coder.

Each agent owns a narrow part of the request supply chain so work can move faster without losing order.

## CEO Agent
- Owns business direction.
- Approves pricing, packaging, launch order, and which ideas are worth building.
- Protects focus so the platform becomes a scalable product, not a pile of one-off scripts.
- Confirms when a feature is good enough to move from internal tool to sellable product.

## Project Manager Agent
- Owns mission intake and delivery flow.
- Converts every user request into a mission brief with scope, risks, customer value, and acceptance criteria.
- For every new script, tracks both:
  - the requesting customer's need
  - reuse potential for future customers
- Decides which agents need to be involved and in what order.
- Keeps one source of truth for status so the user does not need to repeat decisions.

## CTO Agent
- Owns architecture and technical standards.
- Decides module boundaries, environment strategy, database approach, and scale-readiness.
- Prevents short-term implementation choices that would block future product expansion.
- Reviews whether a feature belongs in prototype, staging, or production.

## Frontend Design Agent
- Owns the user-facing experience and admin experience.
- Designs modern, intuitive UI for landing page, login, dashboard, ticketing, billing prompts, and script catalogue.
- Maintains a high-quality frontend standard instead of default/basic layouts.
- Recommends the component library, design system, and interaction patterns before frontend coding starts.

## Backend Agent
- Owns Flask routes, business logic, script orchestration, user provisioning, licensing rules, ticket flows, and integrations.
- Builds reusable platform primitives so each new script plugs into the same system.
- Implements the code approved by PM, CTO, and frontend planning.

## QA Agent
- Owns verification at every meaningful milestone.
- Checks that new work does not break existing scripts or shared platform behavior.
- Verifies customer-facing text, output quality, and critical business flows.
- Blocks release if regression risk is not understood.

## Operations Agent
- Owns day-to-day operational readiness.
- Tracks admin alerts, support workflow, internal handoffs, and which tasks require user action in the terminal.
- Makes sure tickets, customer requests, and production follow-ups are routed and visible.
- Defines what should trigger an email, an admin task, or a manual review.

## Security Agent
- Owns authentication, password handling, permissions, session safety, and secrets management.
- Reviews trial access, licensing transitions, admin access, uploads, and customer data exposure risks.
- Prevents insecure shortcuts in billing and onboarding flows.

## DevOps Agent
- Owns deployment reliability.
- Manages Render services, environment variables, branch-to-environment mapping, backups, logs, and monitoring.
- Keeps `zman-emet-platform` as the testing environment and `scriptly-platform` as the production environment target.

## Data Product Agent
- Owns the usefulness of script outputs.
- Translates customer operational pain into outputs customers will pay for.
- Ensures each script produces a clear deliverable, not just raw processed data.

## Memory Updater Agent
- Owns durable project memory.
- Updates approved vision, tasks, specs, decisions, risks, and release notes.
- Records what was approved, what was rejected, and what still needs validation so the user does not need to repeat context.

## Agent Rules
- No coding starts without a mission brief from Project Manager.
- Every new script must have:
  - customer problem definition
  - reuse hypothesis for future customers
  - sample file analysis
  - acceptance criteria
  - test and rollback path
- QA must review every meaningful feature before release.
- Memory must be updated after each meaningful change.
- If a fact is uncertain, ask instead of guessing.

## Latest Delivery Memory
- Completed deliverables in the current delivery cycle:
  - Matan manual corrections report
  - Rimon home office summary report
  - Organizational hierarchy report
- Production status:
  - These features were tested successfully on production.
  - Production dashboard URL: `https://scriptly-platform.onrender.com/dashboard`
- Organizational hierarchy report V1 status:
  - Excel output exists.
  - PowerPoint output exists.
  - ZIP output exists.
  - Output selector exists with:
    - Excel only
    - PowerPoint only
    - both together
  - PowerPoint currently focuses on managers and departments.
  - Employees are not shown inside the org diagram itself.
  - Employee names appear only in manager/team summary slides.
  - Names in the PowerPoint are displayed as first-name then last-name.
  - Current state is good enough for V1.
- Preserve for future work:
  - Smart field mapping with user confirmation
  - Admin Hebrew localization / language switch
- Process lesson:
  - In this environment, long or broad terminal commands are unreliable.
  - Future agent work should be split into very small, fast steps with visible progress.
