# CEO Reporting System

This file defines what the CEO agent should receive regularly from all company agents.

## Purpose

The CEO should not need to read raw technical conversations.

The CEO should receive short management reports that answer:
- What was built
- What is blocked
- What is risky
- What is next
- Whether the company is becoming stronger or weaker

## Report Types

## 1. Daily Executive Report

Audience:
- CEO

Prepared by:
- Project Manager Agent

Inputs:
- CTO
- Backend
- QA
- DevOps
- Memory Updater

Format:
- 1 short page

Contents:
- completed today
- blocked today
- production status
- customer-impacting issues
- next priority tomorrow

## 2. Weekly Company Report

Audience:
- CEO
- CTO

Prepared by:
- Project Manager Agent

Contents:
- features completed this week
- bugs fixed this week
- regressions introduced
- staging status
- production incidents
- current top risks
- top opportunities
- workload by team agent

## 3. Product Health Report

Audience:
- CEO
- CTO
- Data Product Agent

Contents:
- active scripts
- scripts in development
- scripts with quality issues
- average processing time
- customer complaints by script
- scripts with highest revenue potential

## 4. Security Report

Audience:
- CEO
- CTO
- Security Agent

Contents:
- password/security weaknesses
- failed login risks
- secrets handling issues
- upload/file risks
- permission-model risks
- required actions

## 5. Delivery Reliability Report

Audience:
- CEO
- Project Manager
- DevOps

Contents:
- how many tasks reached staging
- how many tasks reached production
- how many rollbacks happened
- how many releases passed without incident
- what caused delays

## 6. Scale Readiness Report

Audience:
- CEO
- CTO

Contents:
- current active user capacity
- current production risks
- missing infrastructure for 1000 users
- progress against scale plan

## CEO Dashboard Metrics

The CEO agent should always have these numbers available:
- active customers
- active users
- monthly recurring revenue
- scripts available for sale
- premium scripts available
- script usage per customer
- failed runs per week
- average processing time
- open critical bugs
- open security issues
- open staging-only issues

## Mission Status Model

Every mission should be reported in one of these states:
- Discovery
- Defined
- Building
- QA
- Staging
- Ready for Production
- Live
- Blocked

## CEO Escalation Triggers

The CEO must be alerted immediately if:
- production is broken
- customer data may be at risk
- a release regresses an existing paid script
- security risk is discovered
- a feature misses launch timing due to avoidable internal disorder

## Required Report Cadence

- Daily executive report: every workday
- Weekly company report: once a week
- Security report: weekly or on incident
- Scale readiness report: every two weeks

## Current Immediate CEO Priorities

1. Stabilise the release workflow.
2. Create staging before further serious expansion.
3. Clean the UI text layer.
4. Fix the Matan production regression.
5. Resume product expansion only after the foundation is stable.
