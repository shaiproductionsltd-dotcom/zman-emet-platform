#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Assistant smoke / regression suite.

Runnable by hand or in CI:

    python assistant_smoke.py

Verifies:
  1. AI_TOOL_METADATA parity with SCRIPT_REGISTRY (every tool has metadata).
  2. Hebrew keyword shortlister matches realistic user phrasing.
  3. Tools knowledge block carries the access-status invariant per tool.
  4. Recommendation marker parser accepts/rejects correctly.
  5. Retention helpers + DB tables exist.

Exits non-zero on any failure so CI can gate on it.
"""

from __future__ import annotations

import os
import sys

# Prevent accidental Anthropic calls — none of these tests need a key
os.environ.setdefault("ANTHROPIC_API_KEY", "")

# Make the repo root importable regardless of CWD
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import app as A  # noqa: E402


# Each case: (user query, expected tool_id OR None for "no recommendation")
HEBREW_RECOMMENDATION_CASES = [
    # nikuy
    ("הדוח שלי מלא בכוכביות וסימני שאלה, איך אני מנקה אותו?", "nikuy"),
    ("צריך לנקות את הדוח לפני שעושים חישוב שכר", "nikuy"),
    ("יש לי דוח גולמי מאוד מלוכלך", "nikuy"),
    # flamingo_payroll
    ("איך מחשבים שכר לעובדים לפי תעריף שעתי?", "flamingo_payroll"),
    ("צריך לחשב משכורות חודשיות", "flamingo_payroll"),
    ("רוצה סיכום שכר לפי שעות לתשלום", "flamingo_payroll"),
    ("שכר שעתי לעובדים על בסיס תעריף קבוע", "flamingo_payroll"),
    # dept_payroll
    ("יש לי חברת כוח אדם, צריך לחייב לקוחות וגם לשלם לעובדים", "dept_payroll"),
    ("דוח מחלקתי עם ניכוי דירה ומס 3%", "dept_payroll"),
    ("חשבוניות ללקוחות לפי תעריף גביה", "dept_payroll"),
    ("צריך סיכום חברה עם חיוב ללקוח ותשלום לעובד", "dept_payroll"),
    # matan_missing
    ("אני רוצה לראות מי עבד פחות מהתקן החודש", "matan_missing"),
    ("צריך לסנן עובדים לפי שעות חוסר", "matan_missing"),
    ("הפרש מול תקן השעות החודשי", "matan_missing"),
    ("מי שעבד מתחת לתקן", "matan_missing"),
    # inactive_workers
    ("מי הפסיק לעבוד אצלנו בחודשיים האחרונים?", "inactive_workers"),
    ("רשימת עובדים שלא הופיעו לאחרונה", "inactive_workers"),
    ("יש לנו עובדים רדומים שצריך לנקות", "inactive_workers"),
    ("איך אני מוצא עובדים שעזבו בלי לעדכן?", "inactive_workers"),
    ("עובדים ללא דיווח נוכחות", "inactive_workers"),
    # matan_manual_corrections
    ("בקרה על תיקונים ידניים בשעון נוכחות", "matan_manual_corrections"),
    ("יש לי חשד לשינויים ידניים בשעון", "matan_manual_corrections"),
    ("רוצה לראות את כל תיקוני הכניסה והיציאה הידניים", "matan_manual_corrections"),
    # rimon_home_office_summary
    ("צריך סיכום של ימי בית מול ימי משרד", "rimon_home_office_summary"),
    ("מעקב על עבודה היברידית", "rimon_home_office_summary"),
    ("home office report לחודש שעבר", "rimon_home_office_summary"),
    ("עבודה מרחוק לעובדים שלי", "rimon_home_office_summary"),
    # attendance_alerts
    ("צריך לבדוק חריגות של שעות נוספות ומנוחה שבועית", "attendance_alerts"),
    ("מי עבד שעות לילה או בשבת?", "attendance_alerts"),
    ("בעיות נוכחות לבני נוער שעובדים אצלי", "attendance_alerts"),
    ("התראות על יציאות מאוחרות וימים ארוכים", "attendance_alerts"),
    # office_occupancy_heatmap
    ("רוצה מפת חום של עומס במשרד לפי ימים ושעות", "office_occupancy_heatmap"),
    ("תכנון מקומות ישיבה ו-hot desking", "office_occupancy_heatmap"),
    ("עומס במשרד לפי יום בשבוע", "office_occupancy_heatmap"),
    # org_hierarchy_report
    ("צריך תרשים מבנה ארגוני להנהלה", "org_hierarchy_report"),
    ("היררכיה של מנהלים ומחלקות מקובץ CSV", "org_hierarchy_report"),
    ("אורגנוגרם של החברה", "org_hierarchy_report"),
    # vague but should still resolve correctly
    ("אני צריך משהו שיעזור לי לראות מי לא עובד יותר", "inactive_workers"),
    ("יש לי קובץ נוכחות ואני רוצה לחשב כמה לשלם לכל עובד", "flamingo_payroll"),
    # negative cases — should NOT recommend any tool
    ("מה החוק לגבי דמי הבראה?", None),
    ("שלום, מה שלומך היום?", None),
    ("איך אני מחשב נקודות זיכוי?", None),
]


def _passed(label: str):
    print(f"  PASS  {label}")


def _failed(label: str, detail: str = ""):
    msg = f"  FAIL  {label}"
    if detail:
        msg += f"  --  {detail}"
    print(msg)


def check_metadata_parity():
    print("[1] AI_TOOL_METADATA parity with SCRIPT_REGISTRY")
    reg = set(A.SCRIPT_REGISTRY.keys())
    ai = set(A.AI_TOOL_METADATA.keys())
    failures = 0
    extra_in_reg = sorted(reg - ai)
    extra_in_ai = sorted(ai - reg)
    if extra_in_reg:
        _failed("registry tools missing from AI_TOOL_METADATA", str(extra_in_reg))
        failures += 1
    if extra_in_ai:
        _failed("AI_TOOL_METADATA tools missing from registry", str(extra_in_ai))
        failures += 1
    for tid, meta in A.AI_TOOL_METADATA.items():
        for f in ("summary", "inputs", "outputs", "when_to_use", "keywords"):
            if not meta.get(f):
                _failed(f"{tid} missing field {f}")
                failures += 1
    if not failures:
        _passed(f"{len(reg)} tools, all paired with full metadata")
    return failures


def check_hebrew_recommendations():
    """The contract is: the right tool must be in the shortlist (top 3) so the
    full-detail block reaches the model. Strict #1 ranking is not required —
    Claude picks within the shortlist. Negative cases must produce zero hits."""
    print(f"[2] Hebrew shortlister regression ({len(HEBREW_RECOMMENDATION_CASES)} cases)")
    failures = 0
    for query, expected in HEBREW_RECOMMENDATION_CASES:
        top = A.shortlist_tools_for_message(query, limit=3)
        if expected is None:
            ok = len(top) == 0
        else:
            ok = expected in top
        if ok:
            _passed(f"{query[:45]:<45}  -> {top!r}")
        else:
            _failed(f"{query[:45]:<45}  expected={expected!r} got={top!r}")
            failures += 1
    return failures


def check_access_invariant():
    print("[3] Access-status invariant per tool")
    failures = 0
    block = A.render_tools_knowledge_block(
        accessible_tool_ids=["flamingo_payroll", "dept_payroll"],
        last_user_message="חיוב לקוחות של חברת כוח אדם",
    )
    # Per-tool inspection: the section the tool falls into should match its access tag
    for tid, expected_label in (
        ("flamingo_payroll", "זמין לשימוש"),
        ("dept_payroll", "זמין לשימוש"),
        ("nikuy", "זמין"),  # one-line list, [זמין] would be wrong here — it's not accessible
    ):
        idx = block.find(f"`{tid}`")
        if idx < 0:
            _failed(f"{tid} missing from knowledge block")
            failures += 1
            continue
        tail = block[idx:idx + 600]
        if tid in ("flamingo_payroll", "dept_payroll"):
            if "זמין לשימוש" not in tail:
                _failed(f"{tid} should be marked available", tail[:200])
                failures += 1
            elif "אין הרשאה (אם הכלי מתאים" in tail:
                _failed(f"{tid} wrongly marked no-permission", tail[:200])
                failures += 1
            else:
                _passed(f"{tid} marked available correctly")
        elif tid == "nikuy":
            if "[אין הרשאה]" not in tail:
                _failed(f"nikuy should be tagged [אין הרשאה] in others list", tail[:200])
                failures += 1
            else:
                _passed("nikuy correctly tagged in others list")
    if "מקור האמת היחיד לגישת המשתמש" not in block:
        _failed("absolute access rule missing from prompt")
        failures += 1
    else:
        _passed("absolute access rule present at top of section")
    return failures


def check_marker_parser():
    print("[4] Recommendation marker parser")
    failures = 0
    cases = [
        (
            "הנה.\n---RECOMMEND---\ntool_id: flamingo_payroll\nreason: שעות לתשלום × תעריף\n---END---",
            ("recommend", "flamingo_payroll"),
        ),
        (
            "אין כלי קיים מתאים.\n---SUGGEST_BUILD---\nbrief: כלי לחיזוי עזיבת עובדים\n---END---",
            ("build", None),
        ),
        ("סתם תשובה רגילה.", (None, None)),
        (
            "פייק:\n---RECOMMEND---\ntool_id: bogus_tool\nreason: x\n---END---",
            ("none", None),  # invalid tool_id must drop
        ),
    ]
    for text, expected in cases:
        clean, rec, sb = A.parse_assistant_output(text)
        kind, tool_id = expected
        if kind == "recommend":
            if rec and rec.get("tool_id") == tool_id:
                _passed(f"recommend -> {tool_id}")
            else:
                _failed("recommend parsing", f"{rec!r}")
                failures += 1
        elif kind == "build":
            if sb and sb.get("brief"):
                _passed("suggest_build with brief")
            else:
                _failed("build parsing", f"{sb!r}")
                failures += 1
        elif kind == "none":
            if rec is None:
                _passed("invalid tool_id correctly dropped")
            else:
                _failed("invalid tool_id leaked", f"{rec!r}")
                failures += 1
        else:
            if rec is None and sb is None:
                _passed("plain text -> no markers")
            else:
                _failed("plain text wrongly parsed", f"rec={rec} sb={sb}")
                failures += 1
        # Ensure markers stripped
        if "---RECOMMEND---" in clean or "---SUGGEST_BUILD---" in clean or "---END---" in clean:
            _failed("markers leaked into displayed text", clean)
            failures += 1
    return failures


def check_retention_plumbing():
    print("[5] Retention plumbing")
    failures = 0
    try:
        with A.get_db() as db:
            for table in ("session_artifacts", "assistant_recommendations"):
                row = db.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                    (table,),
                ).fetchone()
                # On Postgres sqlite_master doesn't exist; skip test
                if row is not None:
                    _passed(f"{table} present")
                else:
                    # Try a select to confirm table exists on Postgres
                    try:
                        db.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone()
                        _passed(f"{table} present (postgres)")
                    except Exception as e:
                        _failed(f"{table} missing", str(e))
                        failures += 1
    except Exception as e:
        _failed("DB access during retention check", str(e))
        failures += 1
    try:
        stats = A.run_scheduled_cleanup()
        if isinstance(stats, dict) and "artifacts_expired" in stats:
            _passed(f"run_scheduled_cleanup() returned stats {sorted(stats.keys())}")
        else:
            _failed("run_scheduled_cleanup() returned unexpected shape", str(stats))
            failures += 1
    except Exception as e:
        _failed("run_scheduled_cleanup raised", str(e))
        failures += 1
    return failures


def main():
    total = 0
    total += check_metadata_parity()
    print()
    total += check_hebrew_recommendations()
    print()
    total += check_access_invariant()
    print()
    total += check_marker_parser()
    print()
    total += check_retention_plumbing()
    print()
    if total == 0:
        print("==== ALL ASSISTANT SMOKE CHECKS PASSED ====")
        return 0
    print(f"==== {total} FAILURE(S) ====")
    return 1


if __name__ == "__main__":
    sys.exit(main())
