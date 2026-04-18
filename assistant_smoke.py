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
    print("[4] Recommendation marker parser (3-way classification)")
    failures = 0
    # Each case: (label, text, predicate_on_(rec, sb))
    def _expect_recommend(tool_id):
        return lambda rec, sb: bool(rec) and rec.get("tool_id") == tool_id and sb is None

    def _expect_self_serve():
        return lambda rec, sb: (
            rec is None and bool(sb)
            and sb.get("kind") == "self_serve"
            and "/tools/create" in (sb.get("url") or "")
            and "mode=self_serve" in (sb.get("url") or "")
        )

    def _expect_escalate():
        return lambda rec, sb: (
            rec is None and bool(sb)
            and sb.get("kind") == "escalate"
            and "/tools/create" in (sb.get("url") or "")
            and "mode=escalate" in (sb.get("url") or "")
        )

    def _expect_legacy_self_serve():
        # Legacy SUGGEST_BUILD must route to self_serve for back-compat
        return lambda rec, sb: rec is None and bool(sb) and sb.get("kind") == "self_serve"

    def _expect_nothing():
        return lambda rec, sb: rec is None and sb is None

    def _expect_self_serve_wins_over_escalate():
        # If both markers leak (model error), self-serve must win
        return lambda rec, sb: rec is None and bool(sb) and sb.get("kind") == "self_serve"

    cases = [
        ("recommend valid",
         "הנה.\n---RECOMMEND---\ntool_id: flamingo_payroll\nreason: שעות לתשלום × תעריף\n---END---",
         _expect_recommend("flamingo_payroll")),
        ("self-serve build",
         "זה מסוג הכלים שאתה יכול לבנות בעצמך.\n---SELF_SERVE_BUILD---\nbrief: כלי שמסנן עובדים מתחת ל-100 שעות חודשיות\n---END---",
         _expect_self_serve()),
        ("escalate to platform",
         "זו בקשה שדורשת אינטגרציה עם מערכת חיצונית.\n---ESCALATE_TO_PLATFORM_TEAM---\nbrief: סנכרון אוטומטי עם מערכת ה-HR החיצונית\n---END---",
         _expect_escalate()),
        ("legacy SUGGEST_BUILD routes to self_serve",
         "אין כלי קיים מתאים.\n---SUGGEST_BUILD---\nbrief: כלי לחיזוי עזיבת עובדים\n---END---",
         _expect_legacy_self_serve()),
        ("plain text -> no markers", "סתם תשובה רגילה.", _expect_nothing()),
        ("invalid tool_id dropped",
         "פייק:\n---RECOMMEND---\ntool_id: bogus_tool\nreason: x\n---END---",
         _expect_nothing()),
        ("self-serve wins when both markers present",
         "כפול:\n---SELF_SERVE_BUILD---\nbrief: סינון עובדים\n---END---\n\n---ESCALATE_TO_PLATFORM_TEAM---\nbrief: צריך אינטגרציה\n---END---",
         _expect_self_serve_wins_over_escalate()),
    ]
    for label, text, predicate in cases:
        clean, rec, sb = A.parse_assistant_output(text)
        if predicate(rec, sb):
            _passed(label)
        else:
            _failed(label, f"rec={rec!r} sb={sb!r}")
            failures += 1
        # Ensure NO marker token leaks into display text in any case
        for token in ("---RECOMMEND---", "---SELF_SERVE_BUILD---",
                      "---ESCALATE_TO_PLATFORM_TEAM---", "---SUGGEST_BUILD---",
                      "---END---"):
            if token in clean:
                _failed(f"{label}: marker token leaked into display", token)
                failures += 1
    return failures


def check_widget_blank_state():
    """The dashboard assistant widget must render a multi-role blank-state
    with starter chips covering all six supported intents, and reset/refresh
    must restore the same starter experience (single source of truth)."""
    print("[10] Dashboard widget blank-state covers all 6 intents")
    failures = 0
    html = A.render_assistant_widget()
    must_contain = [
        # Multi-role intro copy
        "יועץ Scriptly",
        "למצוא לך כלי שמתאים",
        "self-serve",
        "דורשת פיתוח",
        "לנתח קובץ",
        "HR / שכר / דוחות",
        # Six chip texts (one per intent the user listed)
        "איזה כלי מתאים לי?",                  # find existing
        "אני רוצה לבנות כלי חדש",              # build new
        "לבנות לבד או צריך פיתוח?",            # boundary
        "תעזור לי לנתח קובץ שאעלה",            # upload + analyze
        "שאלה מהירה על שעות / שכר / חוקי עבודה",  # quick HR Q
        "מה אתה יודע לעשות?",                  # what-can-you-do
        # Free-typing reassurance
        "אתה תמיד חופשי לכתוב כל דבר",
        # JS plumbing — single source of truth + restore behavior
        "var astBlankHtml=",
        "function astRenderBlankState(",
        "function astTriggerUpload(",
        # Reset path uses the renderer (no more one-line welcome)
        "astRenderBlankState();",
        # Empty-init path uses the renderer too (so refresh after server
        # cleanup still shows the chips)
        # The renderer is referenced inside astInitSession's else branch
    ]
    for needle in must_contain:
        if needle in html:
            _passed(f"widget contains: {needle[:55]}")
        else:
            _failed(f"widget missing required anchor: {needle[:55]}")
            failures += 1
    # Reset must NOT use the old one-line welcome string
    if "שיחה חדשה נפתחה. אפשר לשאול." in html:
        _failed("old one-line reset welcome still present — should be replaced by astRenderBlankState()")
        failures += 1
    else:
        _passed("old one-line reset welcome removed")
    # Upload chip must trigger the file picker, not send text
    if 'onclick="astTriggerUpload()"' in html:
        _passed("upload chip wired to file picker (not text send)")
    else:
        _failed("upload chip not wired to astTriggerUpload")
        failures += 1
    # The wipe selectors that hide the blank-state when chat starts must
    # also include the new hint line so it doesn't linger after first message
    wipe_count = html.count('".ast-welcome,.ast-chips,.ast-bs-hint"')
    if wipe_count >= 2:
        _passed(f"hint line wipe wired in {wipe_count} sites (addMessage + renderHistory)")
    else:
        _failed(f"hint wipe wiring incomplete (count={wipe_count}, expected >=2)")
        failures += 1
    return failures


def check_tool_creation_early_classification():
    """The /tools/create assistant must commit to a classification in its
    first response and explain any mid-conversation path change. This guards
    against late-surprise escalation — the trust failure we just fixed."""
    print("[9] Tool-creation prompt enforces early classification")
    failures = 0
    p = A.TOOL_CREATION_SYSTEM_PROMPT_BASE
    must_contain = [
        # Section anchor
        "סיווג מוקדם וברור",
        "חוק ברזל",
        # First-response commitment language
        "בתשובה הראשונה שלך",
        "אסור לרוץ לאיסוף פרטים מפורט לפני שהסיווג הוצהר",
        # Four canonical Hebrew framings (as actual user-facing strings)
        "כבר קיים כלי שעושה את זה",
        "זה נראה כמו כלי שתוכל לבנות כאן בעצמך",
        "זו נראית בקשה שכנראה דורשת פיתוח של צוות הפלטפורמה",
        "אני צריך הבהרה אחת",
        # Path-change rule
        "שינוי סיווג באמצע השיחה",
        "בהתחלה זה נראה כמו כלי שאפשר לבנות כאן",
        "בהתחלה חשבתי שזה דורש פיתוח",
        # Hard prohibitions
        "אסור להציג developer_brief בסוף שיחה שניהלת כ-self-serve",
        "אסור להציג JSON של כלי בסוף שיחה שניהלת כצריכת פיתוח",
        "אסור להמתין לסוף השיחה כדי \"להפתיע\" את המשתמש",
        # Output-section guards reinforce the rule at the emission point
        "אסור להציג את הבריף הזה אלא אם המשתמש כבר יודע מתחילת השיחה",
        "אסור להציג JSON של כלי אלא אם המשתמש כבר יודע מתחילת השיחה",
        # Mode B opener now requires classification first
        "שלב 0 (חובה לפני כל שלב אחר)",
        # Honor seeded mode-aware arrivals from the dashboard assistant
        "אם המשתמש הגיע משיחה עם היועץ עם סיווג מוצהר",
    ]
    for needle in must_contain:
        if needle in p:
            _passed(f"prompt contains: {needle[:55]}")
        else:
            _failed(f"prompt missing required anchor: {needle[:55]}")
            failures += 1
    return failures


def check_create_mode_framing():
    """The /tools/create page must look meaningfully different depending on
    how the user arrived (self-serve vs escalate vs no-mode)."""
    print("[8] /tools/create mode-aware framing helper")
    failures = 0

    # No mode -> default behavior preserved
    banner, welcome, hide_sugg, seed = A._build_create_mode_framing("", "")
    if banner == "" and welcome is None and hide_sugg is False and seed is None:
        _passed("no mode -> page renders with original defaults")
    else:
        _failed("no-mode arrival altered the default page",
                f"banner={bool(banner)} welcome={welcome is not None} hide={hide_sugg} seed={seed is not None}")
        failures += 1

    # Self-serve mode with brief
    banner, welcome, hide_sugg, seed = A._build_create_mode_framing(
        "self_serve",
        "כלי שמסנן עובדים מתחת ל-100 שעות חודשיות\nקלט: דוח נוכחות חודשי\nפלט: רשימת עובדים עם שעות בפועל"
    )
    if banner and "לבנות בעצמך" in banner and "כלי שמסנן עובדים" in banner:
        _passed("self_serve banner shows empowering framing + brief")
    else:
        _failed("self_serve banner missing or wrong", banner[:200])
        failures += 1
    if welcome and "בנייה" in welcome or (welcome and "בעצמך" in welcome):
        _passed("self_serve welcome bubble overrides default")
    else:
        _failed("self_serve welcome override missing")
        failures += 1
    if hide_sugg is False:
        _passed("self_serve keeps generic suggestion chips")
    else:
        _failed("self_serve should keep suggestion chips")
        failures += 1
    if seed and "בעצמך" in seed and "כלי שמסנן עובדים" in seed:
        _passed("self_serve seed message includes brief verbatim")
    else:
        _failed("self_serve seed text wrong", (seed or "")[:200])
        failures += 1

    # Escalate mode with brief
    banner, welcome, hide_sugg, seed = A._build_create_mode_framing(
        "escalate",
        "אינטגרציה עם מערכת HR חיצונית — סנכרון אוטומטי של עובדים"
    )
    if banner and "פיתוח" in banner and "אינטגרציה" in banner:
        _passed("escalate banner shows platform-team framing + brief")
    else:
        _failed("escalate banner missing or wrong", banner[:200])
        failures += 1
    if welcome and ("צוות" in welcome or "פיתוח" in welcome):
        _passed("escalate welcome bubble overrides default")
    else:
        _failed("escalate welcome override missing")
        failures += 1
    if hide_sugg is True:
        _passed("escalate hides generic self-serve suggestion chips")
    else:
        _failed("escalate should hide suggestion chips")
        failures += 1
    if seed and "developer brief" in seed and "אינטגרציה" in seed:
        _passed("escalate seed message frames spec handoff + carries brief")
    else:
        _failed("escalate seed text wrong", (seed or "")[:200])
        failures += 1

    # Brief truncation safety
    huge = "x" * 5000
    _, _, _, seed = A._build_create_mode_framing("self_serve", huge)
    if seed and len([c for c in seed if c == "x"]) <= 1500:
        _passed("brief truncated to <=1500 chars before seeding")
    else:
        _failed("brief truncation failed", str(len(seed or "")))
        failures += 1

    # Unknown mode is treated as no-mode
    banner, welcome, hide_sugg, seed = A._build_create_mode_framing("garbage", "x")
    if banner == "" and welcome is None and seed is None:
        _passed("unknown mode falls back to no-op (defensive)")
    else:
        _failed("unknown mode leaked framing", str(banner)[:120])
        failures += 1

    # Parser URLs must include mode= so plain anchor clicks (no JS) still
    # land in the right framing
    _, _, sb1 = A.parse_assistant_output(
        "ok\n---SELF_SERVE_BUILD---\nbrief: סינון עובדים\n---END---"
    )
    if sb1 and "mode=self_serve" in (sb1.get("url") or ""):
        _passed("parser sets mode=self_serve in url")
    else:
        _failed("parser URL missing mode=self_serve", str(sb1))
        failures += 1
    _, _, sb2 = A.parse_assistant_output(
        "ok\n---ESCALATE_TO_PLATFORM_TEAM---\nbrief: אינטגרציה\n---END---"
    )
    if sb2 and "mode=escalate" in (sb2.get("url") or ""):
        _passed("parser sets mode=escalate in url")
    else:
        _failed("parser URL missing mode=escalate", str(sb2))
        failures += 1
    return failures


def check_three_way_prompt_doc():
    """Make sure the prompt actually documents the 3-way classification with both markers,
    so Claude has the contract to follow. This catches accidental drift."""
    print("[7] Assistant prompt documents 3-way classification")
    failures = 0
    p = A.ASSISTANT_CHAT_SYSTEM_PROMPT
    must_contain = [
        "RECOMMEND_EXISTING_TOOL",
        "SELF_SERVE_BUILD",
        "ESCALATE_TO_PLATFORM_TEAM",
        "---SELF_SERVE_BUILD---",
        "---ESCALATE_TO_PLATFORM_TEAM---",
        # Hebrew anchors so we know the explanatory section is present
        "ברירת מחדל = SELF_SERVE_BUILD",
        "אינטגרצי",  # match אינטגרציה / אינטגרציות
    ]
    for needle in must_contain:
        if needle in p:
            _passed(f"prompt contains: {needle}")
        else:
            _failed(f"prompt missing required anchor: {needle}")
            failures += 1
    return failures


def check_marketplace_handling():
    """Verify marketplace surfacing + name resolution."""
    print("[5] Marketplace tool surfacing + name resolution")
    failures = 0
    fake_installed = [
        {"id": 101, "name": "מחשבון פיצויים", "description": "חישוב פיצויי פיטורים לעובדים מסיימים"},
        {"id": 102, "name": "סיכום נסיעות חודשי", "description": "סיכום הוצאות נסיעה לעובדים"},
    ]
    block = A.render_tools_knowledge_block(
        accessible_tool_ids=[],
        last_user_message="צריך לחשב פיצויי פיטורים לעובד שמסיים השבוע",
        installed_marketplace=fake_installed,
    )
    # The relevant marketplace tool should be highlighted
    if "כלי שוק מותקנים שנראים רלוונטיים" in block and "מחשבון פיצויים" in block:
        _passed("relevant marketplace tool surfaced in priority section")
    else:
        _failed("marketplace surfacing missed", block[:300])
        failures += 1
    # Both tools should appear with [זמין] tag
    for nm in ("מחשבון פיצויים", "סיכום נסיעות"):
        if f"{nm}" in block and "[זמין]" in block:
            _passed(f"marketplace tool '{nm}' tagged as available")
        else:
            _failed(f"marketplace tool '{nm}' missing or untagged")
            failures += 1
    # Name resolution: invalid id must drop, unknown numeric id must drop, format error must drop
    for raw in ("marketplace:99999999", "marketplace:abc", "marketplace:"):
        _, name, url = A._resolve_marketplace_tool(raw)
        if name is None:
            _passed(f"invalid marketplace tool_id '{raw}' correctly dropped")
        else:
            _failed(f"invalid marketplace tool_id '{raw}' resolved to {name}")
            failures += 1
    return failures


def check_retention_plumbing():
    print("[6] Retention plumbing")
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
        expected_keys = {
            "sessions_wiped_inactivity", "sessions_wiped_hard_cap",
            "sessions_deleted", "artifacts_expired", "token_rows_purged",
            "report_jobs_expired", "orphan_files_removed", "report_pii_scrubbed",
        }
        missing = expected_keys - set(stats.keys())
        if missing:
            _failed("run_scheduled_cleanup missing stat keys", str(missing))
            failures += 1
        else:
            _passed(f"run_scheduled_cleanup returns all {len(expected_keys)} stat keys")
        # Last-run tracking was captured
        if A._LAST_CLEANUP.get("ran_at") and A._LAST_CLEANUP.get("stats"):
            _passed(f"_LAST_CLEANUP populated (duration={A._LAST_CLEANUP['duration_ms']}ms)")
        else:
            _failed("_LAST_CLEANUP not captured")
            failures += 1
    except Exception as e:
        _failed("run_scheduled_cleanup raised", str(e))
        failures += 1

    # Per-step isolation: simulate a failed step shouldn't abort others.
    # We can't easily force a DB failure here without mocking; confirm the
    # function has the expected structural guard (stats dict populated).
    stats2 = A.run_scheduled_cleanup()
    if isinstance(stats2.get("orphan_files_removed"), int):
        _passed("orphan_files_removed is an int even on clean runs")
    else:
        _failed("orphan_files_removed not int", str(stats2))
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
    total += check_marketplace_handling()
    print()
    total += check_retention_plumbing()
    print()
    total += check_three_way_prompt_doc()
    print()
    total += check_create_mode_framing()
    print()
    total += check_tool_creation_early_classification()
    print()
    total += check_widget_blank_state()
    print()
    if total == 0:
        print("==== ALL ASSISTANT SMOKE CHECKS PASSED ====")
        return 0
    print(f"==== {total} FAILURE(S) ====")
    return 1


if __name__ == "__main__":
    sys.exit(main())
