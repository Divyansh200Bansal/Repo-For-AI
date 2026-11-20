#!/usr/bin/env python3
"""
varc.py - offline helper for the daily CAT VARC routine.

Use it with any free chat assistant (the free Claude plan works). The chat writes
the news, passage and questions; this script does what a free chat cannot do:
it picks today's question mix from your log, times you, and writes your results
into the Excel mistake log.

  python varc.py brief    pick today's question mix and Q7 type, write prompt_today.txt
  python varc.py timer    stopwatch for START, DONE and Q1-Q7; prints a TIMES line
  python varc.py log      paste the LOG BLOCK from the debrief; adds rows to the log
  python varc.py stats    accuracy by question type, mistake reasons, last 7 days

Options go before the command, e.g.  python varc.py --file other.xlsx stats
Needs Python 3.8+ and openpyxl:  pip install openpyxl
"""
import argparse
import datetime as dt
import json
import re
import shutil
import sys
import time
from copy import copy
from pathlib import Path

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is missing. Install it with:  pip install openpyxl")

HERE = Path(__file__).resolve().parent
LOG_FILE = HERE / "CAT_VARC_Mistake_Log.xlsx"
TEMPLATE = HERE / "templates" / "CAT_VARC_Mistake_Log_template.xlsx"
PROMPT_TEMPLATE = HERE / "prompts" / "free_prompt.md"
PROMPT_OUT = HERE / "prompt_today.txt"
RECORDS = HERE / "records"
PROMPT_MARKER = "<!-- PROMPT STARTS BELOW -->"
BLOCK_END = "=== LOG BLOCK END ==="

# Labels must match the Summary tab's formulas exactly.
Q_TYPES = ["Main idea", "Inference", "Detail", "Tone", "Vocab in context",
           "Strengthen/Weaken", "Para summary", "Para-jumble", "Odd sentence out"]
RC_TYPES = Q_TYPES[:6]
VA_ORDER = ["Para-jumble", "Para summary", "Odd sentence out"]  # Q7 tie-break order
BASE_MIX = ["Main idea", "Inference", "Inference", "Detail", "Tone", "Vocab in context"]
REASONS = ["Misread question", "Extreme option", "Out of scope", "Over-inference",
           "Partial answer", "Vocab gap", "Missed link (PJ)", "Time pressure", "Careless"]
RESULTS = ["Correct", "Wrong"]
LOG_COLS = ["Date", "Q type", "Result", "Time (sec)", "Mistake reason", "My answer",
            "Correct answer", "Topic", "Q no.", "Notes"]
DAILY_COLS = ["Date", "Score", "Out of", "Reading (min)", "Speed (wpm)",
              "Total time (min)", "Topic", "Words"]
VOCAB_COLS = ["Word", "Meaning", "Status", "Example", "Added", "Source"]
WEAK_MIN_ATTEMPTS, WEAK_ACCURACY = 3, 0.60  # swap rule: 3+ attempts and below 60%
SLOW_SEC = 120                              # flag questions over 2 minutes

ALIASES = {
    "pj": "Para-jumble", "parajumbles": "Para-jumble", "jumble": "Para-jumble",
    "summary": "Para summary", "paragraphsummary": "Para summary",
    "oddoneout": "Odd sentence out", "oddsentence": "Odd sentence out", "oso": "Odd sentence out",
    "vocab": "Vocab in context", "vocabulary": "Vocab in context", "vic": "Vocab in context",
    "strengthen": "Strengthen/Weaken", "weaken": "Strengthen/Weaken",
    "main": "Main idea", "centralidea": "Main idea", "mainpoint": "Main idea",
    "right": "Correct", "incorrect": "Wrong",
    "extreme": "Extreme option", "partial": "Partial answer", "misread": "Misread question",
    "missedlink": "Missed link (PJ)",
}


# ---------- small helpers ----------

def _key(text):
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def canon(value, allowed, field):
    """Return the exact label for a loosely typed one ('para jumble' -> 'Para-jumble')."""
    k = _key(value)
    for label in allowed:
        if _key(label) == k:
            return label
    alias = ALIASES.get(k)
    if alias in allowed:
        return alias
    raise ValueError(f"{field} '{value}' is not an allowed label. Use one of: {', '.join(allowed)}")


def to_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d %b %Y", "%d %B %Y"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def num(value, field):
    """Parse a number from the LOG BLOCK; '-', '?' or blank mean 'not given'."""
    text = str(value).strip().lower().replace(",", "")
    text = re.sub(r"\s*(sec|secs|s|min|mins|wpm)$", "", text)
    if text in ("", "-", "?", "na", "n/a", "none"):
        return None
    try:
        f = float(text)
    except ValueError:
        raise ValueError(f"{field} '{value}' is not a number")
    return int(f) if f.is_integer() else f


def mmss(sec):
    m, s = divmod(int(round(sec)), 60)
    return f"{m}m {s:02d}s"


def ask_yes(question, assume_yes=False):
    if assume_yes:
        return True
    try:
        return input(f"{question} [y/N] ").strip().lower() in ("y", "yes")
    except EOFError:
        return False


# ---------- workbook access ----------

def open_book(path):
    if not path.exists():
        if path == LOG_FILE and TEMPLATE.exists():
            shutil.copy(TEMPLATE, path)
            print(f"Created {path.name} from the empty template.")
        else:
            sys.exit(f"Cannot find {path}. Put your mistake log there or pass --file.")
    wb = openpyxl.load_workbook(path)
    for name in ("Log", "Daily", "Vocab"):
        if name not in wb.sheetnames:
            sys.exit(f"{path.name} has no '{name}' tab. Is it the CAT VARC mistake log?")
    return wb


def header_map(ws, wanted):
    heads = {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}
    missing = [w for w in wanted if w not in heads]
    if missing:
        sys.exit(f"Tab '{ws.title}' is missing column(s): {', '.join(missing)}")
    return {w: heads[w] for w in wanted}


def last_filled_row(ws, cols):
    """Last row (1 = header) with a value in any of the given column numbers."""
    for r in range(ws.max_row, 1, -1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in cols):
            return r
    return 1


def read_rows(ws, wanted):
    hm = header_map(ws, wanted)
    rows = []
    for r in range(2, last_filled_row(ws, hm.values()) + 1):
        rec = {name: ws.cell(row=r, column=col).value for name, col in hm.items()}
        if any(v not in (None, "") for v in rec.values()):
            rows.append(rec)
    return rows


def extend_validations(ws, row, cols):
    """Make the tab's dropdowns (Q type, Result, ...) cover a newly added row."""
    for dv in ws.data_validations.dataValidation:
        col_ranges = [g for g in dv.sqref.ranges if g.min_col == g.max_col and g.min_col in cols]
        if col_ranges and not any(g.min_row <= row <= g.max_row for g in col_ranges):
            dv.add(f"{get_column_letter(col_ranges[0].min_col)}{row}")


def append_rows(ws, wanted, records, date_cols=()):
    """Write records below the last filled row. Existing rows are never touched."""
    hm = header_map(ws, wanted)
    last = last_filled_row(ws, hm.values())
    style_row = last if last >= 2 else None  # copy the look of the last data row
    first = last + 1
    for i, rec in enumerate(records):
        r = first + i
        for name, col in hm.items():
            cell = ws.cell(row=r, column=col)
            if style_row:
                src = ws.cell(row=style_row, column=col)
                if src.has_style:
                    cell._style = copy(src._style)
            cell.value = rec.get(name)
            # Keep the column's own date format; use ISO dates only where there is none.
            if name in date_cols and cell.value is not None and cell.number_format == "General":
                cell.number_format = "yyyy-mm-dd"
        extend_validations(ws, r, set(hm.values()))
    return first, first + len(records) - 1


# ---------- stats and today's settings ----------

def summarize(log_rows):
    by_type = {t: {"n": 0, "ok": 0, "secs": []} for t in Q_TYPES}
    reasons = {r: 0 for r in REASONS}
    for rec in log_rows:
        t = rec.get("Q type")
        if t in by_type:
            by_type[t]["n"] += 1
            by_type[t]["ok"] += rec.get("Result") == "Correct"
            if isinstance(rec.get("Time (sec)"), (int, float)):
                by_type[t]["secs"].append(rec["Time (sec)"])
        if rec.get("Mistake reason") in reasons:
            reasons[rec["Mistake reason"]] += 1
    return by_type, reasons


def accuracy(stats):
    return stats["ok"] / stats["n"] if stats["n"] else None


def pick_mix(by_type):
    """Base mix, or one swap toward the weakest RC type (3+ attempts, below 60%)."""
    weak = [t for t in RC_TYPES
            if by_type[t]["n"] >= WEAK_MIN_ATTEMPTS and accuracy(by_type[t]) < WEAK_ACCURACY]
    if not weak:
        return list(BASE_MIX), None
    weak.sort(key=lambda t: (accuracy(by_type[t]), -by_type[t]["n"]))
    weak_type = weak[0]
    # Drop one question of the strongest type with enough data; never Main idea.
    droppable = [t for t in dict.fromkeys(BASE_MIX)
                 if t not in (weak_type, "Main idea") and by_type[t]["n"] >= WEAK_MIN_ATTEMPTS]
    if droppable:
        drop = max(droppable, key=lambda t: (accuracy(by_type[t]), by_type[t]["n"]))
    else:
        drop = "Inference" if weak_type != "Inference" else "Detail"
    mix = list(BASE_MIX)
    mix.remove(drop)
    mix.append(weak_type)
    mix.sort(key=RC_TYPES.index)
    s = by_type[weak_type]
    note = (f"{weak_type} is weak ({s['ok']}/{s['n']} = {accuracy(s):.0%}), so one {drop} "
            f"question is swapped for an extra {weak_type} question.")
    return mix, note


def pick_q7(by_type):
    counts = {t: by_type[t]["n"] for t in VA_ORDER}
    low = min(counts.values())
    return next(t for t in VA_ORDER if counts[t] == low), counts


def build_settings(wb):
    log_rows = read_rows(wb["Log"], LOG_COLS)
    by_type, reasons = summarize(log_rows)
    mix, note = pick_mix(by_type)
    q7, counts = pick_q7(by_type)
    lines = ["- Q1-Q6: " + ", ".join(f"Q{i} {t}" for i, t in enumerate(mix, 1)) + ".",
             "- Swap: " + (note or f"none (no question type has {WEAK_MIN_ATTEMPTS}+ attempts "
                                   f"with accuracy below {WEAK_ACCURACY:.0%})."),
             f"- Q7: {q7} (rows in my log so far: "
             + ", ".join(f"{t} {n}" for t, n in counts.items()) + "; ties go in that order)."]
    total = len(log_rows)
    if total:
        ok = sum(r.get("Result") == "Correct" for r in log_rows)
        lines.append(f"- My log so far: {total} questions, {ok / total:.0%} correct.")
        top = sorted(((n, r) for r, n in reasons.items() if n), reverse=True)[:3]
        if top:
            lines.append("- My most common mistakes: " + ", ".join(f"{r} ({n})" for n, r in top)
                         + ". Build some wrong options that test exactly these.")
    else:
        lines.append("- My log so far: empty (this is my first day).")
    topics = [str(d["Topic"]) for d in read_rows(wb["Daily"], DAILY_COLS)[-10:] if d.get("Topic")]
    if topics:
        lines.append("- Recent passage topics (pick a different one): "
                     + "; ".join(reversed(topics)) + ".")
    return "\n".join(lines)


# ---------- commands ----------

def cmd_brief(args):
    wb = open_book(args.file)
    settings = build_settings(wb)
    today = dt.date.today().isoformat()
    print(f"Settings for {today}:\n{settings}\n")
    if not PROMPT_TEMPLATE.exists():
        sys.exit(f"Missing {PROMPT_TEMPLATE}; paste the settings above into your prompt by hand.")
    text = PROMPT_TEMPLATE.read_text(encoding="utf-8")
    text = text.split(PROMPT_MARKER, 1)[-1].strip()
    text = text.replace("{{DATE}}", today).replace("{{SETTINGS}}", settings)
    PROMPT_OUT.write_text(text + "\n", encoding="utf-8")
    print(f"Wrote {PROMPT_OUT.name}. Open it, copy all of it, and paste it into a new chat.")


def timing_path(day):
    return RECORDS / f"timing_{day.isoformat()}.json"


def times_line(rec):
    parts = [f"reading {mmss(rec['reading_sec'])} ({rec['reading_sec']} s)"]
    for i, s in enumerate(rec["q_sec"], 1):
        parts.append(f"Q{i} {s} s" + (" (over 2 min)" if s > SLOW_SEC else ""))
    parts.append(f"total {rec['total_sec'] / 60:.1f} min")
    return "TIMES: " + " | ".join(parts)


def cmd_timer(args):
    steps = (["you START reading (send START)", "you are DONE reading (send DONE)"]
             + [f"you send your answer to Q{i}" for i in range(1, 8)])
    print("Stopwatch. Press Enter at each step. Type u then Enter to undo the last step,\n"
          "or q then Enter to stop early.\n")
    marks = []
    while len(marks) < len(steps):
        try:
            cmd = input(f"Press Enter when {steps[len(marks)]}: ").strip().lower()
        except EOFError:
            break
        if cmd == "u":
            if marks:
                marks.pop()
                print("  Undone.")
            continue
        if cmd == "q":
            break
        marks.append(time.time())
        if len(marks) >= 2:
            print(f"  +{mmss(marks[-1] - marks[-2])}")
    if len(marks) < 2:
        sys.exit("Stopped before DONE, so nothing was saved.")
    today = dt.date.today()
    rec = {"date": today.isoformat(),
           "marks_epoch": [round(m, 1) for m in marks],
           "reading_sec": round(marks[1] - marks[0]),
           "q_sec": [round(marks[i] - marks[i - 1]) for i in range(2, len(marks))],
           "total_sec": round(marks[-1] - marks[0])}
    RECORDS.mkdir(exist_ok=True)
    timing_path(today).write_text(json.dumps(rec, indent=2), encoding="utf-8")
    print(f"\n{times_line(rec)}\n\nPaste the TIMES line into the chat after Q7. "
          f"Saved to {timing_path(today)}")


def read_block(args):
    if args.from_file:
        return Path(args.from_file).read_text(encoding="utf-8")
    print("Paste the LOG BLOCK from the chat (it ends with the END line):")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
            break
        lines.append(line)
        if BLOCK_END in line:
            break
    return "\n".join(lines)


def parse_block(text):
    data = {"date": None, "topic": "", "words": None, "log": [], "daily": None, "vocab": []}
    for raw in text.splitlines():
        line = re.sub(r"^\s*(?:[-*]\s+)?", "", raw).strip().strip("`").strip()
        m = re.match(r"^(DATE|TOPIC|WORDS|LOG|DAILY|VOCAB)\s*:\s*(.*)$", line, re.I)
        if not m:
            continue
        key, rest = m.group(1).upper(), m.group(2).strip()
        cells = [c.strip() for c in rest.split("|")]
        if key == "DATE":
            data["date"] = to_date(rest)
        elif key == "TOPIC":
            data["topic"] = rest
        elif key == "WORDS":
            data["words"] = num(rest, "WORDS")
        elif key == "LOG":
            data["log"].append(cells)
        elif key == "DAILY":
            data["daily"] = cells
        elif key == "VOCAB":
            data["vocab"].append(cells)
    return data


def build_records(data, timing):
    errors, log_rows, vocab_rows = [], [], []
    for i, cells in enumerate(data["log"], 1):
        cells = (cells + [""] * 8)[:8]
        try:
            qno = num(cells[0], "Q no.")
            if qno is None:
                raise ValueError("Q no. is missing")
            qtype = canon(cells[1], Q_TYPES, "Q type")
            result = canon(cells[2], RESULTS, "Result")
            secs = num(cells[3], "Time (sec)")
            if secs is None and timing and 0 < qno <= len(timing.get("q_sec", [])):
                secs = timing["q_sec"][int(qno) - 1]
            reason = None
            if result == "Wrong":
                if cells[4] in ("", "-"):
                    raise ValueError("a wrong answer needs a Mistake reason")
                reason = canon(cells[4], REASONS, "Mistake reason")
            log_rows.append({"Date": data["date"], "Q type": qtype, "Result": result,
                             "Time (sec)": secs, "Mistake reason": reason,
                             "My answer": cells[5].strip("- ") or None,
                             "Correct answer": cells[6].strip("- ") or None,
                             "Topic": data["topic"] or None, "Q no.": qno,
                             "Notes": None if cells[7] in ("", "-") else cells[7]})
        except ValueError as e:
            errors.append(f"LOG line {i}: {e}")

    d = ((data["daily"] or []) + [""] * 5)[:5]
    daily = {}
    try:
        score, out_of, reading, speed, total = (num(d[0], "Score"), num(d[1], "Out of"),
                                                num(d[2], "Reading (min)"), num(d[3], "Speed (wpm)"),
                                                num(d[4], "Total time (min)"))
        if score is None:
            score = sum(r["Result"] == "Correct" for r in log_rows)
        if out_of is None:
            out_of = len(log_rows)
        if reading is None and timing:
            reading = round(timing["reading_sec"] / 60, 1)
        if total is None and timing:
            total = round(timing["total_sec"] / 60, 1)
        if speed is None and data["words"] and reading:
            speed = round(data["words"] / reading)
        daily = {"Date": data["date"], "Score": score, "Out of": out_of, "Reading (min)": reading,
                 "Speed (wpm)": speed, "Total time (min)": total,
                 "Topic": data["topic"] or None, "Words": data["words"]}
    except ValueError as e:
        errors.append(f"DAILY line: {e}")

    source = f"VARC {data['date'].isoformat()}" + (f" - {data['topic']}" if data["topic"] else "")
    for i, cells in enumerate(data["vocab"], 1):
        cells = (cells + ["", "", ""])[:3]
        if not cells[0] or cells[0] == "-":
            errors.append(f"VOCAB line {i}: the word is missing")
            continue
        vocab_rows.append({"Word": cells[0], "Meaning": cells[1] or None, "Status": "New",
                           "Example": cells[2] or None, "Added": data["date"], "Source": source})
    return log_rows, daily, vocab_rows, errors


def cmd_log(args):
    data = parse_block(read_block(args))
    if not data["log"]:
        sys.exit("No LOG lines found. Paste the whole block, from START to END.")
    if data["date"] is None:
        data["date"] = dt.date.today()
        print(f"No DATE line, so using today ({data['date']}).")
    timing = None
    if timing_path(data["date"]).exists():
        timing = json.loads(timing_path(data["date"]).read_text(encoding="utf-8"))
    log_rows, daily, vocab_rows, errors = build_records(data, timing)
    if errors:
        print("Nothing was written. Fix these lines and run log again:")
        print("\n".join("  - " + e for e in errors))
        sys.exit(1)

    wb = open_book(args.file)
    already = [r for r in read_rows(wb["Log"], LOG_COLS)
               if to_date(r.get("Date")) == data["date"]
               and str(r.get("Topic") or "").strip().lower() == data["topic"].strip().lower()]
    if already and not ask_yes(f"The Log already has {len(already)} rows for {data['date']} "
                               f"'{data['topic']}'. Add these rows anyway?", args.yes):
        sys.exit("Stopped. Nothing was written.")

    RECORDS.mkdir(exist_ok=True)
    (RECORDS / "backups").mkdir(exist_ok=True)
    backup = RECORDS / "backups" / f"{args.file.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(args.file, backup)

    done = [("Log", *append_rows(wb["Log"], LOG_COLS, log_rows, date_cols=("Date",)))]
    if daily:
        done.append(("Daily", *append_rows(wb["Daily"], DAILY_COLS, [daily], date_cols=("Date",))))
    if vocab_rows:
        done.append(("Vocab", *append_rows(wb["Vocab"], VOCAB_COLS, vocab_rows,
                                           date_cols=("Added",))))
    try:
        wb.save(args.file)
    except PermissionError:
        sys.exit(f"Could not save {args.file.name}: close it in Excel and run log again. "
                 "Nothing was written.")
    for tab, first, last in done:
        print(f"{tab}: added rows {first}-{last}" if last > first else f"{tab}: added row {first}")
    print(f"Backup of the previous file: {backup}")
    if daily:
        print(f"Today: {daily['Score']}/{daily['Out of']}"
              + (f", reading {daily['Reading (min)']} min" if daily["Reading (min)"] else "")
              + (f" at {daily['Speed (wpm)']} wpm" if daily["Speed (wpm)"] else "") + ".")


def cmd_stats(args):
    wb = open_book(args.file)
    rows = read_rows(wb["Log"], LOG_COLS)
    by_type, reasons = summarize(rows)
    days = len(read_rows(wb["Daily"], DAILY_COLS))
    total = len(rows)
    ok = sum(r.get("Result") == "Correct" for r in rows)
    print(f"Questions: {total}   Accuracy: {ok / total:.0%}   Days practised: {days}"
          if total else f"Questions: 0   Days practised: {days}")
    cutoff = dt.date.today() - dt.timedelta(days=6)
    recent = [r for r in rows if (to_date(r.get("Date")) or dt.date.min) >= cutoff]
    if recent:
        rok = sum(r.get("Result") == "Correct" for r in recent)
        print(f"Last 7 days: {rok}/{len(recent)} = {rok / len(recent):.0%}")
    print(f"\n{'Question type':<20}{'Tried':>6}{'Right':>7}{'Acc':>7}{'Avg s':>7}")
    for t, s in by_type.items():
        acc = f"{accuracy(s):.0%}" if s["n"] else "-"
        avg = f"{sum(s['secs']) / len(s['secs']):.0f}" if s["secs"] else "-"
        print(f"{t:<20}{s['n']:>6}{s['ok']:>7}{acc:>7}{avg:>7}")
    eligible = [t for t, s in by_type.items() if s["n"] >= WEAK_MIN_ATTEMPTS]
    if eligible:
        weakest = min(eligible, key=lambda t: accuracy(by_type[t]))
        print(f"\nWeakest type (3+ attempts): {weakest}")
    if any(reasons.values()):
        print("\nWhy I lose marks:")
        for r, n in sorted(reasons.items(), key=lambda kv: -kv[1]):
            if n:
                print(f"  {r:<18}{n:>4}")


def main():
    p = argparse.ArgumentParser(description="Offline helper for the daily CAT VARC routine.")
    p.add_argument("--file", type=Path, default=LOG_FILE,
                   help="mistake log .xlsx (default: CAT_VARC_Mistake_Log.xlsx next to this script)")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("brief", help="pick today's question mix and write prompt_today.txt")
    sub.add_parser("timer", help="stopwatch for START, DONE and each question")
    lg = sub.add_parser("log", help="add the debrief's LOG BLOCK to the Excel log")
    lg.add_argument("--from-file", help="read the LOG BLOCK from a text file instead of pasting")
    lg.add_argument("--yes", action="store_true", help="don't ask before adding a repeated day")
    sub.add_parser("stats", help="print the Summary numbers")
    args = p.parse_args()
    {"brief": cmd_brief, "timer": cmd_timer, "log": cmd_log, "stats": cmd_stats}[args.cmd](args)


if __name__ == "__main__":
    main()
