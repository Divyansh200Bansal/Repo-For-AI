#!/usr/bin/env python3
"""
daily.py - the whole morning CAT VARC routine in one command:

    python daily.py

1. Fetches the last 24 hours of finance headlines from free RSS feeds.
2. Asks Claude for 5 stories (3 India, 2 global), an original CAT-style passage and
   7 questions, using the question mix your mistake log calls for.
3. Runs the quiz in this window and times your reading and every answer.
4. Asks Claude for the debrief: score, reasons, grammar fixes, words, what to sharpen.
5. Adds the rows to the Log, Daily and Vocab tabs of your Excel mistake log.

Options:  --model claude-haiku-5-5   far cheaper (see README for what each model costs)
          --effort medium            less thinking, lower cost
          --new                      make a fresh set even if today's set exists

If you stop half-way, run it again: it continues today's set where you stopped.
Needs Python 3.10+, `pip install anthropic openpyxl`, and an Anthropic API key in the
ANTHROPIC_API_KEY environment variable. Today's files go in days/<date>/.
"""
import argparse
import datetime as dt
import email.utils
import html
import json
import os
import re
import shutil
import sys
import textwrap
import time
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

import varc  # varc.py in this folder: settings, labels and Excel writing

try:
    import anthropic
except ImportError:
    sys.exit("The anthropic library is missing. Install it with:  pip install anthropic")

DAYS = varc.HERE / "days"
IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
DEFAULT_MODEL = os.environ.get("VARC_MODEL", "claude-opus-5-5")
DEFAULT_EFFORT = os.environ.get("VARC_EFFORT", "high")
PRICES = {  # US dollars per million tokens (input, output), for the cost line
    "claude-opus-5-5": (4.00, 20.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-5-5": (0.10, 0.50),
}
USD_TO_INR = 93  # approximate; edit when the rate moves
# Models that take server-side refusal fallbacks, and models that take an effort setting.
FALLBACK_MODELS = {"claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5",
                   "claude-fable-5-2", "claude-fable-5-1"}
EFFORT_PREFIXES = ("claude-opus-5", "claude-sonnet-5", "claude-haiku-5", "claude-fable-5",
                   "claude-opus-4-8", "claude-opus-4-7", "claude-opus-4-6", "claude-sonnet-4-6")
NUMBERED = ("Para-jumble", "Odd sentence out")  # answered with digits, not letters
WIDTH = 92

# Free RSS feeds. A feed that fails is skipped; edit this list if one stops working.
FEEDS = {
    "India": [
        ("Economic Times Markets", "https://economictimes.indiatimes.com/markets/rssfeeds/1977021501.cms"),
        ("Economic Times Economy", "https://economictimes.indiatimes.com/news/economy/rssfeeds/1373380680.cms"),
        ("Mint Markets", "https://www.livemint.com/rss/markets"),
        ("Mint Economy", "https://www.livemint.com/rss/economy"),
        ("Business Standard Markets", "https://www.business-standard.com/rss/markets-106.rss"),
        ("Google News India Business",
         "https://news.google.com/rss/headlines/section/topic/BUSINESS?hl=en-IN&gl=IN&ceid=IN:en"),
    ],
    "Global": [
        ("BBC Business", "https://feeds.bbci.co.uk/news/business/rss.xml"),
        ("CNBC Economy", "https://www.cnbc.com/id/20910258/device/rss/rss.html"),
        ("CNBC Finance", "https://www.cnbc.com/id/10000664/device/rss/rss.html"),
        ("Google News World Business",
         "https://news.google.com/rss/headlines/section/topic/BUSINESS?hl=en-US&gl=US&ceid=US:en"),
    ],
}

SYSTEM_PACK = """You write one day's CAT VARC practice set for a working engineer in India preparing for CAT 2027. The user message gives today's date, settings from the student's mistake log, and finance and economy headlines from the last day, each with an id (I = India, G = global). Return JSON that matches the schema.

news
- Pick the 5 most important stories for someone following markets and the economy: exactly 3 India ids and 2 global ids, India first.
- headline: your own short headline. what_happened: one sentence. why_it_matters: one sentence.
- Use only facts in that item's headline and summary. Do not add numbers, names, causes or events the item does not state; if an item is thin, keep the sentence general rather than guess. Never copy the source's wording.

passage
- An original 450-550 word passage in the style of CAT reading comprehension: an essay on the debate or idea behind one of the 5 stories, not a news report. Do not retell the news item.
- Dense and argumentative, with a clear authorial stance, some abstract reasoning, and at least one opposing view that the author answers. 4-7 paragraphs. Count the words before you answer.
- topic: 2-6 words naming that idea. passage_title: a short title. Avoid the recent topics listed in the settings.

questions
- Exactly 7, numbered 1-7, with the types in the settings, in that order. Every question must have exactly one defensible answer.
- Q1-Q6 are on the passage, with exactly 4 options each. Write option text without letters; answer is A, B, C or D. Spread the correct letters across A-D.
  - Main idea: the central argument. Inference: what the passage implies but does not state. Detail: what the passage states (an EXCEPT question is fine). Tone: the author's attitude, with close options. Vocab in context: a word the passage uses in a less common sense, with its everyday meaning as one of the traps. Strengthen/Weaken: which option, if true, would most strengthen or weaken the author's argument.
  - Make wrong options tempting the way CAT does: extreme words (always, never, only), out of scope, partly true, reversed cause and effect, distortions of the passage.
- Q7 is the verbal ability question named in the settings, on a new short text of your own, not the passage:
  - Para-jumble: options are 4 sentences in jumbled order, numbered 1-4 by position; answer is the correct order as digits, such as "3142". Pronouns, linking words and logic must fix one order, and it must not be 1234.
  - Para summary: the stem contains a 4-6 sentence paragraph; 4 summary options; answer A-D.
  - Odd sentence out: options are 5 sentences; 4 of them form a coherent paragraph; answer is the number (1-5) of the odd one.
- explanation: one or two sentences on why the answer is right. traps: a short reason for each wrong option, such as "A: extreme ('never'); B: out of scope; D: partly true".

vocab
- 5 useful words that appear in your passage, each with a plain meaning and a new example sentence."""

SYSTEM_DEBRIEF = """You are a strict, honest CAT VARC coach. The user message has a passage, 7 questions with the answer key, and one student's answers. For each question the student wrote their own answer first, then picked an option, then gave a reason. Return JSON that matches the schema. No flattery.

questions: one entry per question, numbered 1-7.
- mistake_reason: for a wrong answer, the label that best explains the error, judged from the student's own answer and reason: Misread question, Extreme option, Out of scope, Over-inference, Partial answer, Vocab gap, Missed link (PJ), Time pressure, Careless. For a correct answer, "None".
- comment: one or two sentences on the student's reasoning: what was sound, where it went wrong, and what to notice next time. Be specific.

grammar: every grammar mistake in the student's own words (own answers and reasons). you_wrote quotes the exact words; corrected gives the corrected version; rule names the rule in a few words. Watch especially for Hindi-influenced errors: missing or extra articles (a, an, the), wrong word order in questions, and uncountable nouns used with "a". Skip one-letter option answers. Return an empty list if there are none.

sharpen_tomorrow: the one thing to work on tomorrow, in one or two sentences."""


def _obj(props):
    return {"type": "object", "properties": props, "required": list(props),
            "additionalProperties": False}


STR = {"type": "string"}
PACK_SCHEMA = _obj({
    "news": {"type": "array", "items": _obj({"id": STR, "headline": STR,
                                             "what_happened": STR, "why_it_matters": STR})},
    "topic": STR,
    "passage_title": STR,
    "passage_paragraphs": {"type": "array", "items": STR},
    "questions": {"type": "array", "items": _obj({
        "number": {"type": "integer"},
        "type": {"type": "string", "enum": varc.Q_TYPES},
        "stem": STR,
        "options": {"type": "array", "items": STR},
        "answer": STR,
        "explanation": STR,
        "traps": STR})},
    "vocab": {"type": "array", "items": _obj({"word": STR, "meaning": STR, "example": STR})},
})
DEBRIEF_SCHEMA = _obj({
    "questions": {"type": "array", "items": _obj({
        "number": {"type": "integer"},
        "mistake_reason": {"type": "string", "enum": varc.REASONS + ["None"]},
        "comment": STR})},
    "grammar": {"type": "array", "items": _obj({"you_wrote": STR, "corrected": STR, "rule": STR})},
    "sharpen_tomorrow": STR,
})


# ---------- news from RSS feeds ----------

def _clean(text, limit=None):
    text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", text or ""))).strip()
    if limit and len(text) > limit:
        text = text[:limit].rsplit(" ", 1)[0] + "..."
    return text


def _when(stamp):
    stamp = (stamp or "").strip()
    if not stamp:
        return None
    try:
        when = email.utils.parsedate_to_datetime(stamp)
    except (TypeError, ValueError):
        try:
            when = dt.datetime.fromisoformat(stamp.replace("Z", "+00:00"))
        except ValueError:
            return None
    return when if when.tzinfo else when.replace(tzinfo=dt.timezone.utc)


def parse_feed(data, feed_name):
    """Items from an RSS 2.0 or Atom feed: title, link, when, summary, source."""
    root = ET.fromstring(data)
    atom = "{http://www.w3.org/2005/Atom}"
    items = []
    for it in root.iter("item"):
        items.append({"title": it.findtext("title", ""), "link": it.findtext("link", ""),
                      "when": _when(it.findtext("pubDate", "")),
                      "summary": it.findtext("description", ""),
                      "source": it.findtext("source", "") or feed_name})
    for it in root.iter(atom + "entry"):
        link = it.find(atom + "link")
        items.append({"title": it.findtext(atom + "title", ""),
                      "link": link.get("href", "") if link is not None else "",
                      "when": _when(it.findtext(atom + "updated", "")
                                    or it.findtext(atom + "published", "")),
                      "summary": it.findtext(atom + "summary", ""), "source": feed_name})
    return items


def fetch_feeds():
    raw, failed = {}, []
    for region, feeds in FEEDS.items():
        raw[region] = []
        for name, url in feeds:
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (varc-daily)"})
                with urllib.request.urlopen(req, timeout=15) as resp:
                    raw[region].append(parse_feed(resp.read(), name))
            except Exception as e:  # one broken feed must not stop the run
                failed.append(f"{name} ({type(e).__name__})")
    return raw, failed


def pick_headlines(raw, now, hours, per_region=35):
    """Recent, de-duplicated headlines per region, taking turns between feeds."""
    picked = {}
    for region, feeds in raw.items():
        seen, lists = set(), []
        for items in feeds:
            fresh = []
            for it in items:
                title, source = _clean(it["title"]), _clean(it["source"])
                if source and title.endswith(" - " + source):
                    title = title[: -len(source) - 3]
                key = re.sub(r"[^a-z0-9]", "", title.lower())[:70]
                when = it["when"]
                if not title or key in seen or when is None:
                    continue
                if not now - dt.timedelta(hours=hours) <= when <= now + dt.timedelta(hours=1):
                    continue
                seen.add(key)
                summary = _clean(it["summary"], 240)
                if re.sub(r"[^a-z0-9]", "", summary.lower()).startswith(key[:40]):
                    summary = ""  # Google News repeats the headline as its summary
                fresh.append({"title": title, "source": source, "link": it["link"].strip(),
                              "when": when, "summary": summary, "region": region})
            lists.append(sorted(fresh, key=lambda x: x["when"], reverse=True))
        merged = []
        while len(merged) < per_region and any(lists):
            for lst in lists:
                if lst and len(merged) < per_region:
                    merged.append(lst.pop(0))
        picked[region] = merged
    return picked


def gather_news(now):
    print("Fetching headlines from the last 24 hours...", flush=True)
    raw, failed = fetch_feeds()
    for hours in (24, 36, 48):
        picked = pick_headlines(raw, now, hours)
        if len(picked["India"]) >= 3 and len(picked["Global"]) >= 2:
            break
    if failed:
        print("Skipped feeds that did not load: " + ", ".join(failed))
    if len(picked["India"]) < 3 or len(picked["Global"]) < 2:
        sys.exit("Not enough headlines (need 3 India and 2 global). Check your internet connection, "
                 "or edit FEEDS at the top of daily.py.")
    if hours > 24:
        print(f"Few stories in the last 24 hours, so this uses the last {hours} hours.")
    items = []
    for prefix, region in (("I", "India"), ("G", "Global")):
        for i, it in enumerate(picked[region], 1):
            items.append(dict(it, id=f"{prefix}{i}"))
    print(f"Got {len(picked['India'])} India and {len(picked['Global'])} global headlines.")
    return items, hours


# ---------- Claude ----------

def call_claude(client, model, effort, system, user, schema, label, usage):
    """One request with a JSON-schema reply. Adds the tokens used to `usage`."""
    kwargs = {"model": model, "max_tokens": 64000, "system": system,
              "messages": [{"role": "user", "content": user}],
              "output_config": {"format": {"type": "json_schema", "schema": schema}}}
    if model.startswith(EFFORT_PREFIXES):
        kwargs["output_config"]["effort"] = effort
    if model in FALLBACK_MODELS:  # if the model declines, the API retries on another model
        stream = client.beta.messages.stream(betas=["server-side-fallback-2026-07-01"],
                                             fallbacks="default", **kwargs)
    else:
        stream = client.messages.stream(**kwargs)
    print(f"{label} with {model}", end="", flush=True)
    started = time.time()
    with stream as s:
        for i, _event in enumerate(s):
            if i % 200 == 0:
                print(".", end="", flush=True)
        msg = s.get_final_message()
    print(f" done in {time.time() - started:.0f} s.", flush=True)
    usage[0] += msg.usage.input_tokens or 0
    usage[1] += msg.usage.output_tokens or 0
    if msg.stop_reason == "refusal":
        raise RuntimeError("Claude declined this request. Run again, or try another --model.")
    if msg.stop_reason == "max_tokens":
        raise RuntimeError("The reply was cut off before it finished. Run again.")
    text = next((b.text for b in reversed(msg.content) if b.type == "text"), "")
    return json.loads(text)


def words_in(paragraphs):
    return len(re.findall(r"[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*", " ".join(paragraphs)))


def answer_of(q):
    ans = str(q.get("answer", ""))
    if q.get("type") in NUMBERED:
        return re.sub(r"\D", "", ans)
    m = re.search(r"\b([ABCD])\b", ans.upper())  # "C", "(C)" or "Option C" -> "C"
    return m.group(1) if m else ans.strip().upper()[:1]


def check_pack(pack, mix, q7, ids):
    problems = []
    news = pack.get("news", [])
    regions = [str(n.get("id", ""))[:1] for n in news]
    if len(news) != 5 or regions.count("I") != 3 or regions.count("G") != 2:
        problems.append("news needs exactly 5 items: 3 India ids (I...) and 2 global ids (G...)")
    unknown = [n.get("id") for n in news if n.get("id") not in ids]
    if unknown:
        problems.append("these news ids are not in the list: " + ", ".join(map(str, unknown)))
    if len({n.get("id") for n in news}) != len(news):
        problems.append("each news item needs a different id")
    words = words_in(pack.get("passage_paragraphs", []))
    if not 430 <= words <= 570:
        problems.append(f"the passage has {words} words; it must have 450-550")
    qs = pack.get("questions", [])
    if [q.get("type") for q in qs] != mix + [q7]:
        problems.append("the questions must be, in order: " + ", ".join(mix + [q7]))
    if [q.get("number") for q in qs] != list(range(1, 8)):
        problems.append("the questions must be numbered 1 to 7")
    for q in qs:
        n, opts, ans = q.get("number"), q.get("options", []), answer_of(q)
        if q.get("type") == "Para-jumble":
            order = [str(i) for i in range(1, len(opts) + 1)]
            if len(opts) not in (4, 5) or sorted(ans) != order or ans == "".join(order):
                problems.append(f"Q{n}: a para-jumble needs 4 or 5 sentences and an answer that "
                                "reorders their numbers (not already in order)")
        elif q.get("type") == "Odd sentence out":
            if len(opts) != 5 or ans not in list("12345"):
                problems.append(f"Q{n}: odd sentence out needs 5 sentences and the odd one's number")
        elif len(opts) != 4 or ans not in list("ABCD"):
            problems.append(f"Q{n} needs exactly 4 options and an answer A, B, C or D")
    if len(pack.get("vocab", [])) != 5:
        problems.append("vocab needs exactly 5 words")
    return problems, words


def make_pack(client, args, settings, mix, q7, items, hours, today, usage):
    lines = []
    for it in items:
        when = it["when"].astimezone(IST).strftime("%d %b %H:%M IST")
        lines.append(f'{it["id"]} | {it["source"]} | {when} | {it["title"]}'
                     + (f' | {it["summary"]}' if it["summary"] else ""))
    user = (f"Date: {today:%d %B %Y} (India time)\n\nSettings from my mistake log:\n{settings}\n\n"
            f"Headlines from the last {hours} hours (id | source | time | headline | summary):\n"
            + "\n".join(lines))
    by_id = {it["id"]: it for it in items}
    for attempt in (1, 2):
        pack = call_claude(client, args.model, args.effort, SYSTEM_PACK, user, PACK_SCHEMA,
                           "Writing today's news, passage and questions", usage)
        problems, words = check_pack(pack, mix, q7, set(by_id))
        if not problems:
            break
        print("That set had problems: " + "; ".join(problems))
        if attempt == 2:
            sys.exit("The set still had problems after a second try, so nothing was saved. "
                     "Run daily.py again.")
        print("Asking for a corrected set.")
        user += ("\n\nYour previous answer had these problems; fix all of them and return the "
                 "whole set again: " + "; ".join(problems))
    if not 450 <= words <= 550:
        print(f"Note: the passage has {words} words (target 450-550).")
    for n in pack["news"]:
        src = by_id[n["id"]]
        n.update(region=src["region"], source=src["source"], link=src["link"])
    opt_prefix = re.compile(r"^\s*\(?(?:[A-Da-d]|[1-5])[.):]\s+")
    for q in pack["questions"]:
        q["options"] = [opt_prefix.sub("", o).strip() for o in q["options"]]
        q["answer"] = answer_of(q)
    pack.update(date=today.isoformat(), words=words, model=args.model,
                made=dt.datetime.now(IST).isoformat(timespec="seconds"))
    return pack


def debrief_request(pack, answers):
    lines = [f"Passage: {pack['passage_title']}", *pack["passage_paragraphs"], "",
             "Questions, key and the student's answers:"]
    for q, a in zip(pack["questions"], answers):
        lines += [f"Q{q['number']} ({q['type']}): {q['stem']}"]
        lines += [f"  {lab}. {opt}" for lab, opt in zip(labels(q), q["options"])]
        lines += [f"  Key: {q['answer']}. {q['explanation']} Traps: {q['traps']}",
                  f"  Student's own answer: {a['own']}",
                  f"  Student's choice: {a['choice']} ({'correct' if a['correct'] else 'wrong'}), "
                  f"{a['secs']} seconds",
                  f"  Student's reason: {a['reason']}", ""]
    return "\n".join(lines)


# ---------- showing the set and running the quiz ----------

def wrap(text, indent="", hang=None):
    """Wrap to WIDTH; continuation lines line up under the text, not under a bullet."""
    hang = " " * len(indent) if hang is None else hang
    return textwrap.fill(text, width=WIDTH, initial_indent=indent, subsequent_indent=hang)


def labels(q):
    n = len(q["options"])
    return [str(i) for i in range(1, n + 1)] if q["type"] in NUMBERED else list("ABCDE")[:n]


def news_text(pack):
    out = []
    for region in ("India", "Global"):
        out.append(f"\n{region}")
        for n in (x for x in pack["news"] if x["region"] == region):
            out += [wrap(n["headline"], "  * "), wrap("What happened: " + n["what_happened"], "    "),
                    wrap("Why it matters: " + n["why_it_matters"], "    "),
                    f"    Source: {n['source']}  {n['link']}"]
    return "\n".join(out)


def passage_text(pack):
    body = "\n\n".join(wrap(p) for p in pack["passage_paragraphs"])
    return f"{pack['passage_title']}  ({pack['words']} words)\n\n{body}"


def question_text(q):
    lines = [wrap(f"Q{q['number']} ({q['type']}). {q['stem']}")]
    lines += [wrap(f"{lab}. {opt}", "   ") for lab, opt in zip(labels(q), q["options"])]
    return "\n".join(lines)


def banner(title):
    print("\n" + "=" * WIDTH + f"\n{title}\n" + "=" * WIDTH)


def wait(prompt):
    try:
        input(prompt)
    except EOFError:
        sys.exit("\nStopped. Run daily.py again to continue today's set.")


def ask(prompt, check=None, hint="Please type something."):
    while True:
        try:
            text = input(prompt).strip()
        except EOFError:
            sys.exit("\nStopped. Your answers so far are saved; run daily.py again to continue.")
        if text and (check is None or check(text)):
            return text
        print("  " + hint)


def read_choice(q, text):
    if q["type"] in NUMBERED:
        return re.sub(r"\D", "", text)
    m = re.fullmatch(r"\(?\s*([A-Da-d])\s*[).]?", text.strip())
    return m.group(1).upper() if m else ""


def choice_ok(q, text):
    c, n = read_choice(q, text), len(q["options"])
    if q["type"] == "Para-jumble":
        return sorted(c) == [str(i) for i in range(1, n + 1)]
    if q["type"] == "Odd sentence out":
        return len(c) == 1 and 1 <= int(c) <= n
    return c in list("ABCD")


def run_quiz(pack, progress, save):
    banner(f"FINANCE NEWS - {pack['date']}")
    print(news_text(pack))
    banner("PASSAGE")
    if "reading_sec" not in progress:
        wait("Press Enter to START reading (the passage appears and the clock starts) ")
        start = time.time()
        print("\n" + passage_text(pack))
        wait("\nPress Enter when you are DONE reading ")
        progress["reading_sec"] = round(time.time() - start)
        save()
        print(f"Reading time: {varc.mmss(progress['reading_sec'])}.")
    else:
        print("\n" + passage_text(pack))
        print(f"\n(Continuing today's set. Reading time already saved: "
              f"{varc.mmss(progress['reading_sec'])}.)")
    answers = progress.setdefault("answers", [])
    prev = time.time()
    for q in pack["questions"][len(answers):]:
        banner(f"QUESTION {q['number']} OF 7")
        print(question_text(q) + "\n")
        own = ask("Your own answer first, in your words: ")
        if q["type"] == "Para-jumble":
            prompt, hint = "Your order (for example 3142): ", "Type each sentence number once."
        elif q["type"] == "Odd sentence out":
            prompt, hint = "Number of the odd sentence: ", f"Type one number from 1 to {len(q['options'])}."
        else:
            prompt, hint = "Your option (A, B, C or D): ", "Type one letter: A, B, C or D."
        choice = read_choice(q, ask(prompt, lambda t: choice_ok(q, t), hint))
        reason = ask("Your reason: ")
        now = time.time()
        answers.append({"number": q["number"], "own": own, "choice": choice, "reason": reason,
                        "secs": round(now - prev), "correct": choice == q["answer"]})
        prev = now
        save()


# ---------- debrief and log ----------

def show_debrief(pack, progress, debrief, cost):
    answers = progress["answers"]
    score = sum(a["correct"] for a in answers)
    by_no = {d["number"]: d for d in (debrief or {}).get("questions", [])}
    out = [f"DEBRIEF - {pack['date']} - {pack['topic']}", f"\nScore: {score}/7"]
    for q, a in zip(pack["questions"], answers):
        mark = "Correct" if a["correct"] else "Wrong"
        slow = "  (over 2 min)" if a["secs"] > varc.SLOW_SEC else ""
        out.append(f"\nQ{q['number']} {q['type']}: {mark}. You chose {a['choice']}, key {q['answer']}. "
                   f"{a['secs']} s{slow}")
        out.append(wrap(f"Why: {q['explanation']}", "   "))
        out.append(wrap(f"Traps: {q['traps']}", "   "))
        d = by_no.get(q["number"])
        if d:
            if not a["correct"] and d["mistake_reason"] != "None":
                out.append(f"   Mistake type: {d['mistake_reason']}")
            out.append(wrap(f"Coach: {d['comment']}", "   "))
    reading_min = progress["reading_sec"] / 60
    total_min = (progress["reading_sec"] + sum(a["secs"] for a in answers)) / 60
    wpm = round(pack["words"] / reading_min) if reading_min else 0
    slow = [f"Q{a['number']}" for a in answers if a["secs"] > varc.SLOW_SEC]
    out.append(f"\nTime: reading {varc.mmss(progress['reading_sec'])} ({wpm} wpm), total {total_min:.1f} min."
               + (f" Over 2 minutes: {', '.join(slow)}." if slow else " No question took over 2 minutes."))
    if debrief:
        out.append("\nGrammar in your answers:")
        if debrief["grammar"]:
            for g in debrief["grammar"]:
                out.append(wrap(f'You wrote: "{g["you_wrote"]}"  ->  "{g["corrected"]}"  ({g["rule"]})', "  - "))
        else:
            out.append("  No grammar mistakes found.")
    out.append("\nNew words:")
    for v in pack["vocab"]:
        out.append(wrap(f"{v['word']}: {v['meaning']}. Example: {v['example']}", "  - "))
    if debrief:
        out.append("\n" + wrap("Sharpen tomorrow: " + debrief["sharpen_tomorrow"]))
    out.append("\n" + cost)
    return "\n".join(out), score, reading_min, total_min, wpm


def log_results(path, pack, progress, debrief, score, reading_min, total_min, wpm):
    wb = varc.open_book(path)
    day = dt.date.fromisoformat(pack["date"])
    already = [r for r in varc.read_rows(wb["Log"], varc.LOG_COLS)
               if varc.to_date(r.get("Date")) == day
               and str(r.get("Topic") or "").strip().lower() == pack["topic"].strip().lower()]
    if already:
        print(f"The Log already has rows for {day} '{pack['topic']}', so nothing new was added.")
        return
    by_no = {d["number"]: d for d in (debrief or {}).get("questions", [])}
    log_rows = []
    for q, a in zip(pack["questions"], progress["answers"]):
        d = by_no.get(q["number"], {})
        reason = None if a["correct"] else d.get("mistake_reason")
        log_rows.append({"Date": day, "Q type": q["type"], "Result": "Correct" if a["correct"] else "Wrong",
                         "Time (sec)": a["secs"], "Mistake reason": None if reason == "None" else reason,
                         "My answer": a["choice"], "Correct answer": q["answer"], "Topic": pack["topic"],
                         "Q no.": q["number"], "Notes": (d.get("comment") or "")[:250] or None})
    daily = {"Date": day, "Score": score, "Out of": 7, "Reading (min)": round(reading_min, 1),
             "Speed (wpm)": wpm, "Total time (min)": round(total_min, 1), "Topic": pack["topic"],
             "Words": pack["words"]}
    vocab_rows = [{"Word": v["word"], "Meaning": v["meaning"], "Status": "New", "Example": v["example"],
                   "Added": day, "Source": f"VARC {day.isoformat()} - {pack['topic']}"}
                  for v in pack["vocab"]]
    (varc.RECORDS / "backups").mkdir(parents=True, exist_ok=True)
    backup = varc.RECORDS / "backups" / f"{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy2(path, backup)
    done = [("Log", *varc.append_rows(wb["Log"], varc.LOG_COLS, log_rows, date_cols=("Date",))),
            ("Daily", *varc.append_rows(wb["Daily"], varc.DAILY_COLS, [daily], date_cols=("Date",))),
            ("Vocab", *varc.append_rows(wb["Vocab"], varc.VOCAB_COLS, vocab_rows, date_cols=("Added",)))]
    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"Could not save {path.name}: close it in Excel and run daily.py again "
                 "(your answers are saved, so it will only redo the log).")
    print(f"Logged in {path.name}: "
          + ", ".join(f"{t} row {a}" if a == b else f"{t} rows {a}-{b}" for t, a, b in done) + ".")


def cost_line(model, usage):
    tin, tout = usage
    if model not in PRICES:
        return f"This run used {tin:,} input and {tout:,} output tokens."
    usd = tin / 1e6 * PRICES[model][0] + tout / 1e6 * PRICES[model][1]
    return (f"This run used {tin:,} input and {tout:,} output tokens: about ${usd:.3f} "
            f"(about Rs {usd * USD_TO_INR:.1f}).")


def brief_markdown(pack):
    lines = [f"# VARC - {pack['date']} - {pack['topic']}", "", "## Finance news, last 24 hours"]
    for region in ("India", "Global"):
        lines.append(f"\n### {region}")
        for n in (x for x in pack["news"] if x["region"] == region):
            lines += [f"\n**{n['headline']}**", f"- What happened: {n['what_happened']}",
                      f"- Why it matters: {n['why_it_matters']}", f"- Source: [{n['source']}]({n['link']})"]
    lines += [f"\n## Passage: {pack['passage_title']} ({pack['words']} words)", ""]
    lines += [p + "\n" for p in pack["passage_paragraphs"]]
    lines.append("## Questions")
    for q in pack["questions"]:
        lines.append(f"\n**Q{q['number']} ({q['type']}).** {q['stem']}\n")
        lines += [f"- {lab}. {opt}" for lab, opt in zip(labels(q), q["options"])]
    return "\n".join(lines) + "\n"


def api_errors(fn):
    """Run fn() and turn API errors into plain advice."""
    try:
        return fn()
    except anthropic.AuthenticationError:
        sys.exit("The API key is missing or wrong. Set ANTHROPIC_API_KEY (README: 'Fully automatic').")
    except anthropic.PermissionDeniedError as e:
        sys.exit(f"Your API key can't do this: {e.message}")
    except anthropic.NotFoundError:
        sys.exit("That model name was not found. Check --model (for example claude-opus-5-5).")
    except anthropic.RateLimitError:
        sys.exit("The API is rate-limiting you. Wait a minute and run again.")
    except anthropic.BadRequestError as e:
        sys.exit(f"The API refused the request: {e.message}")  # e.g. credit balance too low
    except anthropic.APIConnectionError:
        sys.exit("Could not reach the Claude API. Check your internet connection.")
    except anthropic.APIStatusError as e:
        sys.exit(f"Claude API error {e.status_code}: {e.message}. Try again in a few minutes.")
    except anthropic.AnthropicError as e:
        sys.exit(f"Claude client problem: {e}. Is ANTHROPIC_API_KEY set?")
    except json.JSONDecodeError:
        sys.exit("Claude's reply was not valid JSON. Run daily.py again.")
    except RuntimeError as e:  # refusal or a reply cut off, raised by call_claude
        sys.exit(str(e))


def main():
    p = argparse.ArgumentParser(description="The whole daily CAT VARC routine in one command.")
    p.add_argument("--model", default=DEFAULT_MODEL, help="Claude model (default %(default)s)")
    p.add_argument("--effort", default=DEFAULT_EFFORT, choices=["low", "medium", "high", "xhigh", "max"],
                   help="how hard the model thinks while writing the set (default %(default)s)")
    p.add_argument("--new", action="store_true", help="make a fresh set even if today's set exists")
    p.add_argument("--file", type=Path, default=varc.LOG_FILE, help="mistake log .xlsx")
    args = p.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(line_buffering=True)  # show progress at once in Colab

    now = dt.datetime.now(IST)
    today = now.date()
    sets = sorted(DAYS.glob(f"{today.isoformat()}*/pack.json"), key=lambda f: f.stat().st_mtime)
    usage = [0, 0]
    client = anthropic.Anthropic(max_retries=3)

    if sets and not args.new:
        day_dir = sets[-1].parent
        pack = json.loads(sets[-1].read_text(encoding="utf-8"))
        print(f"Using today's set from {day_dir.name} (run with --new for a fresh one).")
    else:
        wb = varc.open_book(args.file)
        settings, mix, q7 = varc.build_settings(wb)
        print(f"Today's settings:\n{settings}\n")
        items, hours = gather_news(now.astimezone(dt.timezone.utc))
        pack = api_errors(lambda: make_pack(client, args, settings, mix, q7, items, hours, today, usage))
        n = 1
        day_dir = DAYS / today.isoformat()
        while day_dir.exists():
            n += 1
            day_dir = DAYS / f"{today.isoformat()}-{n}"
        day_dir.mkdir(parents=True)
        (day_dir / "pack.json").write_text(json.dumps(pack, indent=2, ensure_ascii=False), encoding="utf-8")
        (day_dir / "brief.md").write_text(brief_markdown(pack), encoding="utf-8")
        print(f"Saved today's set in {day_dir} (pack.json holds the answers: don't open it yet).")

    progress_file = day_dir / "progress.json"
    progress = json.loads(progress_file.read_text(encoding="utf-8")) if progress_file.exists() else {}
    if progress.get("logged"):
        print("You have already finished and logged this set. Run with --new for another one.")
        return

    def save():
        progress_file.write_text(json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8")

    if len(progress.get("answers", [])) < len(pack["questions"]):
        run_quiz(pack, progress, save)

    debrief = progress.get("debrief")  # saved, so a re-run never pays for it twice
    if debrief is None:
        try:
            debrief = api_errors(lambda: call_claude(
                client, args.model, "medium", SYSTEM_DEBRIEF, debrief_request(pack, progress["answers"]),
                DEBRIEF_SCHEMA, "\nWriting your debrief", usage))
            progress["debrief"] = debrief
            save()
        except SystemExit as e:  # api_errors turns API problems into SystemExit with advice
            print(f"The debrief could not be written ({e}). Your score and the key are below.")
    text, score, reading_min, total_min, wpm = show_debrief(pack, progress, debrief,
                                                            cost_line(args.model, usage))
    banner("DEBRIEF")
    print(text)
    (day_dir / "debrief.md").write_text(text + "\n", encoding="utf-8")
    log_results(args.file, pack, progress, debrief, score, reading_min, total_min, wpm)
    progress["logged"] = True
    save()


if __name__ == "__main__":
    main()
