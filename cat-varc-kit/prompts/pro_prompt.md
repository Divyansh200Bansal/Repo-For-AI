# Pro daily prompt (Claude Code cloud session)

Paste everything below the line into a new cloud session at claude.ai/code. Replace
`<SHEET_LINK>` with the link to your mistake-log Sheet first. It needs the Claude Docs
connector (for the doc and the Sheet) and web search.

Changes from the first version, learned on the first run (2026-11-20):
- Pick the topic before creating the doc: the connector cannot rename a doc's title later.
- A doc tab cannot be empty, so the Feedback tab starts with a one-line placeholder.
- Q7 ties go in a fixed order, so the three verbal-ability types rotate.
- If the session's network blocks news sites, use search results checked across two or more reports, and say so.
- The swap rule now says which question gets dropped.

---

Daily CAT finance and VARC brief

Who I am and what I want
I'm a working engineer preparing for CAT 2027 (VARC section). Each morning I want to (1) catch up on finance news, (2) practise CAT reading comprehension, and (3) fix my English, especially Hindi-influenced mistakes: missing articles, wrong word order in questions, and uncountable nouns used with "a". Be direct and honest. Don't flatter me.

Each run, do this in order

1. Finance news. Search the web for the 5 most important finance, economy and market stories of the last 24 hours: 3 from India, 2 global. For each, give a headline, two lines in your own words (what happened, why it matters) and the source link. Never copy article text. If news pages can't be opened, use search results checked across two or more reports, and say so.
2. Passage. Write an original 450-550 word CAT-style passage on the debate or idea behind one of today's stories, not a news report. Make it dense and argumentative, with a clear authorial stance. State the word count.
3. Questions. First read the Summary tab of my mistake log (link below).
   * Q1-Q6, four options each: 1 Main idea, 2 Inference, 1 Detail, 1 Tone, 1 Vocab in context. If any type has 3+ attempts and accuracy below 60%, swap one question for an extra one of that type: drop a question of my strongest type with 3+ attempts (never Main idea); if none qualifies, drop the second Inference question (or Detail, if Inference is the weak type).
   * Q7: one verbal ability question, of whichever type has the fewest rows in the Log: Para-jumble, Para summary or Odd sentence out. Break ties in that order. For para-jumbles and odd sentences, I type the numbers (e.g. 3142).
   * Make the wrong options tempting the way CAT does: extreme words, out of scope, partly true.
4. Doc. Only after choosing the topic, create a doc titled "VARC - <YYYY-MM-DD> - <topic>" with two tabs: "Today" (news, passage, questions) and "Feedback - open after Q7" (a one-line placeholder for now). Keep the answer key out of the doc; save it in a scratchpad file instead.
5. Run message. End with today's topic, the doc link, and: "Send START when you begin reading and DONE when you finish." If the doc fails, put everything in this message, with the answer key at the very end.

When I reply in this session

* Timing: when each of my messages arrives, run `date +%s` in the shell and save it to a scratchpad file. Reading time = START to DONE. Each question's time = the gap since my previous message. If the shell isn't available, use the message times.
* Answers: for each question, I'll write my own answer first, then pick an option and give my reason. Reply only "Noted" until I've answered Q7, but write your feedback into the Feedback tab as we go.
* Debrief after Q7: my score; the correct answers with reasons; my time per question (flag anything over 2 minutes); every grammar mistake in my answers, quoting my words and the corrected version; 5 new words from the passage; and the one thing to sharpen tomorrow.

Log everything
Mistake log (a Sheet): <SHEET_LINK>
Write to it with the Claude Docs connector, and read its guide (topic.sheets) first. Add rows below the last filled row. Never overwrite existing rows.

* Log tab, one row per question: Date, Q type, Result (Correct or Wrong), Time (sec), Mistake reason (only if wrong), My answer, Correct answer, Topic, Q no., Notes.
* Daily tab, one row per day: Date, Score, Out of, Reading (min), Speed (wpm), Total time (min), Topic, Words.
* Vocab tab, today's 5 words: Word, Meaning, Status (New), Example, Added, Source.

Use these labels exactly, or the Summary formulas won't count them:

* Q type: Main idea, Inference, Detail, Tone, Vocab in context, Strengthen/Weaken, Para summary, Para-jumble, Odd sentence out
* Mistake reason: Misread question, Extreme option, Out of scope, Over-inference, Partial answer, Vocab gap, Missed link (PJ), Time pressure, Careless

If any step fails, say which one in the run message and carry on with the rest.
