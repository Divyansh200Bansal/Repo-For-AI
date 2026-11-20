# Free-version daily prompt

Don't paste this file as it is. Run `python varc.py brief`: it fills in `{{DATE}}` and
`{{SETTINGS}}` from your mistake log and writes the finished prompt to `prompt_today.txt`.
Paste that into a new chat. It works on the free Claude plan; turn on web search if your
plan shows the option. Everything below the marker line is the prompt.

<!-- PROMPT STARTS BELOW -->
Daily CAT finance and VARC brief for {{DATE}} (India time).

Who I am: a working engineer preparing for CAT 2027 (VARC section). Each morning I want to (1) catch up on finance news, (2) practise CAT reading comprehension, and (3) fix my English, especially Hindi-influenced mistakes: missing articles, wrong word order in questions, and uncountable nouns used with "a". Be direct and honest. Don't flatter me.

Settings from my mistake log (follow them exactly):
{{SETTINGS}}

Do all of this in your first reply, in this order:

1. Finance news. Search the web for the 5 most important finance, economy and market stories of the last 24 hours: 3 from India, 2 global. For each, give a headline, two lines in your own words (what happened, why it matters) and the source link. Never copy article text. If you cannot search the web, say so in one line and ask me to paste five headlines. Never invent news.
2. Passage. Write an original 450-550 word CAT-style passage on the debate or idea behind one of today's stories, not a news report. Make it dense and argumentative, with a clear authorial stance. Give it a title and the exact word count.
3. Questions. Q1-Q6 exactly as listed in Settings, four options each (A-D). Q7 is one verbal ability question of the type in Settings; for para-jumbles and odd-sentence-out questions I answer with digits (e.g. 3142). Make the wrong options tempting the way CAT does: extreme words, out of scope, partly true. Label each question with its type.
4. Close with today's topic and this line: "Send START when you begin reading and DONE when you finish." Then leave several blank lines, write "ANSWER KEY - scroll down only after Q7", and give the answers with a one-line reason each. Put nothing after the answer key.

When I reply:
- I send START, then DONE, then one message per question: my own answer first, then the option, then my reason. If I am short on messages, I may send all seven answers in one message.
- Until I have answered Q7, reply only "Noted". Don't hint at answers.
- I time myself with a stopwatch script. After Q7 I paste a line that starts with "TIMES:". If I forget, ask me for it once; then carry on without times.
- After Q7, debrief: my score; the correct answers with reasons; my time per question (flag anything over 2 minutes); every grammar mistake in my messages, quoting my words and the corrected version; 5 new words from the passage; and the one thing to sharpen tomorrow.
- End the debrief with the LOG BLOCK below, filled in, inside a code block, so that my script can copy it into my Excel log. Write one LOG line per question (Q1 to Q7), one DAILY line and five VOCAB lines. Use "-" for an empty field, and never put "|" inside a field.

=== LOG BLOCK START ===
DATE: {{DATE}}
TOPIC: <today's topic in 2-6 words>
WORDS: <passage word count>
LOG: <Q no.> | <Q type> | <Correct or Wrong> | <time in seconds> | <mistake reason, or - if correct> | <my answer> | <correct answer> | <short note>
DAILY: <score> | <out of> | <reading time in minutes> | <reading speed in wpm> | <total time in minutes>
VOCAB: <word> | <meaning> | <example sentence>
=== LOG BLOCK END ===

Use these labels exactly, or my log's formulas won't count them:
- Q type: Main idea, Inference, Detail, Tone, Vocab in context, Strengthen/Weaken, Para summary, Para-jumble, Odd sentence out
- Mistake reason (wrong answers only): Misread question, Extreme option, Out of scope, Over-inference, Partial answer, Vocab gap, Missed link (PJ), Time pressure, Careless
