# CAT VARC daily kit

A 30-minute morning routine for CAT VARC: 5 finance stories, an original CAT-style passage with 7 questions, timing, grammar feedback, and a mistake log that picks the next day's questions. It runs two ways.

| Step | Pro: Claude cloud session | Free: any plan, on your laptop |
| --- | --- | --- |
| News, passage, questions | Claude writes them into a doc | Claude writes them in the chat |
| Today's question mix | Claude reads your Sheet's Summary tab | `python varc.py brief` reads your Excel log |
| Timing | Claude runs `date +%s` per message | `python varc.py timer` (press Enter at each step) |
| Logging | Claude writes rows into your Sheet | `python varc.py log` writes rows into the Excel file |

## What's in this folder

- `prompts/pro_prompt.md`: the routine for a Claude Code cloud session. Add your Sheet link, then paste it.
- `prompts/free_prompt.md`: the routine for a free chat. `varc.py brief` fills it in for you.
- `varc.py`: the helper script (Python 3.8+ and openpyxl).
- `templates/CAT_VARC_Mistake_Log_template.xlsx`: an empty mistake log with the Summary formulas.

## Free version: one-time setup

1. Install Python 3 from python.org. On Windows, tick "Add python.exe to PATH" in the installer.
2. Download the kit: [ZIP of this branch](https://github.com/Divyansh200Bansal/Repo-For-AI/archive/refs/heads/claude/awesome-mccarthy-csg9ps.zip). Unzip it and open a terminal in the `cat-varc-kit` folder. (If this branch is merged into `main`, use Code, then Download ZIP, on the repository page instead.)
3. Install the one library: `pip install openpyxl`. On Windows, use `py -m pip install openpyxl` if `pip` is not found.
4. Keep your history: open your mistake-log Sheet on claude.ai, download it as an .xlsx file, and save it in this folder as `CAT_VARC_Mistake_Log.xlsx`. To start fresh, skip this step; the script copies the empty template the first time.

## Free version: every morning

1. Run `python varc.py brief`. It shows today's question mix and writes `prompt_today.txt`.
2. Open `prompt_today.txt`, copy all of it, and paste it into a new chat. Turn on web search if you see the option.
3. Run `python varc.py timer` in a terminal beside the chat. Press Enter each time you send START, DONE and each answer. Short on messages? Send all seven answers in one message; the timer still times each question.
4. After Q7, paste the `TIMES:` line from the timer into the chat, then read the debrief.
5. Run `python varc.py log` and paste the LOG BLOCK from the end of the debrief. Rows go below the last filled row of the Log, Daily and Vocab tabs. Existing rows are never changed, and a backup goes to `records/backups/` first.
6. Any time: `python varc.py stats` prints the Summary numbers (accuracy by type, reasons for mistakes, last 7 days).

On Windows, type `py` instead of `python` if `python` is not found.

## Rules the script follows

- Question mix: 1 Main idea, 2 Inference, 1 Detail, 1 Tone, 1 Vocab in context. If a type has 3+ attempts and accuracy below 60%, one question of your strongest type (never Main idea) is swapped for an extra one of the weak type.
- Q7: the verbal ability type with the fewest rows in the Log. Ties go Para-jumble, then Para summary, then Odd sentence out, so the three types rotate.
- Labels are checked against the lists the Summary formulas count. A near miss such as "para jumble" is corrected; an unknown label stops the write and says which line to fix.

## Privacy

This repository is public. Your filled-in log and timing records hold your answers, so `.gitignore` keeps `CAT_VARC_Mistake_Log.xlsx`, `records/` and `prompt_today.txt` out of git. Don't force-add them.

## If something goes wrong

- "Could not save": close the Excel file and run `log` again. Nothing was written.
- "is not an allowed label": fix that word in the block, or save the block to a file and run `python varc.py log --from-file block.txt`.
- Excel shows old Summary numbers: press Ctrl+Alt+F9 to recalculate. `python varc.py stats` always counts the rows directly.
- The chat can't search the web: it will ask for five headlines. Paste them from any news app.
