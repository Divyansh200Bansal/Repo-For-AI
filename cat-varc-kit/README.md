# CAT VARC daily kit

A 30-minute morning routine for CAT VARC: 5 finance stories, an original CAT-style passage with 7 questions, timing, grammar feedback, and a mistake log that picks the next day's questions. It runs three ways.

| Step | Automatic: `python daily.py` | Pro: Claude cloud session | Free: any chat, by hand |
| --- | --- | --- | --- |
| News, passage, questions | Free RSS feeds + the Claude API | Claude writes them into a doc | Claude writes them in the chat |
| Today's question mix | Read from your Excel log | Claude reads your Sheet's Summary tab | `python varc.py brief` |
| Timing | Built into the quiz | Claude runs `date +%s` per message | `python varc.py timer` |
| Debrief and logging | Claude API + your Excel file | Claude writes rows into your Sheet | `python varc.py log` |
| Cost | Pay per use (see below) | Pro plan | Free |

## What's in this folder

- `daily.py`: the whole routine in one command (needs an Anthropic API key).
- `prompts/pro_prompt.md`: the routine for a Claude Code cloud session. Add your Sheet link, then paste it.
- `prompts/free_prompt.md`: the routine for a free chat. `varc.py brief` fills it in for you.
- `varc.py`: the helper script (Python 3.8+ and openpyxl).
- `templates/CAT_VARC_Mistake_Log_template.xlsx`: an empty mistake log with the Summary formulas.

## Fully automatic: `daily.py`

One command does the whole morning. It fetches the last 24 hours of finance headlines from free RSS feeds (Economic Times, Mint, Business Standard, Google News, BBC, CNBC), and Claude writes the 5 stories, the passage and the 7 questions with your log's question mix. You then take the timed quiz in the same window. Claude writes the debrief, and the rows go into your Excel log.

**What it needs:** Python 3.10 or newer, `pip install anthropic openpyxl`, and an Anthropic API key. The API is pay-as-you-go and separate from a Claude Pro subscription: create a key in the Claude Console (platform.claude.com) and add a small prepaid credit.

**What a run costs.** These are rough figures; the script prints the real cost at the end of each run.

| Model | Option | About per day | About per month |
| --- | --- | --- | --- |
| Claude Opus 5.5 (default; the best questions) | none | $0.30-0.40 (Rs 30) | $10 (Rs 900) |
| Claude Sonnet 5.5 | `--model claude-sonnet-5-5` | $0.15-0.20 | $5 |
| Claude Haiku 5.5 (cheapest; weaker questions) | `--model claude-haiku-5-5` | under $0.01 (Rs 1) | $0.25 |

`--effort medium` makes the model think less and costs less again. If Opus or Sonnet declines a request, the API retries it on another model automatically.

**In Colab:** add the key once under Secrets (the key icon on the left) as `ANTHROPIC_API_KEY`, with notebook access on. Then run this cell each morning:

```python
from google.colab import drive, userdata
import os
drive.mount('/content/drive')
%cd /content/drive/MyDrive/cat-varc-kit
!pip install -q anthropic
os.environ["ANTHROPIC_API_KEY"] = userdata.get("ANTHROPIC_API_KEY")
!python daily.py
```

**On a laptop:** set the key once (`export ANTHROPIC_API_KEY=...` on Mac or Linux; `setx ANTHROPIC_API_KEY "..."` on Windows, then open a new terminal), then run `python daily.py`.

**During the run:** read the news, press Enter to start the passage and again when you finish reading, then answer each question: your own answer first, then the option, then your reason. If the run stops half-way, run it again and it continues where you stopped. `--new` makes a second set on the same day.

**Files:** each day gets `days/<date>/`, with `brief.md` (news, passage and questions, to read on your phone), `debrief.md`, and `pack.json`, which holds the answer key, so don't open it before you finish.

**Limits:** Claude sees only the headlines and their short summaries; it does not read the articles. So "what happened" is only as detailed as the headline. If a feed stops working, the run skips it; you can swap feeds by editing `FEEDS` at the top of `daily.py`.

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

## Run it in Google Colab

Don't paste `varc.py` into a cell. A cell is not a file, so Python has no `__file__` and the script can't find its folder. Keep the kit in Google Drive and run the file with `!python`.

One time: upload `cat-varc-kit.zip` to My Drive, then run this cell:

```python
from google.colab import drive
drive.mount('/content/drive')
%cd /content/drive/MyDrive
!unzip -o cat-varc-kit.zip
```

Save your log in `My Drive/cat-varc-kit/` as `CAT_VARC_Mistake_Log.xlsx`. Every time you open Colab, run this first, because Colab forgets everything when its runtime restarts:

```python
from google.colab import drive
drive.mount('/content/drive')
%cd /content/drive/MyDrive/cat-varc-kit
```

Then use the commands with `!` in front: `!python varc.py brief`, then `print(open('prompt_today.txt').read())` to copy the prompt; `!python -u varc.py timer`, pressing Enter in the box under the cell; and `!python varc.py stats`. Colab's input box takes one line at a time, so for `log`, save the block to a file first: start a cell with `%%writefile block.txt`, paste the block below that line, run it, then run `!python varc.py log --from-file block.txt`.

Keep the folder in Drive. Colab's own disk is wiped when the runtime resets, and your log would go with it.

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
