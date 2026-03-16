# delaware Timesheet Automator

Reads your Outlook calendar, uses AI to map meetings to project codes, and generates a draft timesheet for `time.delaware.pro`.

## What it does

1. Reads your Outlook desktop calendar for the selected day(s)
2. Uses Claude AI to map each meeting to the right project code
3. Opens a review page in your browser with the draft
4. You review, edit if needed, and copy entries to paste into time.delaware.pro

**Time saved: ~15 minutes → ~3 minutes per day**

## Requirements

- Windows laptop
- Outlook desktop app (must be open when running)
- Python 3.x ([download](https://www.python.org/downloads/)) — tick "Add Python to PATH" during install
- An Anthropic API key ([get one here](https://console.anthropic.com)) — costs ~€0.01 per run

## Setup (one time, ~10 minutes)

### 1. Install Python
Download from https://www.python.org/downloads/ and install.  
**Important:** tick "Add Python to PATH" on the first screen.

### 2. Get an Anthropic API key
1. Go to https://console.anthropic.com
2. Sign in or create a free account
3. Go to API Keys → Create Key
4. Copy the key (starts with `sk-ant-`)

### 3. Configure
1. Copy `config.template.json` → rename to `config.json`
2. Open `config.json` and fill in:
   - `anthropic_key`: your API key from step 2
   - `user_id`: your delaware username (e.g. `werbrouckv`)

### 4. Set up your projects
1. Copy `projects.template.xlsx` → rename to `projects.xlsx`
2. Open in Excel and replace the example rows with your real project codes
3. See the Instructions sheet inside the file for how to find your project codes

## Daily use

1. Make sure Outlook is open
2. Double-click `run.bat`
3. Paste your bearer token when asked (see below)
4. Select which day(s) to process
5. Review the draft in your browser
6. Click "Copy all entries" → paste into time.delaware.pro

### Getting your bearer token
Open `time.delaware.pro` in Edge → F12 → Network → Fetch/XHR → click any `timeentry` request → Headers → copy the value after `Bearer `

This takes about 20 seconds once you know where to look. The token is valid for your whole browser session.

## Catch-up mode

Forgot to fill in timesheets for a few days? When you run `run.bat`, it shows you the last 7 working days and lets you select multiple days at once:

```
Which days to process?
[1] Friday, 13 March 2026  <- yesterday
[2] Thursday, 12 March 2026
[3] Wednesday, 11 March 2026
...

Enter number(s), e.g. 1 or 1,2,3:
```

## Project code mapping

The AI maps your meetings to project codes using the tags in `projects.xlsx`. The more keywords you add that match your actual meeting titles, the better the mapping.

Example: if your meeting is called "DATS biweekly sprintstatus", add `dats biweekly, sprint, sprintstatus` as tags for that project.

When you correct a mapping in the browser, the script saves it to the "Corrections Log" sheet in `projects.xlsx` and applies it automatically next time.

## Files

| File | Purpose |
|------|---------|
| `run.bat` | Double-click to run |
| `timesheet.py` | Main script |
| `projects.xlsx` | Your project codes and tags (personal, not in git) |
| `config.json` | Your API key and user ID (personal, not in git) |
| `projects.template.xlsx` | Template to copy for new users |
| `config.template.json` | Template to copy for new users |
| `test_suite.py` | Automated tests |

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "python is not recognized" | Reinstall Python, tick "Add Python to PATH" |
| "Could not connect to Outlook" | Make sure Outlook desktop app is open |
| No events found | Check Outlook is synced for that date |
| AI maps wrong project | Add better keywords to `projects.xlsx`, or edit in browser |
| Wrong project codes in dropdown | Check your bearer token is fresh (re-copy from F12) |

## Privacy

- Calendar events are read locally from your Outlook app
- Only meeting titles and durations are sent to the Anthropic API (no email content, no attendees)
- Your bearer token and API key are never saved to disk beyond `config.json`
- `config.json` and `projects.xlsx` are excluded from git via `.gitignore`

## Contributing

Found a bug or want to improve it? Open an issue or pull request. When adding features, run `python test_suite.py` first to make sure nothing breaks.
