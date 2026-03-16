# delaware Timesheet Automator — Setup Guide
## One-time setup (~10 minutes)

---

## Step 1 — Install Python

1. Go to https://www.python.org/downloads/
2. Click the big yellow "Download Python 3.x.x" button
3. Run the installer
4. **Important**: tick **"Add Python to PATH"** before clicking Install
5. Verify: open Start menu → type `cmd` → run `python --version`
   You should see something like `Python 3.12.0`

---

## Step 2 — Get your Anthropic API key

1. Go to https://console.anthropic.com
2. Sign in or create a free account
3. Go to **API Keys** in the left menu → click **Create Key**
4. Give it a name like "timesheet-automator" and copy the key (starts with `sk-ant-`)
5. Open `config.json` and paste it as the value for `"anthropic_key"`

```json
"anthropic_key": "sk-ant-api03-xxxxxxxxxxxxx"
```

---

## Step 3 — Fill in your user ID

Open `config.json` and replace the placeholder with your delaware email:

```json
"user_id": "firstname.lastname@delawareconsulting.com"
```

---

## Step 4 — First run

1. Make sure the **Outlook desktop app is open** on your laptop
2. Double-click **`run.bat`**

The script will install a small required package (pywin32) automatically on first run,
then ask you to restart it once. After that it runs normally every time.

---

## Step 5 — Submitting entries (each time you use it)

The review page needs a bearer token to submit. Here is how to get it in under 30 seconds:

1. Open **time.delaware.pro** in Edge (keep it open)
2. Press **F12** → click the **Network** tab → click **Fetch/XHR**
3. Click any request named `timeentry?date=...` in the list
4. Click the **Headers** tab on the right panel
5. Scroll to **Authorization** — you see `Bearer eyJ...`
6. Copy everything after the word `Bearer ` (the long string)
7. Paste it into the token field at the top of the review page

You only need to do this once per browser session. If you keep Edge open all day,
the token stays valid and you can reuse it.

---

## Daily use (after setup)

1. Open Outlook (if not already open)
2. Double-click `run.bat`
3. Browser opens with your AI-generated draft
4. Paste bearer token once (30 sec)
5. Review and edit entries if needed
6. Click **Submit to timesheet**

Total time: 2-3 minutes instead of 15-20.

---

## Command options

| Command | What it does |
|---------|-------------|
| Double-click `run.bat` | Process yesterday (default, skips weekends) |
| `python timesheet.py --today` | Process today so far |

---

## Updating project codes

When you start a new project, open `config.json` and add it to `"project_codes"`:

```json
"NEWCODE001": {
  "taskCode": "NEWCODE001.1.1",
  "label": "Client name — project description",
  "keywords": ["words that appear in your meeting titles for this project"]
}
```

The more keywords match your actual calendar event titles, the better the AI mapping.

---

## Upgrading to full auto-login later (Option A)

When IT approves your Azure app, add two lines to `config.json`:

```json
"ms_client_id": "your-azure-app-client-id",
"ms_tenant":    "common"
```

Then ask Claude to swap in the Option A script. No other changes needed.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "python is not recognized" | Reinstall Python, tick "Add Python to PATH" |
| "Could not connect to Outlook" | Make sure Outlook desktop app is open |
| "HTTP 401" when submitting | Token expired — re-copy from network tab |
| Wrong project codes | Edit in browser before submitting, or add keywords to config.json |
| No events found | Check Outlook is open and synced for that date |

---

## Privacy

- Calendar is read locally from your Outlook app — no data leaves your laptop for this step
- Claude AI receives only event titles and durations (no email bodies, no attendees)
- Bearer token is used in memory only and never saved to disk
- Keep `config.json` private — it contains your Anthropic API key
