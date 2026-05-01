
---

## How it Works

```
OneNote (cloud) → Graph API → Python script → .md files → VS Code + Copilot
```

Run the script manually or on a schedule (e.g. every hour via Task Scheduler / cron). It only rewrites files that changed.

---

## Step 1: Register a Free Azure App (one-time, 5 min)

1. Go to **portal.azure.com** → **Azure Active Directory → App registrations → New registration**
2. Name it anything (e.g. `OneNoteSync`), choose **Personal accounts only**
3. Under **API permissions** → Add → Microsoft Graph → Delegated → add `Notes.Read`
4. Under **Authentication** → Add platform → Mobile/Desktop → check `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Copy your **Application (client) ID**

---

## Step 2: The Sync Script

Save this as `onenote_sync.py` in your VS Code workspace:

Here's the full sync setup:

---

## Setup (one-time)

**1. Install dependencies:**
```bash
pip install msal requests html2text
```

**2. Register your Azure App** (free, 5 min):
- Go to **portal.azure.com** → App registrations → New registration
- Name it anything, choose **Personal accounts only**
- API permissions → Microsoft Graph → Delegated → add `Notes.Read`
- Authentication → Mobile/Desktop → check the `login.microsoftonline.com` redirect URI
- Copy your **Application (client) ID**

**3. Paste your client ID** into the `CLIENT_ID` field at the top of the script, and set `OUTPUT_DIR` to your VS Code workspace folder.

---

## Running it

```bash
python onenote_sync.py
```

- **First run:** opens a browser for Microsoft login (one-time — token is cached)
- **Subsequent runs:** fully automatic, no login needed
- **Smart sync:** skips pages that haven't changed since last run

---

## Auto-sync (keep in sync continuously)

**Windows Task Scheduler:** run the script every hour
**Mac/Linux cron:**
```bash
# runs every hour
0 * * * * python /path/to/onenote_sync.py
```

---

## Output structure in VS Code

```
onenote-notes/
├── Work Notebook/
│   ├── Meetings/
│   │   ├── Sprint Planning.md
│   │   └── 1on1 Notes.md
│   └── Projects/
│       └── Architecture.md
└── Personal/
    └── Ideas/
        └── Side Project.md
```

Then in VS Code, use **`@workspace`** in Copilot Chat to query across all your notes instantly.
