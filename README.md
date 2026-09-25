# Recruiting Agent (networking)

A Google Apps Script add-on that lives inside your **Recruiting Tracker** Google Sheet and:

1. **Finds people** at your target companies (Penn alumni first) with Apollo.io, including their **LinkedIn profile, email, and a short "About Them" background** (current role, past companies).
2. **Drafts outreach** for each person with Claude: a cold **email**, a **LinkedIn connection note** (fits the 200/300-character limit), and a longer **LinkedIn message** for after they accept. Tone adjusts by industry (formal for consulting/finance, warmer for music/consumer).
3. **Puts emails in your Gmail Drafts** and waits for your approval. Nothing is sent without you.
4. **LinkedIn queue**: a sidebar with each approved person's profile link, their background, and copy buttons for the note and message.
5. **Tracks everything**: marks emails Sent/Replied from your inbox, drafts follow-ups after 7 days with no reply, logs every touch into your existing **Networking 2027** tab, and keeps a **Referrals** tab.

It skips anyone already in Networking 2027 / Networking 2026 / the queue, so you don't email the same person twice or waste Apollo credits.

## Why LinkedIn is "one click" and not fully automatic

LinkedIn doesn't offer an API for sending connection requests or messages, and tools that automate your account break LinkedIn's rules and regularly get accounts restricted. That's a bad risk during recruiting season. So the agent does everything up to the send: it finds the person, links the profile, writes the note, and copies it for you. You click **Connect → Add a note → paste → Send**, which takes about 10 seconds per person. Email is fully automated after you approve it.

Free LinkedIn accounts only get a few personalized connection notes per month. If you hit that limit, send the request without a note and use the "after they accept" message instead.

## Setup (about 10 minutes, one time)

### 1. Get API keys
- **Apollo.io**: sign up (free) → Settings → Integrations → API → create a key. If Apollo says your plan can't use an endpoint, the error message will show up in the sheet. Searching people is free; each person the agent adds costs **1 credit** (that's how it gets the email and full name).
- **Anthropic** (optional but recommended): [console.anthropic.com](https://console.anthropic.com) → API Keys. Add $5 of credit; each person costs roughly a few cents to draft. Without a key the agent uses a basic template.

### 2. Paste the code into your sheet
1. Open **Recruiting Tracker** → **Extensions → Apps Script**.
2. For each `.gs` file in [`apps-script/`](apps-script/) (`Config`, `Menu`, `Setup`, `Apollo`, `Drafts`, `Outreach`): click **+ → Script**, name it the same (without `.gs`), and paste the contents. Delete the default `Code.gs` if it's empty.
3. Click **+ → HTML**, name it `LinkedInQueue`, and paste [`apps-script/LinkedInQueue.html`](apps-script/LinkedInQueue.html).
4. Optional: Project Settings → check "Show appsscript.json", then paste [`apps-script/appsscript.json`](apps-script/appsscript.json) (sets the time zone to New York).
5. Save, then reload the sheet. A **Recruiting Agent** menu appears.

### 3. First run
1. **Recruiting Agent → Setup → Set up tabs.** This creates `Targets`, `Outreach Queue`, `Referrals`, `Agent Settings`. Your existing tabs are not changed.
2. Google will ask for permission (Sheets, Gmail, external requests). The warning says "unverified app" because it's your own script: click **Advanced → Go to project**.
3. **Setup → Set API keys.** Keys are stored privately in your Google account, not in the sheet.
4. Fill in **Agent Settings**, especially *School & year*, *About me*, and *Email signature*. The more specific *About me* is, the better the drafts.
5. Edit **Targets**: check **Active** for the companies you want, adjust titles, locations, and how many contacts per company. MBB are pre-filled.
6. Optional: **Setup → Turn on daily inbox check** (runs every morning to catch replies and draft follow-ups).

## Daily workflow

| Step | Menu item | What happens |
|---|---|---|
| 1 | **Find contacts** | New people appear in `Outreach Queue` with Status **New**, LinkedIn link, and About Them. Set **Penn Alum?** to Yes/No if you can tell (only "Yes" lets the draft mention Penn). |
| 2 | **Write drafts** | Email, LinkedIn note, and LinkedIn message filled in; email saved to Gmail Drafts; Status → **Draft Ready**. |
| — | *You review* | Edit anything. Set Status to **Approved** or **Skip**. Change **Channel** (Email / LinkedIn / Both) if you want. |
| 3 | **Send approved emails** | Sends the Gmail drafts for Approved rows (asks you to confirm first) and logs them. You can also just hit Send in Gmail; the inbox check will notice. |
| 4 | **Open LinkedIn queue** | Sidebar with each approved person: open profile, copy note, copy message, **I sent it**. |
| — | **Check inbox now** | Marks Sent/Replied, creates follow-up drafts in the same Gmail thread after 7 days. For LinkedIn replies, set Status to **Replied** yourself. |

Other menu items:
- **Rebuild Gmail draft for selected row(s)**: if you edited the email in the sheet, this pushes the edit to Gmail.
- **Add selected contact(s) to Referrals**: starts a referral row you can track (To Ask → Asked → Agreed → Submitted).

## Development

The code is plain Apps Script (V8). Tests run the real files in Node against fake Sheets/Gmail/Apollo/Claude:

```
node tests/run.js
```

## Next phases
- Application autofill (browser-based, with you approving each submit)
- Pulling coffee chats from Google Calendar into the log
