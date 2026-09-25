# Recruiting Agent (networking)

A Google Apps Script add-on that lives inside your **Recruiting Tracker** Google Sheet. It looks at the **jobs you applied to**, finds people at those companies who **have things in common with you**, and hands you a ready-to-send message for each one: **open the link, copy, paste, go.**

1. **Reads your applications** (`2027 Recruiting Tracker` tab). Every company you applied to becomes a target, unless all your applications there are Rejected. It works out job titles to search for from the roles you applied to (e.g. "Firmwide Strategy Analyst" → Strategy Analyst, Analyst, Associate).
2. **Finds similar people** with Apollo.io. It runs one search per thing you share with people (Penn, Miami Beach, Consult for America, WUME, Pan Asian Dance Troupe; editable in Agent Settings) and ranks people by how many they match. Each person gets:
   - their **LinkedIn profile link**
   - **Why Them**: what you share and which of your applications they're closest to
   - **About Them**: current role and past companies
3. **Writes the messages** with Claude: a **LinkedIn connection note** (fits the 200/300-character limit), a longer **message for after they accept**, and optionally an **email** (saved to Gmail Drafts). The note leads with what you share and mentions the role you applied for. Tone matches the industry.
4. **LinkedIn queue**: a sidebar opens with every person: **Open profile → Copy note → paste → I sent it.** No separate approval step.
5. **Tracks everything**: logs each LinkedIn/email touch into your **Networking 2027** tab, marks email replies, drafts follow-ups after 7 days, and keeps a **Referrals** tab.

Your networking tabs are **not** used to find people. They're only checked so the agent never suggests someone you've already reached out to (and doesn't spend Apollo credits on them).

## Why LinkedIn is "one click" and not fully automatic

LinkedIn doesn't offer an API for sending connection requests or messages, and tools that automate your account break LinkedIn's rules and regularly get accounts restricted. That's a bad risk during recruiting season. So the agent does everything up to the send: it finds the person, links the profile, writes the note, and copies it for you. You click **Connect → Add a note → paste → Send**, which takes about 10 seconds per person. Email is fully automated after you approve it.

Free LinkedIn accounts only get a few personalized connection notes per month. If you hit that limit, send the request without a note and use the "after they accept" message instead.

## Setup (about 10 minutes, one time)

### 1. Get API keys
- **Apollo.io**: sign up (free) → Settings → Integrations → API → create a key. If Apollo says your plan can't use an endpoint, the error message will show up in the sheet. Searching people is free; each person the agent adds costs **1 credit** (that's how it gets the full name, LinkedIn and email). Looking up each new company once may also use a credit.
- **Anthropic** (optional but recommended): [console.anthropic.com](https://console.anthropic.com) → API Keys. Add $5 of credit; each person costs roughly a few cents to draft. Without a key the agent uses a basic template.

### 2. Paste the code into your sheet
1. Open **Recruiting Tracker** → **Extensions → Apps Script**.
2. For each `.gs` file in [`apps-script/`](apps-script/) (`Config`, `Menu`, `Setup`, `Sources`, `Apollo`, `Drafts`, `Outreach`): click **+ → Script**, name it the same (without `.gs`), and paste the contents. Delete the default `Code.gs` if it's empty.
3. Click **+ → HTML**, name it `LinkedInQueue`, and paste [`apps-script/LinkedInQueue.html`](apps-script/LinkedInQueue.html).
4. Optional: Project Settings → check "Show appsscript.json", then paste [`apps-script/appsscript.json`](apps-script/appsscript.json) (sets the time zone to New York).
5. Save, then reload the sheet. A **Recruiting Agent** menu appears.

### 3. First run
1. **Recruiting Agent → Setup → Set up tabs.** This creates `Targets`, `Outreach Queue`, `Referrals`, `Agent Settings`. Your existing tabs are not changed. (If you set up an earlier version, delete the old `Targets`, `Outreach Queue` and `Agent Settings` tabs first so they get the new columns.)
2. Google will ask for permission (Sheets, Gmail, external requests). The warning says "unverified app" because it's your own script: click **Advanced → Go to project**.
3. **Setup → Set API keys.** Keys are stored privately in your Google account, not in the sheet.
4. Check **Agent Settings**. It's pre-filled from your resume (Penn '28 PPE, Miami Beach, Pivot Tokyo, Consult for America, WUME, Pan Asian Dance Troupe). Edit **Similarity keywords** to change who ranks highest, and **Also draft emails** = No if you only want LinkedIn.
5. You don't need to fill in **Targets**. Step 1 fills it from your applications. You can uncheck **Active** on a company to skip it, tweak titles/locations, or add a company by hand.
6. Optional: **Setup → Turn on daily inbox check** (runs every morning to catch email replies and draft follow-ups).

## Daily workflow

| Step | Menu item | What happens |
|---|---|---|
| 1 | **Find people at companies I applied to** | Adds any new companies from your applications to Targets, then adds the best-matching people to `Outreach Queue` (link, Why Them, About Them). |
| 2 | **Write messages** | Writes the LinkedIn note, the after-accept message, and the email (if any), then **opens the LinkedIn queue**. |
| 4 | **LinkedIn queue** | For each person: **Open profile → Connect → Add a note → paste → I sent it.** Check the "shared" item on their profile first, since it comes from a keyword match. When they accept, copy the longer message into a DM. |
| 3 | **Send approved emails** | Optional. Emails sit in Gmail Drafts: send from Gmail, or set Status to **Approved** and use this. |
| — | **Check inbox now** | Marks email Sent/Replied, drafts follow-ups after 7 days. For LinkedIn replies, set Status to **Replied** yourself; it updates Networking 2027. |

Every time you apply somewhere new, add it to your applications tab as usual. The next step 1 picks it up.

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
