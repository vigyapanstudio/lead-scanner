# Eventra Lead Scanner – Version 2 setup guide

This guide is for Vigyapan Studio, not for the client. Everything below is done once.

---

## 1. Upload the new files to GitHub

Open your `lead-scanner` repository on github.com. Click **Add file → Upload files**, drag in **all** the files from this folder, then click **Commit changes**. Upload the new versions even where a file name already exists.

These files are new in Version 2. Make sure they are uploaded:
`cards.py`, `cleanup.py`, `mini_report.py`, `storage.py`, `packages.txt`, plus the updated `requirements.txt`.

**The colour-theme file** (`.streamlit/config.toml`). A Mac hides folders whose names start with a dot, so create this one directly on GitHub:

1. Click **Add file → Create new file**.
2. In the name box, type `.streamlit/config.toml`. Typing the `/` creates the folder automatically.
3. Paste this in, then click **Commit changes**:

```
[theme]
base = "dark"
primaryColor = "#FFB72B"
backgroundColor = "#06201F"
secondaryBackgroundColor = "#0E3634"
textColor = "#F3F7F2"

[client]
toolbarMode = "minimal"

[browser]
gatherUsageStats = false

[server]
maxUploadSize = 50
```

Streamlit rebuilds the website by itself. The first rebuild takes 3–5 minutes, because it installs the card-reading software.

---

## 2. Put Eventra's contact details on the mini-reports

Each company's free mini-report ends with a "Want to talk?" box. On GitHub, open `brand.py`, click the ✏️ pencil icon and fill in:

```
CLIENT_PHONE = "+91 98XXX XXXXX"
CLIENT_EMAIL = "hello@eventravlogs.com"
CLIENT_WEBSITE = "www.eventravlogs.com"
CLIENT_INSTAGRAM = "@eventravlogs"
```

Any line left as `""` simply doesn't show. Then click **Commit changes**.

---

## 3. Optional add-ons (Streamlit → your app → ⋮ → Settings → Secrets)

Everything works without these. Each one switches on an extra feature. Put the plain lines (`NAME = "value"`) **at the top** of the Secrets box.

| What it unlocks | Line to add | Cost |
|---|---|---|
| Password on the website | `APP_PASSWORD = "your-password"` | Free |
| **Google Maps check** (rating, reviews, "not on Maps") | `GOOGLE_MAPS_KEY = "..."` | Google gives free monthly credit, then charges per lookup. Check current pricing on Google Cloud. |
| **Full Instagram check** (followers, last post, engagement) | `APIFY_TOKEN = "..."` | Paid per profile on apify.com |
| **Sharper visiting-card reading** (reads messy or angled cards much better) | `ANTHROPIC_API_KEY = "..."` and `CLAUDE_MODEL = "..."` | Paid per card on console.anthropic.com, roughly a fraction of a rupee per card. Use a current model name from Anthropic's models page. |
| **Permanent lead tracker** (Google Sheet) | See section 4 | Free |

### Getting a Google Maps key
1. Go to **console.cloud.google.com** and create a project (for example "Eventra Scanner").
2. Open **APIs & Services → Library**, search **"Places API (New)"** and click **Enable**.
3. Open **APIs & Services → Credentials → Create credentials → API key** and copy the key.
4. Google will ask you to add a billing card. The free monthly credit usually covers normal use.

---

## 4. Make the lead tracker permanent (Google Sheet)

Without this, the lead tracker is saved on the website but **wiped when the website restarts** (for example after an update, or after a few days without use). The client can download a backup from the tracker tab, but connecting a Google Sheet is better.

1. **Create the sheet.** In Google Drive, make a new empty Google Sheet, for example "Eventra Lead Tracker". Copy its ID from the address bar: it's the long code between `/d/` and `/edit`.
2. **Create a "robot" Google account (a service account).**
   In console.cloud.google.com, using the same project as above:
   - Open **APIs & Services → Library**. Enable **Google Sheets API** and **Google Drive API**.
   - Open **IAM & Admin → Service Accounts → Create service account**. Give it any name, then click **Done**.
   - Click the new service account, then **Keys → Add key → Create new key → JSON**. A file downloads.
3. **Share the sheet with the robot.** Open the downloaded JSON file in Notepad or TextEdit and copy the `client_email` (it ends in `iam.gserviceaccount.com`). In your Google Sheet, click **Share**, paste that email and give it **Editor** access.
4. **Add it to Streamlit Secrets** in this exact layout. The plain lines go first, and the `[gcp_service_account]` section goes last:

```
APP_PASSWORD = "your-password"
GSHEET_ID = "paste-the-sheet-id-here"

[gcp_service_account]
type = "service_account"
project_id = "copy from the JSON file"
private_key_id = "copy from the JSON file"
private_key = "-----BEGIN PRIVATE KEY-----\n...copy exactly from the JSON file...\n-----END PRIVATE KEY-----\n"
client_email = "copy from the JSON file"
client_id = "copy from the JSON file"
token_uri = "https://oauth2.googleapis.com/token"
```

Click **Save**. The yellow "Temporary storage" warning in the Lead tracker tab disappears once the sheet is connected. Every scan is then saved to the Google Sheet, so the client can also open their leads directly in Google Sheets.

---

## 5. What's new for the client (Version 2)

1. **WhatsApp in one tap.** Every lead has a green WhatsApp button that opens a personalised message (English or Hinglish), mentioning the event and what was found.
2. **Free mini-reports.** A branded 1-page "Digital Health Check" PDF for each company, downloadable for one company or as a ZIP for HOT + WARM leads.
3. **List clean-up.** Duplicates are removed, phone numbers are tidied, email typos are fixed (`gamil.com` becomes `gmail.com`), and wrong-looking numbers or emails are flagged.
4. **Missing details filled in.** If the list has no phone or email, the scanner picks them up from the company's website.
5. **Coverage / podcast prospects.** Companies already strong online are tagged as good targets for paid event coverage.
6. **Hinglish pitch lines** next to the English ones.
7. **Visiting-card scanner.** Upload card photos or use the camera, check the table, then send the contacts to the scanner.
8. **Lead tracker.** Status (New → Contacted → Interested → Won), follow-up dates and notes, kept across every event.
9. **Google Maps check** (with key): listed or not, rating, number of reviews.
10. **Full Instagram check** (with Apify token).
11. **Facebook, LinkedIn and YouTube** presence shown for every company.
12. **Events dashboard.** Events compared side by side: contacts, HOT leads, contacted, interested, won.
