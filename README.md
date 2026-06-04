# Outlook Delay Send Add-in

An Office.js add-in for New Outlook (Windows) that automatically applies a **120-second server-side delivery delay** to every outbound email. The message sits in Drafts during the hold window — editable and deletable — then Exchange Online delivers it automatically.

## How it works

When you click Send, the `OnMessageSend` event fires, and the add-in calls `item.delayDeliveryTime.setAsync()` to schedule delivery 120 seconds in the future. Exchange holds the message server-side; your Outlook client doesn't need to stay open.

## Setup

### 1. Deploy to GitHub Pages

1. Create a public GitHub repository named `outlook-delay-send`
2. Push this repo to it
3. Go to **Settings → Pages**, set source to **Deploy from branch**, branch `main`, folder `/`
4. Wait ~2 minutes; confirm `https://smar000.github.io/outlook-delay-send/src/commands.html` loads

### 2. Update the manifest

Replace `smar000` in `manifest.xml` (two occurrences) with your GitHub username:

```xml
DefaultValue="https://smar000.github.io/outlook-delay-send/src/commands.html"
```

### 3. Sideload in New Outlook

**Option A — New Outlook UI:**
1. Click the **Apps** icon in the left sidebar
2. Click **Add apps → Upload a custom app**
3. Select `manifest.xml`

**Option B — Outlook on the Web:**
1. Go to https://outlook.office.com/mail/options/general/manage-add-ins
2. Click **+** → **Add from file** → upload `manifest.xml`

## Adjusting the delay

Edit `DELAY_SECONDS` in [src/commands.js](src/commands.js) and push — GitHub Pages updates within ~1 minute.

```javascript
const DELAY_SECONDS = 120; // change this value
```

## Requirements

- Microsoft 365 / Exchange Online account (not IMAP/POP3)
- New Outlook for Windows (Mailbox requirement set 1.13+, mid-2023 or later)

## Testing

1. Compose an email to yourself and click Send
2. Immediately check **Drafts** — the message should appear there
3. After 120 seconds it moves to **Sent Items**
4. To cancel: open the message in Drafts and delete it during the hold window
