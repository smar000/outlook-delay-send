# Outlook Delay Send Add-in

An Office.js add-in for New Outlook (Windows) that automatically applies a configurable server-side delivery delay to every outbound email. The message sits in Drafts during the hold window — editable and deletable — then Exchange Online delivers it automatically, even if Outlook is closed.

## How it works

When you click Send, the `OnMessageSend` event fires and the add-in schedules delivery for a future time using `item.delayDeliveryTime.setAsync()`. Exchange holds the message server-side until that time.

Settings and the Send Now bypass are accessible via a task pane, opened from the **More Apps** button in the compose or reading ribbon.

## Setup

### 1. Deploy to GitHub Pages

1. Fork or clone this repository and push to your own public GitHub repo named `outlook-delay-send`
2. Replace `smar000` in `manifest.xml` (two occurrences) with your GitHub username
3. Go to **Settings → Pages**, set source to **Deploy from branch**, branch `main`, folder `/`
4. Wait ~2 minutes; confirm `https://YOUR_USERNAME.github.io/outlook-delay-send/src/commands.html` loads in a browser

### 2. Sideload the add-in

**Option A — New Outlook UI:**
1. Click the **Apps** icon in the left sidebar
2. Click **Add apps → Upload a custom app**
3. Select `manifest.xml`

**Option B — Outlook on the Web (recommended if Option A isn't available):**

If you don't see an upload option in the New Outlook UI, use this direct link instead:

👉 **https://aka.ms/olksideload**

1. Go to that URL in your browser
2. Click **My add-ins** → **Add a custom add-in** → **Add from file**
3. Upload `manifest.xml`

The add-in will then appear in both Outlook on the Web and New Outlook on Windows.

## Usage

### Automatic delay (default behaviour)
Just compose and send as normal. Every email is automatically held for the configured delay period before delivery. During the hold window, the message appears in **Drafts** — you can open and delete it there to cancel the send.

### Opening the settings panel
The task pane is accessed via the **More Apps** button in the compose ribbon:

1. Open a new message
2. Click **More Apps** in the ribbon
3. Click **Delay Send** in the popup
4. The settings panel opens on the right side of the window

### Send Now (bypass the delay for one email)
1. Open the settings panel (see above)
2. Toggle **Send Next Email Immediately** on
3. Close the panel and click **Send** — the message sends without any delay

The bypass applies to that email only and resets automatically after sending.

### Changing the delay duration
1. Open the settings panel (see above)
2. Enter your preferred delay in seconds (e.g. 300 for 5 minutes)
3. Click **Save**

The new delay applies to all future sends. The setting is stored in your Microsoft 365 account and roams across devices.

## Requirements

- Microsoft 365 / Exchange Online account (not IMAP/POP3)
- New Outlook for Windows (Mailbox requirement set 1.13+, mid-2023 or later)

## Testing

1. Compose an email to yourself and click Send
2. Immediately check **Drafts** — the message should appear there
3. After the configured delay it moves to **Sent Items**
4. To cancel: open the message in Drafts and delete it during the hold window

## Adjusting the default delay in code

The fallback delay (used if no setting has been saved) can be changed in [src/commands.js](src/commands.js):

```javascript
const delaySeconds = settings.get('delaySeconds') || 120; // change 120 to your preferred default
```

Commit and push — GitHub Pages updates within ~1 minute. No manifest change needed.

## Limitations

| Concern | Detail |
|---|---|
| Exchange/M365 only | Won't apply to IMAP/POP3 accounts added to Outlook |
| Minimum Mailbox 1.13 | Requires a recent version of New Outlook (mid-2023+) |
| SoftBlock behaviour | If GitHub Pages is unreachable at send time, Outlook warns the user but allows them to override and send immediately |
| Ribbon button | For sideloaded add-ins in New Outlook, the task pane is accessed via More Apps rather than a dedicated ribbon button |
| No automatic Drafts UI | To cancel a queued message, go to Drafts manually and delete it |
