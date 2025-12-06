# Google-Apps-Script
Repository holding scripts to automate email distribution and email acknowledgements via scripts.google.com

## Newsletter Automation

### Overview
This Google Apps Script automates the process of sending personalized newsletters via Gmail. It reads a draft email containing "Haptigation Newsletter" in the subject line and sends it to all contacts with a specific label, personalizing each email with the recipient's name.
Implementing it this way ensures data safety and allows for personalization in the newsletter.

---

### Prerequisites
**1. Enable Gmail API Access**

- Go to https://script.google.com
- Open your project (or create a new one)
- Click on the Services icon (⊕) in the left sidebar
- Find Gmail API and click Add

**2. Enable People API**

- In the same Google Apps Script project
- Click on the Services icon (⊕) in the left sidebar
- Search for People API
- Click Add to enable it
- This allows the script to access your Gmail contacts and labels

**3. Gmail Settings (if needed)**

- Check if in your gmail account setting IMAP is enabled.
- Ensure that "Allow less secure apps" is NOT required for Google Apps Script (it runs with your Google account permissions)
- Make sure your Gmail account has sufficient daily sending quota (typically 500 emails/day for regular Gmail, 2000/day for Google Workspace)

---

### Setup Instructions
**Step 1:** Access Gmail and Google Apps Script

Open [Gmail](https://mail.google.com) and [Google Apps Script](https://script.google.com/home) for your account.

**Step 2:** Create Newsletter Draft in Gmail

In Gmail, create a new email draft containing your complete newsletter content
IMPORTANT: The subject line MUST include "Haptigation Newsletter" somewhere (e.g., "Haptigation Newsletter - December 2025")

Use [NAME] as a placeholder where you want the recipient's name to appear.

If users [NAME] is "" or "unbekannt" [NAME] will be replaced with "LeserIn" instead of the contact name. 

**DO NOT SEND - Save as draft only**

**Step 3:** Configure Settings

Replace the CONFIG variables with your actual values:

```
javascript

const CONFIG = {
  contactLabel: "Newsletter Subscriber",  // The label of your newsletter contacts
  emailSubject: "{Enter your subject}",   // This replaces the draft's subject
  senderName: "{Your Name} von Company",  // Shows as "Max von Company via company@gmail.com"
  testMode: false  // Set to true for testing, false for live sending
};
```

**Step 4:** Save the Script
Click the disk icon (💾) or press Ctrl+S / Cmd+S to save your changes.

<img width="1227" height="134" alt="save icon" src="https://github.com/user-attachments/assets/d5c373b4-24ea-4bc5-aea0-cbb9c12048b4" />

**Step 5:** Run the Script

In the function dropdown (top center), select "sendNewsletter"
Click the Run button (▶️) to the left

<img width="1227" height="144" alt="run function" src="https://github.com/user-attachments/assets/d0c9576f-31a4-489e-b2e4-4e1a9b0df6c7" />

On first run, you'll need to authorize the script:

Click "Review permissions"
Select your Google account
Click "Advanced" → "Go to [Project Name] (unsafe)"
Click "Allow"

**Step 6:** Reset Safety Settings
After successful sending:

Change testMode back to true
Change contactLabel back to "safetyFirst"
This prevents accidental sends

**Step 7:** Save Again
Click the disk icon (💾) or press Ctrl+S / Cmd+S to save your changes.

<img width="1227" height="134" alt="save icon" src="https://github.com/user-attachments/assets/d5c373b4-24ea-4bc5-aea0-cbb9c12048b4" />

**Step 8:** Cleanup

Verify the newsletter was sent successfully (check your Sent folder)
Delete the draft or modify its subject to NOT include "Haptigation Newsletter"
You can restore the subject line when preparing the next newsletter edition

---

### Features

#### Personalization

Use [NAME], [Name], or [name] in your draft
Recipients with stored names: "Hallo [NAME]" → "Hallo Max"
Recipients marked as "unknown" or empty: "Hallo [NAME]" → "Hallo LeserIn"

#### Test Mode
When testMode: true, the script only sends to your own email address for testing.

#### Rate Limiting
The script automatically pauses every 50 emails for one cycle (default 1min) to avoid Gmail's rate limits.

#### Helpful Functions
**Check Your Contacts**
Before sending, verify your contacts:
In the function dropdown, select "showContacts" and click Run

**List All Contact Groups**
To see all available contact labels:
In the function dropdown, select "listAllContactGroups" and click Run

---

#### Troubleshooting
**"No Newsletter Draft Found"**

Ensure your draft's subject contains "Haptigation Newsletter"
Check the execution log for available draft subjects

**"No Contacts Found"**

Verify the contact label name matches exactly
Ensure contacts have the specified label in Gmail Contacts
Check that People API is enabled

**"Permission Denied"**

Re-run authorization process
Ensure Gmail API and People API are both enabled in Services

**Daily Sending Limits**

Regular Gmail: 500 emails/day
Google Workspace: 2,000 emails/day
If you exceed limits, wait 24 hours

---

### Security & Data Protection
**⚠️ CRITICAL - READ CAREFULLY:**

You are working with sensitive personal data
Contact information MUST NOT be shared outside your organization
Never export or share contact lists
Only authorized personnel should access this script
Always keep testMode: true when not actively sending
Use contactLabel: "safetyFirst" as default to prevent accidental sends

### Contact Management
DO NOT manually edit contacts - designated personnel will maintain contact data integrity. 

If you do please consider all consequences of your action (e.g. false names, accidentally unsubscribing newsletters, etc.)

### Support
If you encounter issues:

Check the Execution log (View → Logs or Ctrl+Enter)
Verify all configuration settings
Ensure APIs are enabled
Test with testMode: true first

---

Last Updated: December 2025
Version: 1.0

## Email Auto-Responder

### Overview
This Google Apps Script automatically responds to incoming unread emails in your Gmail inbox. It monitors for new messages and sends a pre-written response from a draft email, ensuring each sender receives only one automated reply.

---

### Prerequisites
**1. Enable Gmail API Access**

Go to https://script.google.com
Open your project (or create a new one)
Click on the Services icon (⊕) in the left sidebar
Find Gmail API and click Add

**2. Gmail Settings**

Ensure your Gmail account has sufficient daily sending quota (typically 500 emails/day for regular Gmail, 2000/day for Google Workspace)
The script runs with your Google account permissions (no additional settings needed)

---

### Setup Instructions
**Step 1:** Create Response Draft in Gmail

In Gmail, compose a new email draft with your auto-response message
**CRITICAL: Set the subject line to exactly:**

```E-Mail Response Automation Draft - DO NOT DELETE```

Write your response message (supports HTML formatting and inline images)

**DO NOT SEND - Save as draft only**

**Never delete this draft - the script requires it to function**

**Step 2:** Open Google Apps Script

Go to https://script.google.com

**Step 3:** Configure Settings (Optional)
You can customize the CONFIG settings at the top of the script:

```
javascript
const CONFIG = {
  DRAFT_SUBJECT: 'E-Mail Response Automation Draft - DO NOT DELETE',  // Draft subject (must match exactly)
  REPLY_SUBJECT: 'Re: ',                                              // Reply subject prefix
  MAX_EMAILS_PER_RUN: 50,                                             // Max emails processed per check
  CHECK_INTERVAL_MINUTES: 1,                                          // How often to check (in minutes)
  LABEL_NAME: 'AutoResponded'                                         // Label for processed emails
};
```

**Settings Explained:**

DRAFT_SUBJECT: The exact subject of your response draft (don't change unless you change the draft)
REPLY_SUBJECT: Prefix for replies (default "Re: " creates "Re: Original Subject")
MAX_EMAILS_PER_RUN: Limits emails processed per check (prevents quota issues)
CHECK_INTERVAL_MINUTES: Frequency of checks (1 = every minute, 5 = every 5 minutes, etc.)
LABEL_NAME: Gmail label applied to processed emails (prevents duplicate responses)

**Step 4:** Save the Script
Click the disk icon (💾) or press Ctrl+S / Cmd+S to save your changes.

<img width="1302" height="127" alt="save icon" src="https://github.com/user-attachments/assets/9f713499-f400-4042-8c1b-537d62180816" />

**Step 5:** Set Up Automatic Trigger

In the function dropdown (top center), select "setupTrigger"
Click the Run button (▶️)

<img width="1223" height="141" alt="run function" src="https://github.com/user-attachments/assets/1fa2cedf-4218-4e49-bcb0-c0d7946974e2" />

On first run, authorize the script:

Click "Review permissions"
Select your Google account
Click "Advanced" → "Go to [Project Name] (unsafe)"
Click "Allow"

Check the execution log - you should see: "Trigger erfolgreich eingerichtet"

**Step 6:** Verify Setup
The auto-responder is now active! It will:

Check for unread emails every X minutes (as configured)
Send your draft response to each new sender
Apply the "AutoResponded" label to prevent duplicates
Continue running automatically

---

### Workflow

Every X minutes (default: 1 minute), the script checks your inbox
Searches for unread emails without the "AutoResponded" label
For each new email:

Reads the sender's address
Sends your draft message as a reply
Applies the "AutoResponded" label
Logs the action

Prevents duplicates: Once labeled, the script won't respond again to that sender

---

### Smart Features

- Inline Images: Supports embedded images in your draft - Only supports pictures if added via URL (e.g. [imgur](https://imgur.com) pictures)
- HTML Formatting: Preserves rich text formatting from your draft
- Thread-Safe: Works with email threads/conversations
- Duplicate Prevention: Uses Gmail labels to track responded emails
- Rate Limiting: Processes maximum 50 emails per run by default

---

### Managing the Auto-Responder
**Check if Trigger is Active**

In Google Apps Script, click the clock icon (⏰) in the left sidebar
You should see "autoResponseToEmails" listed with its schedule

**Temporarily Disable**

Select "removeTrigger" from the function dropdown
Click Run (▶️)
Verify in execution log: "Trigger entfernt"

**Re-Enable**

Select "setupTrigger" from the function dropdown
Click Run (▶️)

**Update Response Message**

Edit your draft in Gmail (the one with "E-Mail Response Automation Draft - DO NOT DELETE" subject)
Save changes - the script will automatically use the updated version
No need to restart the trigger

**Testing**
1. Before setting up the trigger, test manually:
2. Send yourself a test email (from another account)
3. In Google Apps Script, select "autoResponseToEmails"
4. Click Run (▶️)
5. Check execution log for "Auto-Response gesendet an: [your-email]"
6. Verify you received the response
7. Check Processed Emails
- In Gmail: Search for label:AutoResponded to see all emails that received auto-responses

**Configuration Options**
- Change Check Frequency
Modify CHECK_INTERVAL_MINUTES:

1 = Every minute (most responsive, uses more quota)
5 = Every 5 minutes (recommended for moderate volume)
15 = Every 15 minutes (for low volume)
60 = Every hour (minimal quota usage)

After changing, run setupTrigger again to apply.

- Customize Reply Subject

If: 'Re: ' → "Re: Original Subject" (default, wont work if it doesn't start with 'Re: ')
If: 'Thank you for your email!' → Uses this exact subject (change content as you want)

- Adjust Max Emails Per Run
If you receive high email volume:

Increase MAX_EMAILS_PER_RUN to 100 or 200
Be mindful of Gmail's daily sending limits

---

### Troubleshooting
**"No Draft Found" Error**

Verify your draft's subject is exactly: E-Mail Response Automation Draft - DO NOT DELETE
Check for extra spaces or typos
Ensure the draft wasn't accidentally deleted

**No Responses Being Sent**

Check trigger is active (clock icon in script editor)
Verify you have unread emails without "AutoResponded" label
Check execution log for errors (View → Executions)
Test manually with autoResponseToEmails function

**Duplicate Responses**

Shouldn't happen due to label system
If it does, check that LABEL_NAME hasn't changed
Verify the label is being applied (check in Gmail)

**Daily Sending Limit Reached**

Regular Gmail: 500 emails/day
Google Workspace: 2,000 emails/day
Reduce CHECK_INTERVAL_MINUTES or MAX_EMAILS_PER_RUN
Consider upgrading to Google Workspace

**Images Not Appearing**

Ensure images are embedded in the draft using URL (not attached as files)
Gmail sometimes blocks inline images - recipients may need to "Display images"

---

### Advanced Features
Mark Emails as Read (Optional)
Uncomment this line in the script:
```
javascript
// latestMessage.markRead();
```
Remove the // to enable auto-marking as read.

# Now you are all set. Good Luck!
