# Contact-Extractor
Contact Extractor – Professional Contact Discovery for Microsoft 365 and Outlook Archives


Contact Extractor – Detailed Features List
A powerful desktop tool to extract personal and professional contacts from Microsoft 365 (Online) and local Outlook data sources.
Core Functionality

Extracts email addresses, names, phone numbers, and company names from multiple sources
Automatically deduplicates contacts by email address
Smart name parsing (splits full names into First/Last Name)
Signature detection in email bodies to extract structured contact info
Exports results to Excel (.xlsx), CSV, and HTML

Online Mode (Microsoft 365 / Exchange Online via Microsoft Graph API)
Scans your live Microsoft 365 account (requires Azure AD app registration)








SourceData ExtractedNotesEmailsSender, To/CC/BCC recipients, body text, signatures, attachmentsDate-range filteredEmail AttachmentsExcel files fully parsed row-by-row, other docs scanned by filename/textSupports .xlsxCalendar EventsAll attendees (organizer + invited users)Date-range filteredContacts FolderFull contacts with name, emails, phone numbers, companyAll available emailsMicrosoft TeamsUsers from joined teams, channels, recent chat/channel messagesLimited to recent 50 msgs/channel
Additional Online Features:

Custom date/time range selection (with quick buttons: 7/30/90/365 days)
Real-time progress bar and debug log
Secure OAuth2 authentication (interactive browser + device code flow fallback)
Token caching (token_cache.bin) for seamless re-login
Supports multi-account and admin consent scenarios

Local Mode (Offline / Forensic Style)
Scans local data without internet or login







SourceSupported Formats / RequirementsPST & OST FilesFull recursive scanning (requires pypff / libpff)Local DocumentsScans ~/Documents, ~/Desktop, ~/Downloads for .txt, .csv, .xlsx, .docx, .htmlRunning Outlook Profile (Windows only)Scans all folders via COM (requires pywin32)Email Attachments in PST/OSTFilename-based extraction (content parsing planned)
Local Options:

Toggle: Scan attachments
Toggle: Scan subfolders (in PSTs and local folders)

Intelligent Contact Enrichment

Extracts phone numbers using multiple regex patterns (US/international)
Detects company names from signatures and text
Finds names near email addresses using patterns like:
John Doe <john@example.com>
john@example.com - Jane Smith

Falls back to parsing name from email (e.g., john.doe@company.com → John Doe)

User Interface (Modern Tkinter GUI)

Clean, scrollable, responsive layout
Mode switcher: Online vs Local
Live debug console (timestamps + detailed logs)
Progress bar + status updates
One-click export buttons
Configuration help & debug info dialogs

Security & Privacy

No data sent to third parties
All processing done locally
Tokens stored encrypted in MSAL cache
Config stored in plain config.ini (Client/Tenant ID only

Supported Platforms

Windows (full feature set: PST + Outlook COM)
macOS / Linux (PST scanning + local files only)


How to Set Up for Online Scanning (Microsoft 365)
Follow these steps exactly to enable Online Mode:
Step 1: Register an App in Azure AD

Go to: https://portal.azure.com
Navigate to Entra Admin → App registrations → New registration
Name: Contact Extractor (or anything)
Supported account types: Accounts in any organizational directory (Any Azure AD directory - Multitenant)
Redirect URI:
Platform: Mobile and desktop applications
URI: http://localhost

Click Register

Step 2: Add API Permissions (Delegated)
Go to API permissions → Add a permission → Microsoft Graph → Delegated permissions
Add these permissions:

Mail.Read
Contacts.Read
Calendars.Read
User.Read
Team.ReadBasic.All
Chat.Read (optional but recommended)

Click Grant admin consent for [Your Organization] (requires admin)
Step 3: Enable Public Client Flow
Go to Authentication tab:

Under Advanced settings → Allow public client flows → Set to Yes
Ensure http://localhost is listed under Redirect URIs
Default client type → Treat application as a public client → Yes

Step 4: Get Your IDs
From the app Overview page, copy:

Application (client) ID
Directory (tenant) ID

Step 5: Configure the App

Run the Contact Extractor
Go to Configuration → Azure AD Credentials
Paste:
Application (Client) ID
Directory (Tenant) ID

Click Save Configuration

Step 6: Authenticate

Click Authenticate
A browser window will open → log in with your Microsoft 365 account
If blocked, it falls back to device code flow (copy code from popup)

After login → you’ll see “Authenticated”

You’re ready to scan!

Recommended Permissions Summary (Minimal Set)




PermissionTypeRequired
Mail.Read,Delegated,Emails + attachments
Contacts.Read,Delegated,Contacts folder
Calendars.Read,Delegated,Calendar events
User.Read,Delegated,Basic profile
Team.ReadBasic.All,Delegated,Teams & channel members
Chat.Read,Delegated,Teams chats

Tips for Best Results

Use Last 365 days for maximum coverage
Run on a weekend if mailbox is large
Excel attachments are deeply parsed — great for lead lists
Teams scanning works best if you’re active in many channels
Use Clear Cache if switching between different M365 accounts
