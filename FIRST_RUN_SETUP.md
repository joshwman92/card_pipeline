# L.U.C.A.S First Run Setup

This guide assumes you are setting up L.U.C.A.S on a Windows computer for the first time and do not already know what the app needs.

L.U.C.A.S stands for Lot Upload, Comping & Assignment System. It helps you create card sheets, comp cards through Card Ladder, receive cards, assign cards to buying companies, and track payouts.

## Quick Answer

For a complete L.U.C.A.S setup, the computer needs:

- Google Chrome
- the L.U.C.A.S project folder
- Python 3.11 or newer, installed automatically by `install_dependencies.bat` when possible
- a local data folder with `WORKING SHEETS`, `INCOMING SHEETS`, and `RECEIVED SHEETS`
- a Card Ladder account and Chrome extension setup for comping
- a Google AI Studio API key for Photo OCR and screenshot OCR fallback
- Google billing and spend controls for Google API calls
- Google OAuth credentials for Google Sheets rules or payout sheets
- Google Drive for desktop
- company assignment rules and payout files configured inside the app

This guide treats the full workflow as required. Setup is not complete until every section below is done.

## Important Setup URLs

Keep these official links open during setup:

| Need | URL |
| --- | --- |
| L.U.C.A.S repository | `https://github.com/mikegrossbarth/card_pipeline` |
| Google Chrome download | `https://www.google.com/chrome/` |
| Python download | `https://www.python.org/downloads/windows/` |
| Windows Package Manager help | `https://learn.microsoft.com/windows/package-manager/winget/` |
| Google Drive for desktop | `https://www.google.com/drive/download/` |
| Card Ladder | `https://app.cardladder.com/` |
| Chrome extensions page | `chrome://extensions` |
| Google AI Studio API keys | `https://aistudio.google.com/app/apikey` |
| Gemini API billing | `https://ai.google.dev/gemini-api/docs/billing` |
| Gemini API key security | `https://ai.google.dev/gemini-api/docs/api-key` |
| Google AI Studio spend cap | `https://aistudio.google.com/app/spend` |
| Google Cloud Console | `https://console.cloud.google.com/` |
| Google Cloud Billing | `https://console.cloud.google.com/billing` |
| Google Cloud APIs library | `https://console.cloud.google.com/apis/library` |
| Google Sheets API | `https://console.cloud.google.com/apis/library/sheets.googleapis.com` |
| Google Auth Platform | `https://console.cloud.google.com/auth` |
| Google OAuth clients | `https://console.cloud.google.com/auth/clients` |

## What Each Requirement Is For

| Requirement | Needed for | Required for complete setup? |
| --- | --- | --- |
| Google Chrome | Card Ladder automation and extension | Yes |
| Python 3.11+ | running the desktop app | Yes, but installer tries to install it |
| `.venv` dependencies | app libraries like spreadsheet and image support | Yes, created by installer |
| Working folder | saving and finding card sheets | Yes |
| Card Ladder account | comping cards with Card Ladder | Yes |
| Card Ladder Chrome extension | automated comp runs | Yes |
| `GOOGLE_API_KEY` | Photo OCR and Card Ladder screenshot OCR fallback | Yes |
| Google billing and spend controls | paid Gemini API calls, payment method, and cost alerts | Yes |
| Google OAuth Client ID/Secret | reading Google Sheets rule/payout files | Yes |
| Google Drive for desktop | synced Drive folders and Google Sheets shortcuts | Yes |
| assignment companies | best company and estimated payout | Yes |

## Step 1: Get The App Folder

Put the L.U.C.A.S project folder somewhere normal on the computer, for example:

```text
C:\Users\YourName\Documents\card_pipeline
```

The folder should contain files like:

```text
app.py
install_dependencies.bat
Run Team LUCAS.vbs
Run Michael LUCAS.vbs
Run LUCAS Mobile Stack.vbs
README.md
cardladder-autocomp
```

If you are using Git, clone the repository from:

```text
https://github.com/mikegrossbarth/card_pipeline
```

If not, open that same URL in a browser, download the project as a ZIP, and unzip the project folder.

## Shared Team Folder Setup

For a small team, every computer can use the same synced pipeline folder, for example:

```text
G:\My Drive\CARD_PIPELINE
```

Each computer still keeps its own local app folder, `.env`, Google OAuth token, and Card Ladder Chrome extension. The shared folder is only for pipeline data: working sheets, incoming sheets, received sheets, company sheets, payout markers, and profit records.

L.U.C.A.S creates a local `lucas_user_identity.json` on each computer. Do not copy this file between computers. It lets shared locks and profit records identify which user/machine made a change.

When several users share the same pipeline folder, L.U.C.A.S uses lock files in:

```text
CARD_PIPELINE\.locks
```

These locks protect shared workbook writes, receive updates, company-sheet appends, payout markers, and the profit ledger. If another user is writing, the app waits briefly instead of writing over them. If Google Drive creates a conflicted JSON copy, the Home refresh will show a warning so the team can resolve it.

All users should keep Google Drive for desktop running and fully synced before marking cards received, moving sheets, or marking payouts paid.

## Step 2: Install Google Chrome

Install Google Chrome if it is not already installed:

```text
https://www.google.com/chrome/
```

Chrome is needed for Card Ladder comping because the app talks to a local Chrome extension.

## Step 3: Run The Installer

Double-click:

```text
install_dependencies.bat
```

The installer will:

1. look for Python 3.11 or newer
2. try to install Python 3.11 with Windows Package Manager if Python is missing
3. create `.venv`
4. install app dependencies
5. verify Tkinter is available
6. create `.env` from `.env.example` if `.env` does not exist

If Windows asks for permission during Python install, approve it.

If automatic Python install fails, install Python manually from:

```text
https://www.python.org/downloads/windows/
```

During manual install, make sure:

- `Add python.exe to PATH` is checked
- Tcl/Tk support is included

Then run `install_dependencies.bat` again.

## Step 4: Create The Sheet Folders

Create a main folder for your pipeline data inside Google Drive for desktop.

Example:

```text
G:\My Drive\CARD_PIPELINE
```

Inside it, create:

```text
G:\My Drive\CARD_PIPELINE\WORKING SHEETS
G:\My Drive\CARD_PIPELINE\INCOMING SHEETS
G:\My Drive\CARD_PIPELINE\RECEIVED SHEETS
```

The app can create some folders as it works, but starting with these three makes setup clearer.

L.U.C.A.S will also use or create:

```text
G:\My Drive\CARD_PIPELINE\COMPANY SHEETS
G:\My Drive\CARD_PIPELINE\ASSIGNMENT RULES
G:\My Drive\CARD_PIPELINE\sheet_markers.json
```

## Step 5: Launch The App

For normal use, double-click:

```text
Run Team LUCAS.vbs
```

For Personal/Michael LUCAS, double-click:

```text
Run Michael LUCAS.vbs
```

If the app does not open or you want to see error messages, run this from PowerShell in the repo folder:

```powershell
.\.venv\Scripts\python.exe app.py
```

## Step 6: Choose The Working Folder

When L.U.C.A.S opens:

1. click `Working Folder` in the top-right area
2. choose your `WORKING SHEETS` folder
3. confirm the app sees your working sheets

Choose the folder named `WORKING SHEETS`, not the parent `CARD_PIPELINE` folder.

The app remembers this in:

```text
lucas_settings.json
```

That file is local to the computer and should not be committed to Git.

## Step 7: Configure `.env`

Open the `.env` file in the L.U.C.A.S project folder.

It starts as a copy of `.env.example`.

A typical full setup looks like:

```env
GOOGLE_API_KEY=your_google_ai_studio_key
GOOGLE_SHEETS_OAUTH_CLIENT_ID=your_desktop_oauth_client_id
GOOGLE_SHEETS_OAUTH_CLIENT_SECRET=your_desktop_oauth_client_secret
LUCAS_WORKING_SHEETS_DIR=G:\My Drive\CARD_PIPELINE\WORKING SHEETS
```

For a complete setup, fill in every value shown above.

Do not commit `.env` to Git.

## Step 8: Set Up Photo OCR And OCR Fallback

Photo OCR needs:

```env
GOOGLE_API_KEY=...
```

To get this key:

1. go to Google AI Studio:

```text
https://aistudio.google.com/app/apikey
```

2. create or copy an API key
3. paste it into `.env`
4. save `.env`
5. restart L.U.C.A.S

This key is used for:

- scanning photos into card rows
- Card Ladder screenshot OCR fallback

If this key is missing, setup is incomplete. Photo OCR and Card Ladder screenshot OCR fallback will not work.

## Step 9: Set Up Google Billing, Payment, And Spend Controls

L.U.C.A.S uses the Gemini API for Photo OCR and Card Ladder screenshot OCR fallback. Those calls can use paid Google API capacity, so the Google account must have billing, payment, and spending controls set up before real use.

Use the same Google account and project that created the `GOOGLE_API_KEY`.

### Set Up Billing In Google AI Studio

1. Open the official Google AI Studio API keys page:

```text
https://aistudio.google.com/app/apikey
```

2. Find the project connected to the L.U.C.A.S API key.
3. In the billing column, click `Set up billing`.
4. Follow Google's prompts to select or create a Cloud Billing account.
5. Add or confirm the payment method.
6. If Google asks for prepaid credits, add the required minimum credit amount. Google's Gemini billing docs currently describe this as a minimum of `$10` or the equivalent in another currency.
7. Return to the AI Studio billing page and confirm the project no longer says `Set up billing`, `Set up Prepay`, or `No credits`.

Google's billing guide says upgrading to a paid tier means linking a billing account and, where required, prepaying credits. It also says the Gemini API only serves paid-tier requests when the billing/prepay state is valid.

Official references:

- [Gemini API billing](https://ai.google.dev/gemini-api/docs/billing)
- [Google Cloud Billing account management](https://docs.cloud.google.com/billing/docs/how-to/manage-billing-account)

### Add A Payment Method

1. Open Google Cloud Console:

```text
https://console.cloud.google.com/billing
```

2. Select the Cloud Billing account used by the L.U.C.A.S project.
3. Open `Payment method`.
4. Add a valid card, debit card, or bank account.
5. Confirm the billing account is active.

Google notes that projects not linked to an active Cloud Billing account cannot use Google Cloud services, even some services that are otherwise free.

### Set A Monthly Spend Cap

1. Open the AI Studio spend page:

```text
https://aistudio.google.com/app/spend
```

2. Select the L.U.C.A.S project.
3. Find `Monthly spend cap`.
4. Click `Edit spend cap`.
5. Set a cap appropriate for the user.
6. Save.

Start conservative. A practical first cap is usually enough to cover testing and normal OCR volume, then increase later after real usage is understood.

Google's Gemini billing docs say project-level spend caps are available in AI Studio, but there can be a short billing-processing delay. That means a cap helps reduce surprise costs, but it is not a perfect instant shutoff.

### Create A Billing Budget Alert

1. Open Google Cloud Billing:

```text
https://console.cloud.google.com/billing
```

2. Select the billing account.
3. Open `Budgets & alerts`.
4. Create a budget for the L.U.C.A.S project.
5. Add alert thresholds, for example:

```text
50%
75%
90%
100%
```

6. Make sure the owner or operator's email receives billing alerts.

Budget alerts do not replace spend caps. They are warning lights so the user knows costs are moving.

### Secure The Gemini API Key

Treat the Gemini API key like a password.

1. Keep it only in `.env`.
2. Do not paste it into chats, spreadsheets, screenshots, Git commits, or shared docs.
3. Do not commit `.env`.
4. In Google AI Studio, prefer a current auth key.
5. Restrict keys where Google provides a restriction option.
6. If the key may have leaked, create a new key, update `.env`, then disable the old key.

Google's API-key docs warn that leaked keys can consume quota, create unexpected billing charges, and access private resources. They also recommend billing alerts and API-key restrictions.

Official reference:

- [Gemini API key security](https://ai.google.dev/gemini-api/docs/api-key)

After finishing this section, restart L.U.C.A.S so `.env` and the active Google API key are loaded cleanly.

## Step 10: Set Up Card Ladder

Card Ladder comping needs:

- Google Chrome
- a Card Ladder account
- a logged-in Card Ladder session in Chrome
- the L.U.C.A.S Chrome extension loaded

To load the extension:

1. open Chrome
2. go to:

```text
chrome://extensions
```

3. turn on `Developer mode`
4. click `Load unpacked`
5. select this folder inside the L.U.C.A.S project:

```text
cardladder-autocomp\extension
```

6. open Card Ladder in Chrome
7. log in to Card Ladder

The desktop app starts the local Card Ladder bridge automatically when it opens. The extension talks to that local bridge.

If comping does not start, check:

- Chrome is open
- the extension is enabled
- you are logged into Card Ladder
- L.U.C.A.S is running
- the Comp tab says the Card Ladder bridge is running

## Step 11: Set Up Google Sheets Access

Google Sheets access is part of the complete setup because company rules and payout files may live in Google Sheets.

1. go to Google Cloud Console:

```text
https://console.cloud.google.com/
```

2. create or choose a project
3. open the Google Sheets API page:

```text
https://console.cloud.google.com/apis/library/sheets.googleapis.com
```

4. click `Enable`
5. open Google Auth Platform:

```text
https://console.cloud.google.com/auth
```

6. complete the app setup prompts if Google asks for them
7. open OAuth clients:

```text
https://console.cloud.google.com/auth/clients
```

8. create an OAuth Client ID
9. choose application type:

```text
Desktop app
```

10. copy the client ID and client secret into `.env`:

```env
GOOGLE_SHEETS_OAUTH_CLIENT_ID=...
GOOGLE_SHEETS_OAUTH_CLIENT_SECRET=...
```

11. save `.env`
12. restart L.U.C.A.S
13. open Company Rules
14. click `Connect Google`
15. sign in with the Google account that can access the sheets

The app creates:

```text
lucas_google_sheets_token.json
```

That token file is local and should not be committed to Git.

## Step 12: Set Up Google Drive For Desktop

Google Drive for desktop is required for the standard team setup because the pipeline folders and source files can live in synced Drive folders.

Download it here:

```text
https://www.google.com/drive/download/
```

Once installed, Google Drive usually creates a drive like:

```text
G:\My Drive
```

You can put `CARD_PIPELINE` there.

Native Google Sheets files may appear as `.gsheet` shortcuts. If L.U.C.A.S cannot read a `.gsheet` directly as a local workbook, it will use the saved Google Sheet URL and Google OAuth setup to read the live sheet.

For the easiest rule and payout setup, use synced `.xlsx` or `.csv` files when possible.

## Step 13: Set Up Company Rules

Assignment recommendations need company rules and payout rates.

In L.U.C.A.S:

1. open the `Assignment` tab
2. click `Company Rules`
3. create a company
4. choose whether the company is active
5. choose a rule source
6. choose a payout source
7. save the company

Each company can use:

- manual rules
- local rules file
- Google Keep export/text file
- local workbook or CSV
- Google Sheet through OAuth
- payout tiers inside the same workbook on a `Payouts` tab

For manual rules, use **Company Card Year Range** when a company has a broad year cutoff. For example, set Fanatics **Min Year** to `1990` so pre-1990 cards are rejected before sport/value rules are evaluated.

For Google Keep rule sources, sync is done per company:

```text
Company Rules -> select company -> Rule Source -> Sync Google Keep
```

Use this when a Keep note changed and L.U.C.A.S needs the Chrome extension to refresh its cached note text.

The app saves company setup locally in:

```text
assignment_companies.json
```

That file is local and should not be committed unless you intentionally want to share it.

### Optional: Set Up People Rules For Network Mode

Use this only if the computer will buy cards from sellers with preset terms.

Open:

```text
Company Rules -> People Rules
```

People Rules writes shared seller terms to:

```text
CARD_PIPELINE\ASSIGNMENT RULES\seller_terms.csv
```

Each row uses:

```text
Person, Sheet Type, Min Value, Max Value, Seller Rate, Deduction, Balance Share
```

Legacy CSV files with a `Seller` header still load, but the app now labels that field **Person**. In the People Rules popup, the percentage fields are labeled **Seller Rate %**, **Deduction %**, and **Balance Share %**. Type numbers only, without `%`. Decimals are allowed.

Examples:

```text
90
92.5
10
10.5
```

Use **Seller Rate %** for a flat percent of the Sheet Type company's source value. Use **Deduction %** to follow the Sheet Type company's payout logic and subtract that percentage from the company payout. Use one or the other on a seller row, not both. Add multiple rows for the same person and sheet type only when their Min/Max Value ranges do not overlap. Leave **Sheet Type** blank and fill **Balance Share %** to define a normal team profit-share payout, for example `50` for Kevin or James; that does not make the person a seller source.

## Step 14: Check The Full Workflow

After setup, test the app in this order:

1. Create a small working sheet in the `Create` tab.
2. Save it to `WORKING SHEETS`.
3. Open the `Comp` tab and make sure the sheet appears.
4. Run one Card Ladder comp as a test.
5. Open Company Rules and make sure companies load.
6. If Network Mode is needed, open People Rules and make sure People Rules Health is clean.
7. Open the `Assignment` tab and confirm best company and estimated payout can populate.
8. Open the `Receive` tab and test marking a row received.
9. Open `Payouts/Tabs` and confirm assigned-person balances appear when relevant.
10. Open `Profit` and confirm sold cards/sold sheets appear after company sheets have sold rows.

Do not start with a giant sheet until this small test works.

## Files That Stay Local

These files are expected to be different on every computer:

```text
.env
.venv
lucas_settings.json
lucas_google_sheets_token.json
assignment_companies.json
sheet_markers.json
```

Do not send these to Git unless you know exactly why you are doing it.

## Common Problems

### The app does not open

Run:

```powershell
.\.venv\Scripts\python.exe app.py
```

Read the console message. Most startup issues are missing Python, missing dependencies, or a bad `.env` value.

### The installer says Python is missing

Run `install_dependencies.bat` again. If it still cannot install Python automatically, install Python 3.11 or newer from python.org and check `Add python.exe to PATH`.

### Photo OCR says the Google key is missing

Add this to `.env`:

```env
GOOGLE_API_KEY=...
```

Then restart the app.

### Google Sheets cannot connect

Check:

- Google Sheets API is enabled
- OAuth client type is `Desktop app`
- `.env` has both client ID and client secret
- you clicked `Connect Google`
- the signed-in Google account has access to the sheet

### Card Ladder comping does not start

Check:

- Chrome is open
- Card Ladder is logged in
- the unpacked extension is enabled
- L.U.C.A.S is open
- the local bridge status is running

### Sheets are not showing up

Check that the app is pointed at the actual:

```text
WORKING SHEETS
```

folder, not the parent folder.

## Setup Checklist

Use this as the final handoff checklist:

- [ ] Project folder exists on the computer
- [ ] Google Chrome is installed
- [ ] `install_dependencies.bat` completed successfully
- [ ] `.env` exists
- [ ] `WORKING SHEETS`, `INCOMING SHEETS`, and `RECEIVED SHEETS` folders exist
- [ ] L.U.C.A.S opens
- [ ] `Working Folder` points to `WORKING SHEETS`
- [ ] `GOOGLE_API_KEY` is added
- [ ] Google AI Studio billing is connected to the API key project
- [ ] Google Cloud payment method is active
- [ ] Gemini monthly spend cap is set
- [ ] Google Cloud budget alerts are set
- [ ] Gemini API key is secured and stored only in `.env`
- [ ] Card Ladder extension is loaded
- [ ] user is logged into Card Ladder in Chrome
- [ ] Google OAuth credentials are added
- [ ] `Connect Google` has been completed
- [ ] Assignment companies are created and active
- [ ] People Rules are configured if Network Mode seller pricing is needed
- [ ] one small test sheet works end to end
