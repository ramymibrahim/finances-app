# Finances app: setup guide

The setup has three parts and takes about 20 minutes. You only do it once.

- **apps-script/**: the backend. It goes into your Google Sheet.
- **app/**: the app itself. It goes onto GitHub Pages.

The backend is set to the test sheet **"Finances (app test)"**, so you can try everything without touching your real Budget 2026 file.

---

## Part 1: The backend (Google Apps Script)

1. Open the **Finances (app test)** sheet in Google Drive on your computer.
2. Click **Extensions → Apps Script**. A new tab opens.
3. Delete everything in the `Code.gs` file there and paste in the contents of **apps-script/Code.gs**.
4. Click the ⚙️ **Project Settings** icon on the left. Tick **"Show appsscript.json manifest file in editor"** and set the time zone to **(GMT+01:00) Zurich**.
5. Go back to the **Editor** (`< >` icon). Open `appsscript.json` and replace its contents with **apps-script/appsscript.json**. Press **Ctrl+S / Cmd+S** to save.
6. In the function menu at the top, choose **setup** and click **▶ Run**.
   - Google asks for permission. Click **Review permissions** and choose your account.
   - You'll see **"Google hasn't verified this app"**. This is normal for your own scripts. Click **Advanced → Go to (project name) (unsafe)**, then **Allow**.
7. The **Execution log** opens at the bottom and shows `Your app key: ...`. **Copy that key** and keep it private, like a password.
8. Click **Deploy → New deployment**. Click the ⚙️ next to "Select type" and choose **Web app**.
   - Description: `Finances`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then **copy the Web app URL** (it ends in `/exec`).

"Anyone" only means the address can be reached. Without your key, the script refuses every request.

**Later changes to Code.gs:** Deploy → **Manage deployments** → ✏️ edit → Version: **New version** → Deploy. The URL stays the same.

---

## Part 2: Put the app on GitHub Pages

1. Create a free account at **github.com** if you don't have one.
2. Click **+ → New repository**. Name it `finances`, choose **Public**, tick **Add a README file**, then click **Create repository**.
3. Click **Add file → Upload files**. Drag in **everything inside the `app` folder**: `index.html`, `app.js`, `styles.css`, `sw.js`, `manifest.webmanifest`, and the `icons` folder. Drag the files themselves, not the `app` folder. Click **Commit changes**.
4. Go to **Settings → Pages**. Under "Build and deployment", choose Source **Deploy from a branch**, Branch **main**, folder **/ (root)**, then click **Save**.
5. After 1–2 minutes your app is live at:
   `https://YOUR-USERNAME.github.io/finances/`

Only the app's code is public. Your money data stays in your Google Sheet, and the key never goes into the code.

**Updating the app later:** upload the changed files the same way. Also change `VERSION` at the top of `sw.js` (for example to `finances-v2`) so phones load the new files.

---

## Part 3: Install on iPhone and computer

**iPhone**
1. Open the app address in **Safari** (it must be Safari).
2. Tap the **Share** button → **Add to Home Screen** → **Add**.
3. Open **Finances** from the home screen. It goes straight to the **More** screen.
4. Paste the **Web app URL** and the **App key**, then tap **Save and test**. You should see ✓ Connected.
   - Tip: send yourself the URL and key in Apple Notes from your computer, then copy them on the phone.

**Computer**
Open the same address in Chrome, Edge or Safari, and enter the URL and key once. In Chrome or Edge you can also click the install icon in the address bar to get a desktop app.

---

## Using it

- **+**: add an expense, income or transfer. Merchants you've used before fill in the category, account, currency and item automatically.
- **📷 Add receipt**: take a photo, or pick one from Photos or Files (PDF works too). Photos are made smaller before upload. They're saved in the **Finances Receipts** folder in your Google Drive and linked in the Log's **Receipt** column.
- **No signal?** Entries are kept on the phone and uploaded automatically when you're back online. You can see them in Activity marked ⏳.
- **Edit or delete**: in **Activity**, tap any expense, income or transfer. Change it and tap **Save changes**, or tap **Delete entry**. Before changing anything, the app checks that the row in the sheet still holds that entry; if you edited it in the sheet in the meantime, the app asks you to refresh. Deleting clears the row's input cells and keeps its formulas, and a receipt file stays in your Drive folder.
- **Accounts** (More screen): add an account, or tap one to change its name, type or starting balance. A new name is updated everywhere the account is used (Log, Income, Transfers, Planned). An account that has entries can't be deleted or change currency; rename it instead.
- **Categories and budgets** (More screen): add a category, change a monthly budget, or rename one. A new name is updated in the Log, Planned and Items. A category that has entries can't be deleted.
- **Bank rate**: type it only when you want the exact rate your bank used. Otherwise the sheet's rate for that date is used, as before.

## Money left over (Plan tab, at the top)

Everything in CHF, for the next 24 months:
**income expected − planned payments − spending − set aside for funds and goals + paid from funds − savings = left over.**

- **Free money now** = all your accounts in CHF (card and overdraft debt count as minus) minus what is set aside in funds and goals. Each month's "left over" is added to it: the chart shows your free money at the end of each month.
- **Two lines:**
  - **Following your budget** (blue): spending = each category's budget minus what is already planned in it that month (never below 0). For the current month: only what is still left in each budget.
  - **At your real pace** (orange, dashed): spending = what you really spend outside planned payments. First from this month's records (rough until you have a full month), then your average of the last 1–3 full months. Categories above budget are marked ▲ under "Budget and pace by category".
- **Switch** "Following budget / Your real pace" to see the table for either line; tap a month for its lines.
- **Transfers between your own accounts don't count** (CHF → EGP, paying off a card).
- **Savings** is the "Monthly savings goal" in your Cash flow tab; change it with **Savings … ✎**.
- A sinking fund linked to a bill pays for that bill in its month ("Paid from funds").

There is no per-account estimate any more. Coming up and Recurring and planned work as before.

## The Plan tab

- **Coming up (45 days):** tap **Paid** (or **Received**) to add it to your records with one tap. You can change the amount, date and account first. **Skip** moves the item on without recording anything.
- **Paying off credit cards:** add a monthly **Transfer** from your bank account to the card and tick **Pay the full balance**. The amount in Coming up is what the card owes today (marked ≈); when you tap Paid, type what you actually paid. In the sheet this is column R of Planned (**Pay full balance** = Yes).
- **Recurring and planned:** everything in your Planned tab. Tap an item to change it, or **+ Add**. Use **Money in** for salary or money owed to you, and **Transfer** for moving money between your accounts, for example a monthly CHF → EGP transfer or paying off a credit card.

**What the app changes in your sheet the first time it runs:**
- **Planned tab:** four new columns. **Last paid** (the app moves this forward when you tap Paid or Skip), **Ends on**, **If weekend** (Earlier / Later / No change), and **To account** (for transfers). Rows already marked Paid get their first due date as Last paid.
- **Salary:** a monthly "money in" row for your salary is added from your latest Salary in the Income tab. It is moved to the Friday before when the 20th is on a weekend.
- **Accounts tab:** a new **Safety cushion** column (L).
- In the Planned tab, a minus amount means money in, as you already use it.

## Sinking funds and goals (Plan tab)

The money stays in your accounts. The app only keeps track of how much you have set aside, so you see what is really free to spend (shown under Accounts on Home).

- **Sinking fund:** for a bill that comes once or a few times a year, like the half-fare card, Axa or the apartment instalments. The app lists such bills under "Bills that come less than monthly" with a **Create fund** button. Link the fund to the bill. When you tap **Paid** on that bill, it is paid from the fund and the fund's target moves to the next due date.
- **Goal:** for something you are saving towards, like an emergency fund.
- Each one shows how much to set aside each month to reach the target on time. Tap **Set aside** once a month. You can also type your own monthly amount; the app warns you if it isn't enough.
- Tap a fund or goal to change it, or to **take money out** when you used part of it.
- These are kept in a new **Envelopes** tab in your sheet (created automatically).

## Check against your bank (More → Accounts)

Tap an account, type the balance your bank shows, and tap **Check**. If the numbers differ, add the missing entries yourself, or tap **Add correction**. A correction is saved as an expense or income with the category **Balance correction**, which the app leaves out of your budget, spending and income totals. The date you last checked is shown under each account.

## Transfer costs (More)

For every transfer between currencies where you filled in **Amount received**, the app compares what arrived with the market rate that day (from your Rates tab). It shows the cost in % and in CHF, the total for this year, and which route is cheapest. The cost also shows on each transfer in Activity. When a day's rate is missing, today's rate is used and the cost is marked "approx.".

## Search, category details and duplicating

- **Search (Activity):** type to search all years by merchant, item, note, account or amount. Tap ⚲ for filters (type, category, account, dates, amounts). Tap a result to edit it.
- **Category details (Home):** tap a category under "Budget by category" to see the last 12 months against the budget, where the money went, and this month's entries.
- **Duplicate:** open any entry and tap **Duplicate as new entry**. You get a copy with today's date, ready to change and save.

## iPhone Reminders: "To Pay" list

A shortcut copies payments and transfers due in the next 7 days (plus unpaid overdue ones) into the Reminders list **To Pay**. Each reminder has an ID on the last line of its notes (`FIN-…`), which prevents duplicates. When you mark a payment **Paid** or **Skip** in the app, the next sync ticks the reminder off. If you delete a planned item or move it later, its reminder is removed. Income is not added. Ticking a reminder off on the phone does **not** record the payment; use **Paid** in the app for that, because it records the real amount. The full shortcut steps are in the chat; the backend action is `reminderSync`.

## Installing an update

When you get a new version:
1. **Backend:** replace all of `Code.gs` in Apps Script, save, then **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. The URL and key stay the same.
2. **App:** on GitHub, open your `finances` repository and click **Add file → Upload files**. Drag in the new files from the `app` folder; files with the same name are replaced. Click **Commit changes**.
3. On the iPhone, close the app fully (swipe it away) and open it again. If it still looks old, wait a minute and repeat.

## Moving to your real sheet

Note: your Planned tab is a Sheets table with dropdown columns. In the test sheet the dropdowns were widened for the app (Repeats: Every 2 months / Quarterly / Every 6 months; Status: Skipped; dropdowns for If weekend, To account and Pay full balance). Do the same in the real sheet when you switch (ask Claude, or edit the column's dropdown in the table).


When you're happy with the test, change `SHEET_ID` at the top of `Code.gs` to your real sheet's ID. The ID is the long part of the sheet's address between `/d/` and `/edit`. Also add a **Receipt** header in column M of the real Log tab, then deploy a new version as above.

## If something goes wrong

| Message | What to do |
|---|---|
| Wrong app key | Paste the key again in Settings, with no spaces. |
| Unexpected reply from the server | The URL is wrong, or it isn't deployed with access "Anyone". |
| Missing … / Amount must be more than 0 | A required field is empty. |
| Nothing new shows in the app | Switch to another app and back, or close and reopen it. It refreshes every time it opens. |
