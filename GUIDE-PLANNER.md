# Shaadi Planner — setup (about 20 minutes, one time)

> **Money update (vendor edit + Hisaab split) — do this (5 minutes):**
> 1. Google Sheet → **Extensions → Apps Script** → select all, delete, paste the new `Code.gs`. **Check the PIN line is still your PIN.** Save.
> 2. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.** (Same `/exec` link. Allow permission if asked.)
> 3. Replace `index.html` on GitHub (keep your `scriptUrl`). Vercel updates in about a minute.
> 4. Open the planner → tap **Refresh**. The first load adds a **Payments** tab and ID / Split / Share % columns to "Vendors & Payments", moves each existing advance into a payment, and adds about 19 after-wedding tasks. Nothing you already have is lost.

> **Already set up the first version? Do only this (5 minutes):**
> 1. Google Sheet → **Extensions → Apps Script** → select all, delete, paste the new `Code.gs`. **Check the PIN line is still your PIN.** Save.
> 2. **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.** (Same `/exec` link. If Google asks for permission, Advanced → Go to … → Allow.)
> 3. Open `index.html` from this folder, paste your `/exec` link into `scriptUrl: ""`, save.
> 4. On GitHub open your `shaadi-planner` repo → click `index.html` → ✏️ → select all, paste the new file → **Commit changes**. Vercel updates by itself in about a minute.
> (Search-bar update: only step 4 is needed — no script redeploy.)
> 5. Open the planner and tap Refresh. The first load adds the new columns and about 78 suggested tasks (one time only; nothing you already have is changed or duplicated).

Your Excel becomes a **Google Sheet** (the data lives there), and a small **web page** sits on top of it so the whole family can tick and update tasks from their phones.

Files in this folder:
- `index.html` — the planner page
- `Code.gs` — the small script that connects the page to your Google Sheet

---

## Part A — Turn your Excel into a Google Sheet
1. Go to **drive.google.com** (the same Google account you used for the wedding site).
2. Click **New → File upload** and choose `Wedding_Master_Tracker_Excel.xlsx`.
3. Double-click the uploaded file to open it, then click **File → Save as Google Sheets**.
4. In the new Google Sheet click the title and rename it **Wedding Master Tracker**. Use this new Google Sheet from now on, not the Excel.

## Part B — Connect the script
1. In that Google Sheet: **Extensions → Apps Script**.
2. Select all the code in the box, delete it, open `Code.gs` from this folder (Notepad is fine), copy everything and paste it in.
3. On the line `var PIN = '1202';` change `1202` to **your own family PIN** (numbers only). Click **Save** (disk icon).
4. Click **Deploy → New deployment → ⚙ → Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy → Authorize access →** your account. On "Google hasn't verified this app" click **Advanced → Go to … (unsafe) → Allow**.
6. Copy the **Web app URL** (ends with `/exec`).
7. Open the URL in a browser tab. It should say `{"success":false,"code":"PIN","message":"Wrong PIN"}`. That is correct — it means the script is running and is asking for the PIN.

## Part C — Put the page online (Vercel)
1. Open `index.html` in Notepad. Near the top find `scriptUrl: ""` and paste your `/exec` link between the quotes. Save.
2. On **github.com** click **+ → New repository**. Name it `shaadi-planner`, choose **Private**, click **Create repository**.
3. Click **uploading an existing file**, drag in `index.html`, then **Commit changes**.
4. On **vercel.com**: **Add New → Project →** pick `shaadi-planner → Import`. Leave every setting as it is (Framework "Other", no build command) and click **Deploy**.
5. You get an address like `shaadi-planner.vercel.app`. Open it on your phone, enter the PIN and your name.

**Share with family:** send them the address and the PIN on WhatsApp. On a phone they can tap the browser menu → **Add to Home screen** so it opens like an app.

## Part D — How to use it
- **Done button:** every open task has a green **Done · पूरा हुआ** button. It asks "Mark as done?" first, then shows an **Undo · वापस** button for a few seconds. Tap the task itself to change anything else.
- **Last 7 days:** on Today or in All Tasks, see what was marked done or in progress in the last week, grouped by day.
- **Search bar (top of every screen):** type anything you remember — a task word, vendor name, event (mehndi, haldi), a date (1 dec) or a person. Matching tasks appear at once with **Done** and **Started** buttons, and matching vendors show below. Done tasks are found too. Nothing found? Tap the button to add it as a new task.
- **Checklist tab:** every task grouped by function in date order (Before, Faldaan, Mehndi, Engagement & Sangeet, Haldi, Varmala & Shaadi, After). In each group the **critical** tasks come last. Change a task's function by tapping it.
- **Add a task:** the **＋ Add a new task** button is on Today and All Tasks. Only the task name is needed; "More options" is optional.
- **Today** (opens first): days to go, % done, **Overdue**, **This week**, and every **Critical (P0)** task that is still open.
- **Tap the circle** on a task to mark it done. **Tap the task** to change status, priority, who is doing it, due date and notes.
- **All Tasks:** search, filter by status, priority, category or person. **Checklist mode** shows big easy rows for the last days.
- **+ button:** add a new task.
- **Vendors:** quoted, paid and balance for each vendor, with a tap-to-call button.
- Everyone sees the same live data; the page refreshes by itself every 45 seconds. Each change records who made it and when.

## Important
- **Your tasks have no due dates yet**, so "This week" will be empty at first. Open "All Tasks → No date" and give each task a date; the P0 list works without dates.
- The script adds five columns on the right of the "Master Tracker" tab (ID, Owner, Due Date, Last Updated By, Last Updated). **Do not delete the ID column.** You can still edit, sort and add rows in the Google Sheet itself; new rows get an ID automatically.
- **Money tab:** add, edit or delete vendors and expenses from the page. Tap **＋ Payment** to record each payment and who paid (Shrijeet / Shivangi / Family). Balance = quoted − payments, so "I paid the advance, she paid the rest" works with two payments.
- **Hisaab · Split:** shared items are divided by the % on each item (50/50 by default). Mark clothes, jewellery and gifts as **Personal** and they are left out. Family-paid money is not divided. **Settle up** records when one of you pays the other; **Copy** makes a summary for WhatsApp.
- **Today → Critical:** open critical tasks are grouped by category, with a button to see every open task of that category.
- The Guest List and Event Flow tabs are not in the page yet.
- **Privacy:** the PIN keeps casual visitors out, but anyone who has both the link and the PIN can see vendor phone numbers and payments. Share it only with family. Change the PIN in `Code.gs` (then Deploy → Manage deployments → ✏️ → New version → Deploy) if it leaks.
- **If you change `Code.gs` later:** Deploy → Manage deployments → ✏️ → Version: New version → Deploy. The link stays the same.
