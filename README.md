# expense-entry-automation
An automated workflow that works across Google Workspace (Sheets, Forms, Gmail). Created to improve the user experience of entering expenses into my personal finance gsheet from my phone while I'm on the go.

# Mobile Quick-Add for a Zero-Based Budget (Google Apps Script)

A step-by-step guide to turning a Google Form into a fast, mobile-friendly front-end for a zero-based budget spreadsheet — with automatic sheet insertion, an overbudget email alert, and a self-updating category dropdown.

## What you'll end up with

- A Google Form on your phone's homescreen for logging expenses in seconds
- Form submissions land directly in your existing transactions sheet, in the correct columns — no manual copy-paste
- An automatic email whenever a budget category goes over
- A category dropdown that stays in sync with your budget categories, with no manual editing

## Prerequisites

- A budget spreadsheet with a transactions tab (columns for Date, Account, Category, Transaction detail, Outflow, Inflow, Remarks)
- A "helper" tab or range listing your valid budget categories
- A tab with budget-vs-actual totals per category, including some kind of "amount remaining" column

---

## Step 1: Build the Google Form

1. Go to [forms.google.com](https://forms.google.com) → **Blank form**. Title it something like "Quick Expense Entry."
2. Add a **Date** question, type **Date**. Leave **Required** off — this will default to today automatically if left blank.
3. Add an **Account** question, type **Dropdown**. Enter your account names as choices (e.g. Cash, Bank, Credit Card).
4. Add a **Category** question, type **Dropdown**. Paste in your current category list as a starting point — this will be kept in sync automatically later.
5. Add a **Transaction Name** question, type **Short answer**.
6. Add a **Cost** question, type **Short answer** → click the ⋮ menu → **Response validation** → **Number** → **Greater than 0**.
7. Add a **Notes (optional)** question, type **Paragraph**. Leave **Required** off.
8. Go to the **Responses** tab → click the green Sheets icon → **Select existing spreadsheet** → choose your budget spreadsheet. This creates a "Form Responses" tab automatically.

## Step 2: Connect the Form to your transactions sheet

In your spreadsheet, go to **Extensions → Apps Script** and paste in the following. Adjust the sheet name, start row, and column numbers to match your own layout.

```javascript
const TX_SHEET_NAME = "Transactions";
const TX_START_ROW = 13;   // first real data row
const TX_DATE_COL = 2;     // column B

function onFormSubmit(e) {
  const responses = e.namedValues; // { "Question Title": ["answer"] }

  const date     = getAnswer(responses, "Date");
  const account  = getAnswer(responses, "Account");
  const category = getAnswer(responses, "Category");
  const detail   = getAnswer(responses, "Transaction Name");
  const cost     = getAnswer(responses, "Cost");
  const remarks  = getAnswer(responses, "Notes (optional)");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(TX_SHEET_NAME);

  const finalDate = date
    ? date
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");

  const targetRow = findNextEmptyRow(txSheet);

  txSheet.getRange(targetRow, 2).setValue(finalDate);   // DATE
  txSheet.getRange(targetRow, 3).setValue(account);     // ACCOUNT
  txSheet.getRange(targetRow, 4).setValue(category);    // CATEGORY
  txSheet.getRange(targetRow, 5).setValue(detail);      // TRANSACTION DETAIL
  txSheet.getRange(targetRow, 6).setValue(Number(cost)); // OUTFLOW
  txSheet.getRange(targetRow, 8).setValue(remarks);     // REMARKS

  SpreadsheetApp.flush();
  checkBudgetAlerts();
}

function getAnswer(namedValues, questionTitle) {
  return namedValues[questionTitle] ? namedValues[questionTitle][0] : "";
}

function findNextEmptyRow(sheet) {
  const numRows = sheet.getMaxRows() - TX_START_ROW + 1;
  const colValues = sheet
    .getRange(TX_START_ROW, TX_DATE_COL, numRows, 1)
    .getValues();

  for (let i = 0; i < colValues.length; i++) {
    if (colValues[i][0] === "" || colValues[i][0] === null) {
      return TX_START_ROW + i;
    }
  }
  return TX_START_ROW + colValues.length;
}
```

Save the project (give it a name), then set up the trigger:

- Click the **clock icon (Triggers)** → **+ Add Trigger**
- Function: `onFormSubmit`
- Event source: **From spreadsheet**
- Event type: **On form submit**
- Save

Run `onFormSubmit`-dependent functions like `checkBudgetAlerts` manually once from the editor (see Step 3) to grant permissions upfront, so the trigger doesn't fail on its first real run.

## Step 3: Overbudget email alert

Add this to the same script file. Adjust the sheet name and columns to match your totals tab.

```javascript
const BUDGET_SHEET_NAME = "Budget";
const BUDGET_START_ROW = 9;
const CATEGORY_COL = 6;   // column with category names
const AVAILABLE_COL = 11; // column with remaining budget

function checkBudgetAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName(BUDGET_SHEET_NAME);

  const maxRows = budgetSheet.getMaxRows() - BUDGET_START_ROW + 1;
  const categories = budgetSheet
    .getRange(BUDGET_START_ROW, CATEGORY_COL, maxRows, 1)
    .getValues();
  const availables = budgetSheet
    .getRange(BUDGET_START_ROW, AVAILABLE_COL, maxRows, 1)
    .getValues();

  const overBudget = [];
  for (let i = 0; i < categories.length; i++) {
    const category = categories[i][0];
    if (category === "" || category === null) break;

    const available = availables[i][0];
    if (typeof available === "number" && available < 0) {
      overBudget.push({ category: category, amount: available });
    }
  }

  if (overBudget.length === 0) return;

  const lines = overBudget.map(item =>
    `- ${item.category}: over by ${Math.abs(item.amount).toFixed(2)}`
  );

  const subject = `Budget Alert: ${overBudget.length} categor${overBudget.length === 1 ? "y" : "ies"} overbudget`;
  const body = `The following categories are over budget:\n\n${lines.join("\n")}\n\nCheck your Budget tab for details.`;

  MailApp.sendEmail(Session.getActiveUser().getEmail(), subject, body);
}
```

To authorize this before relying on the trigger: select `checkBudgetAlerts` in the function dropdown at the top of the editor and click **Run**. Approve the permissions prompt (Advanced → Go to [project name] → Allow). This is required once per new service the script uses (`MailApp`, `Session`, etc.) — running manually first prevents a silent failure on the real trigger.

## Step 4: Auto-sync the Category dropdown from your helper sheet

This keeps the Form's Category dropdown matched to your budget categories automatically, filtering out any placeholder rows (categories are identified here by having an emoji as their first character — adjust this rule to whatever distinguishes a real category from a placeholder in your own sheet).

```javascript
const FORM_URL = "https://docs.google.com/forms/d/YOUR_FORM_ID/edit"; // use the edit URL, not the viewform/share link
const HELPER_SHEET_NAME = "helper_sheet";
const HELPER_START_ROW = 6;
const CATEGORY_QUESTION_TITLE = "Category";

function syncCategoryDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME);

  const maxRows = helperSheet.getMaxRows() - HELPER_START_ROW + 1;
  const values = helperSheet
    .getRange(HELPER_START_ROW, 1, maxRows, 1)
    .getValues();

  const EMOJI_PREFIX = /^\p{Extended_Pictographic}/u;

  const categories = [];
  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    if (cell === "" || cell === null) break;

    const text = String(cell).trim();
    if (EMOJI_PREFIX.test(text)) {
      categories.push(text);
    }
  }

  if (categories.length === 0) {
    throw new Error("No emoji-prefixed categories found — check sheet contents before this overwrites the dropdown.");
  }

  const form = FormApp.openByUrl(FORM_URL);
  const items = form.getItems();

  const categoryItem = items.find(item => item.getTitle() === CATEGORY_QUESTION_TITLE);
  if (!categoryItem) {
    throw new Error(`No form question titled "${CATEGORY_QUESTION_TITLE}" found.`);
  }

  categoryItem.asListItem().setChoiceValues(categories);
}
```

Run `syncCategoryDropdown` manually once from the editor to authorize the new `FormApp` service and confirm the dropdown updates correctly.

Then set up the schedule:

- **Triggers** → **+ Add Trigger**
- Function: `syncCategoryDropdown`
- Event source: **Time-driven**
- Type: **Month timer**, day **1**, any hour range
- Save

## Step 5: Add the Form to your homescreen

Open the Form's live share link (the `viewform` link, from the **Send** button in the Form editor) on your phone, then use your phone's "Add to Home Screen" browser option so it opens like an app.

---

## Result

Submitting the Form now: inserts a new row into your transactions sheet in the right columns, triggers your existing totals formulas to recalculate, emails you if any category is over budget, and keeps its own category list in sync with your budget categories every month — all without opening the spreadsheet.
