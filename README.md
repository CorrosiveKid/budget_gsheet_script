# 2Up Budget Tracker

A Google Apps Script that turns a Google Sheet of synced transactions into
monthly budget tabs. Each tab covers one budget cycle, lists your budget
categories, and automatically totals what you've spent against each one.

## How it works

- **Cycles** run from the 14th of one month to the 14th of the next
  (`CYCLE_DAY`). The end date is exclusive: a `14 Oct - 14 Nov 2026` tab covers
  14 Oct through 13 Nov.
- **Config is the source of truth.** Budgets are read from the `Config` tab
  each time a new month tab is created. The hidden `Template` tab is rebuilt
  from Config first, then copied to make the new tab.
- **Spending is matched by tag.** Each transaction in the `2Up Spending` tab
  is counted against a category when one of its tags exactly matches the
  category name (case-insensitive).
- **New tabs chain on from the latest one.** The next cycle starts at the end
  date of the most recent month tab, not today's date, so a missed or early
  run still produces the right cycle.

## Sheet layout

The spreadsheet needs these tabs:

### `2Up Spending` — transactions

Filled in by your transaction sync. Row 1 is a header; data starts at row 2.
The script reads these columns:

| Column | Meaning |
| ------ | ------- |
| B | Transaction date |
| D | Amount (negative = spend, positive = income) |
| G | Tags, comma-separated |
| H | **Override** — if filled in, used instead of the tags in G |
| I | **Split** — if filled in, used instead of the amount in D |

The split keeps the transaction's original direction, so on a spend row you
can type your share as either `50` or `-50` and it counts as spend either way.

New transactions can be inserted at the top of the sheet; the formulas use
`INDIRECT` ranges so they keep counting every row.

### `Config` — budgets

Row 1 is a header; one line item per row from row 2:

| Column | Meaning |
| ------ | ------- |
| A | Category (must match the transaction tag) |
| B | Line item (description only) |
| C | Amount |

A category can span several rows (for example, several loans under `Loans`);
its budget is the sum of those rows. Categories appear on the month tab in the
order they first appear in Config.

### `Template`

An empty tab named `Template`. The script overwrites it each time it creates a
month tab, so don't edit it by hand. You can hide it.

### Month tabs (generated)

Each generated tab looks like this:

| | A | B | C | D |
| - | - | - | - | - |
| 1 | Cycle | | | |
| 2 | Start | *start date* | End | *end date* |
| 3 | Expected Incoming | *you fill in* | Actual Incoming | *auto* |

Then from row 5, one row per category:

| Category | Budget | Remaining | Auto Usage | Manual Usage | Total Usage | % Used |
| -------- | ------ | --------- | ---------- | ------------ | ----------- | ------ |

- **Auto Usage** — spend in this cycle tagged with the category.
- **Manual Usage** — anything you want to add by hand (starts at 0).
- **Actual Incoming** — all positive amounts in this cycle.
- A `TOTAL` row sums the columns at the bottom.

## Setup

1. Create the `2Up Spending`, `Config` and `Template` tabs as described above.
2. Open **Extensions > Apps Script**, paste in the contents of
   `budget_system.gs`, and save.
3. Check **Project Settings > Time zone** matches yours; tab names and dates use
   the script's time zone.
4. Reload the spreadsheet. A **Budget** menu appears.
5. Click **Budget > Create Next Month Tab** and approve the permissions prompt.

If there are no month tabs yet, the first one is for the cycle that contains
today.

## Usage

**Budget** menu:

- **Create Next Month Tab** — creates the next cycle's tab after the latest
  one. Does nothing if that tab already exists.
- **Enable Auto-Create the day before each cycle** — adds a time-based trigger
  that creates the next tab on the 13th of each month at around 8pm. Running it
  again replaces the existing trigger rather than adding a second one.

### Changing budgets

Edit `Config`. The change applies to the **next** tab created; existing month
tabs keep the budgets they were created with, so edit their Budget column
directly if needed.

### After changing the formulas

Existing month tabs keep the formulas they were created with. If you change any
of the formula builders in the script, run `refreshAutoUsageFormulas` from the
Apps Script editor to re-apply the Auto Usage and Actual Incoming formulas to
every month tab.

## Configuration

Constants at the top of `budget_system.gs`:

| Constant | Default | Meaning |
| -------- | ------- | ------- |
| `SOURCE_SHEET` | `2Up Spending` | Tab holding the synced transactions |
| `CONFIG_NAME` | `Config` | Tab holding the budgets |
| `TEMPLATE_NAME` | `Template` | Tab used as the month-tab template |
| `CYCLE_DAY` | `14` | Day of the month each cycle starts on |

If you change `CYCLE_DAY`, run **Enable Auto-Create** again so the trigger moves
to the new day.

## License

Apache License 2.0 — see [LICENSE](LICENSE).
