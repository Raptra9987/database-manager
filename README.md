# Database Manager

A Windows desktop tool that merges many files into a single Excel database.
It remembers which file — and whose number — every row came from, and marks
repeated entries in red.

It is **not limited to contacts**: the columns come from a schema, so it works
just as well for products, students, or any list you define yourself.

The interface speaks **English, Russian and Uzbek**, and runs on Windows 10
and Windows 11.

## Download

Grab the latest build from **[Releases](../../releases)** — a single
`Database Manager.exe`, no installer.

- Windows 10 or 11, **64-bit**
- No administrator rights and no internet connection needed
- The app is unsigned, so on first run Windows may warn you:
  **More info → Run anyway**

## Database types

Clicking **New database** lets you choose the shape of the data:

| Type | Columns | Duplicates matched on |
|---|---|---|
| **Contacts** | # · Full Name · Phone · ID · Source | phone number (last 9 digits) |
| **Products** | # · Name · Code/SKU · Price · Quantity · Source | code (text) |
| **Students** | # · Full Name · Group · Phone · ID · Source | phone number |
| **My own columns** | you define them | you choose |

Under **My own columns** you type one column name per line, then pick which
column identifies a duplicate and how it should be compared — as a phone
number, as text, or as a number. Whether to include a row-number column and a
source column is up to you as well.

### The schema travels with the file

- A `.xlsx` saved by this app stores its column definition on a **hidden
  sheet**, so reopening the file restores the columns exactly and the app
  switches to that type by itself.
- Open **someone else's Excel file** and the columns are taken from its
  **header row** instead. Any spreadsheet can become a database.

## Getting started

Double-click `run.bat`. On the first run it installs the dependencies
(`openpyxl`, `sv-ttk`, `pillow`, `tkinterdnd2`).

To produce a standalone binary, run `build.bat`; the result lands in
`dist\Database Manager.exe` and needs no Python on the target machine.

## How it works

1. **New database** starts an empty list, numbering from **1**.
   **Open database** loads an existing Excel file and continues numbering
   **from the last row**. **Recent files** reopens something you had before.

2. **Add contacts** takes one or more files — or you can drag them straight
   onto the window. For each file it asks *"Whose number did these contacts
   come from?"*. **The field starts empty and you fill it in**; the app never
   guesses. Whatever you type goes into the **Source** column for every row
   that came from that file.

3. **Duplicates** — when the same phone number appears more than once, both
   the old and the new row turn **red**. If a file contains duplicates, you
   are asked before they are added: take everything, or skip the repeats.

4. **Save** writes an Excel file. The red marks are preserved there too.

## Features

| What | How |
|---|---|
| Search | One box searches name, number, ID and source at once |
| Sort by column | Click a header: ascending → descending → original order |
| Show duplicates only | The **Duplicates only** switch |
| Edit a whole row | Double-click it (or press `Enter`) |
| Fix the source of several rows | Select them, then **Change source** |
| Add a row by hand | **Add manually** in the sidebar |
| Delete rows | Select them, then **Delete** or the `Delete` key |
| Undo a mistake | `Ctrl+Z` — deletions, edits and imports all come back |
| Copy to Excel | Select rows and press `Ctrl+C` |
| Paste from the clipboard | `Ctrl+V` — rows copied out of Excel or a text list |
| Dark mode | The **Dark mode** button at the bottom of the sidebar |
| Shortcuts | `Ctrl+N` new · `Ctrl+O` open · `Ctrl+S` save · `Ctrl+F` search · `Ctrl+A` select all · `Ctrl+Y` redo |

## Tools

The **Tools** button in the sidebar opens:

| What | What it does |
|---|---|
| **Paste from clipboard** | Copy rows out of Excel and drop them straight into the database — no file needed |
| **Normalize phone format** | Brings every number into one shape: `901112233`, `998901112233`, `+998 90-111-22-33` → **`+998 90 111 22 33`**. Foreign numbers only lose their punctuation, very short ones are left alone. `Ctrl+Z` undoes it |
| **Suspicious rows only** | Shows rows with no phone, a too-short number, letters mixed into the number, or no name. They are marked **yellow** in the table |
| **Export to CSV** | A `.csv` Excel opens correctly (semicolon separated, UTF-8) |
| **Export to vCard** | A `.vcf` for loading contacts back onto a phone |

## Reusing the source value

When several files come from the same person you do not have to retype it:

- Previously entered values sit in a **dropdown** (the last 20 are remembered)
- Pick several files at once and tick **"Use this number for the remaining
  files"** — you will not be asked again

Four counters are always visible: total rows, how many are duplicates, how
many distinct sources were merged, and how many rows are currently shown.

The language, theme, the "last 9 digits" setting, window size, **column
widths**, recent files and recently used source values are all remembered
between runs.

When the database is empty the recent files appear in the middle of the
screen as a clickable list.

## Keeping your data safe

- Saving writes to a **temporary file** first and only swaps it into place
  once it is complete. If anything fails mid-write, the old database is
  untouched.
- Before overwriting, a **`.bak` copy** of the previous file is made
  (`base.xlsx` → `base.xlsx.bak`).
- Unexpected errors are shown on screen and appended to
  `%APPDATA%\DatabaseManager\errors.log`. Settings live in the same folder.

## Windows 10 and Windows 11

Both are supported. What was checked:

| | Windows 11 | Windows 10 |
|---|---|---|
| Icons (28 of them) | `SegoeIcons.ttf` | `segmdl2.ttf` — all present, verified |
| Font | Segoe UI Variable | falls back to Segoe UI |
| Dark title bar | ✅ | ✅ (including 1809–1909) |
| Tinted title bar | ✅ | plain system title bar (a Windows 11 feature) |
| Display scaling 100–200% | ✅ | ✅ |

Window dimensions are computed **from the measured text width**, not from
fixed pixel values. That way a 150% display scaling and the longer Russian
labels both still fit — verified at 100%, 125%, 150%, 175% and 200%.

The `.exe` is built by GitHub Actions on Windows Server 2022 (the Windows 10
kernel), so the same binary runs on both systems.

## Matching phone numbers

By default the **last 9 digits** are compared, so `+998901112233`,
`998901112233` and `901112233` all count as the same number. The
**Match last 9 digits** switch turns that off.

## Supported input files

Contact lists do not have to be Excel:

| Type | Example |
|---|---|
| `.xlsx` `.xlsm` | Excel spreadsheet |
| `.csv` `.tsv` | comma or tab separated |
| `.txt` | a plain text list — every shape below is understood |
| `.vcf` `.vcard` | vCard exported from a phone |
| anything else | read as text |

Inside a `.txt` all of these work:

```
1. Alisher Karimov — +998 90 111 22 33     ← list number, dash, spaces
Dilnoza Yusupova - 998902223344            ← name and number
+998901112233                              ← number only
Alisher Karimov|+998901112233|u1001        ← split by | ; , or tab
```

The encoding is detected automatically (UTF-8, UTF-16, Windows-1251, so
Cyrillic is handled).

Columns are located in two ways: first by header name (`name`, `phone`,
`number`, `id`, `source`, and their Russian and Uzbek equivalents), and if
there is no header, **by content** — which column holds phone-like values,
which holds names made of letters, which is a plain row number.

- Writes: `.xlsx`
- The old `.xls` format is not supported — open it in Excel and save as
  `.xlsx`.

## Sample files

`samples/` contains files to try it out with:

- `existing_database.xlsx` — an existing database of 14 contacts
- `contacts_1.xlsx` — 3 contacts (1 already in the database)
- `contacts_2_no_header.xlsx` — Excel with no header row, 3 contacts
- `contacts_3.txt` — plain text, 4 contacts
- `contacts_4.vcf` — vCard, 3 contacts

Try it: open the existing database, add the contact files one by one giving
each its own source number, and watch the repeated numbers turn red.

## Repository layout

| File | Purpose |
|---|---|
| `database_manager.py` | the application itself |
| `run.bat` | starts it |
| `build.bat` | builds a single `.exe` into `dist\` |
| `icon.ico` | application icon (`build.bat` can regenerate it) |
| `samples/` | files to test with |
| `tests/smoke_test.py` | headless tests (CI runs these) |
| `.github/workflows/build.yml` | builds the `.exe` on a tag and publishes a Release |
| `requirements.txt` | dependencies |
| `LICENSE` | MIT |

## Development

```bash
git clone https://github.com/nasafuriy/database-manager.git
cd database-manager
python -m pip install -r requirements.txt
python database_manager.py
```

Tests:

```bash
python tests/smoke_test.py
```

Cutting a release:

```bash
git tag v1.0.0
git push origin v1.0.0
```

Pushing the tag makes GitHub Actions run the tests, build the `.exe`, check
that it starts, and attach it to a Release.

## License

MIT — see [LICENSE](LICENSE).
