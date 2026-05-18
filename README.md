# Tilbrook Consulting Customer Brief Generator

This repository contains the first version of a Google Apps Script workflow for a Google Sheets-driven customer brief generator.

## What it does

- Reads prospect/customer rows from the active Google Sheet.
- Processes only rows where **Status** is exactly `Ready`.
- Generates a one-page **Growth Opportunity Snapshot** Google Doc.
- Writes the Google Doc link back to the sheet.
- Updates **Status** to `Brief Created` when successful.
- Writes **Status** as `Error` and records **Error Notes** when a row cannot be processed.
- Optionally exports a PDF and creates a Gmail draft when script properties are enabled.
- Never sends email automatically.
- Does not overwrite an existing brief link unless the explicit replace menu item is used.

## Sheet columns

The script expects these columns in order, with headers in row 1 and data starting in row 2:

| Column | Field |
| --- | --- |
| A | Business Name |
| B | Owner / Contact Name |
| C | Email |
| D | Industry |
| E | Location |
| F | Website |
| G | Business Notes |
| H | Growth Challenges |
| I | Offer Angle |
| J | Status |
| K | Google Doc Link |
| L | PDF Link |
| M | Gmail Draft Status |
| N | Last Updated |
| O | Error Notes |

## Apps Script setup

1. Create or open the Google Sheet using the schema above.
2. Open **Extensions → Apps Script**.
3. Add `Code.js` and `appsscript.json` from this repository.
4. Reload the spreadsheet.
5. Use **Tilbrook Briefs → Process ready briefs** to process rows marked `Ready`.

## Optional script properties

In Apps Script, open **Project Settings → Script Properties** and add any of these values:

| Property | Value | Purpose |
| --- | --- | --- |
| `EXPORT_PDF` | `true` | Export a PDF copy and write its link to column L. |
| `CREATE_GMAIL_DRAFTS` | `true` | Create a Gmail draft and update column M. Never sends email. |
| `BRIEF_FOLDER_ID` | Google Drive folder ID | Move generated Docs and optional PDFs into a target folder. |

By default, PDFs and Gmail drafts are disabled.
