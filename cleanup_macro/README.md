# 3 macros executed by 1 custom function to clean up data in Google Sheets

- Follow me on
  [Twitter](https://twitter.com/TechandEco)
- Visit this solution's step-by-step
  [instructions with screenshots here](https://medium.com/@TechandEco/cleanup-exported-data-automatically-with-apps-script-and-advanced-formulas-in-google-sheets-4f77cd8361ca)

## Introduction

- This is 1 of 5 Apps Script tutorials for users in Finance.

- Sample covers how to automatically clean-up .XLS exports once converted to
  Google Sheets using 3 macros within a
  [custom menu](https://developers.google.com/apps-script/guides/menus)
  called **Clean my export**.

- Helps users learn how to write automation scripts by recording logic via
  [macrorecording](https://support.google.com/docs/answer/7665004?co=GENIE.Platform%3DDesktop&hl=en)
  in Google Sheets.

## Try it

1. [Copy this Google Sheet](https://docs.google.com/spreadsheets/d/1ew8WY7UkAdA8QDCh9LLUXba_OB0czhH0dkGWvS1Pnb4/copy)

1. Locate menu **Clean my export** > **Find misspelled statesm** option.
   This creates a _pivot table_ to detect mispelled states in a different tab.

1. Locate menu **Clean my export** > **Lookup office codes** option.
   This performs a `VLOOKUP` of office codes from another sheet.

1. Locate menu **Clean my export** > **Query 2019 records** option.
   This performs a `QUERY` (SQL) table that has an unfriendly date format
   ex: `'201901__'`.

1. Combine `VLOOKUP` and `QUERY`actions into one by locating menu
   **Clean my export** > **Find and query** option.
