# Subaward Reader

This is my .NET console application for the Sponsored Programs Administration programming exercise.

The app reads the Excel budget files from a folder, finds the subawards listed under `G. Other Direct Costs`, prints the subrecipient names for each file, and then prints the total amount for each subrecipient across all files.

## Requirements

- .NET 9 SDK or .NET 10 SDK

## Run

From the repository root:

```bash
dotnet run --project src/SubawardReader -- .
```

The `.` at the end means "use the current folder". Since the three example spreadsheets are in the repository root, this command will read those files.

You can also pass another folder if the spreadsheets are somewhere else:

```bash
dotnet run --project src/SubawardReader -- /path/to/budget/files
```

Expected output with the three provided workbooks:

```text
SubawardBudgetExample1.xlsx
  Indiana: $25,000
  Mayo: $20,637
  Purdue: $25,000
  Florida: $25,000

SubawardBudgetExample2.xlsx
  Ecotek: $25,000
  Purdue: $20,000
  Mayo: $19,782

SubawardBudgetExample3.xlsx
  U WA: $25,000
  U CO: $25,000
  Mayo: $25,000

Subaward totals across all files
--------------------------------
Ecotek: $25,000
Florida: $25,000
Indiana: $25,000
Mayo: $65,419
Purdue: $45,000
U CO: $25,000
U WA: $25,000
```

## Test

```bash
dotnet test
```

The unit test confirms `SubawardBudgetExample1.xlsx` contains exactly 4 subrecipients: `Indiana`, `Mayo`, `Purdue`, and `Florida`.

## How I Approached It

- I read every `.xlsx` file in the selected folder.
- I skip temporary Excel files that start with `~$`.
- I read the `.xlsx` file as a zip file because Excel stores worksheet data as XML inside the workbook.
- I look for the `G. Other Direct Costs` section.
- Inside that section, I look for rows starting with `Subaward:`.
- The spreadsheets had two formats, so I handled both:
  - `Subaward: Mayo`
  - `Subaward:` with the recipient name in the next populated cell to the right.
- I use the row's `Total` column as the amount.
- At the end, I group the same recipient names together and add their amounts.

## Assumptions

- The workbook is an `.xlsx` file, not legacy `.xls`.
- Subaward rows are part of the `G. Other Direct Costs` section and start with `Subaward:`.
- Blank subaward rows are placeholders and should be ignored.
- The row's `Total` column represents the subaward amount to aggregate across files.
- Recipient names are matched case-insensitively when calculating cross-file totals.
- Workbooks are expected to contain cached formula results, which Excel normally stores in `.xlsx` files.

## Questions I Would Ask

- Should cost share amounts be included in subaward totals, or should the report include sponsor and cost share totals separately?
- Should placeholder rows like `Subaward:` with no recipient be reported as data quality issues?
- Are recipient names expected to be normalized across files, for example `University of Washington` vs `U WA`?
- Should the app recurse into subfolders, or only read files directly inside the selected folder?
- What output format would staff prefer long term: console text, CSV, Excel, or a small report file?

## Additional Interview Questions

### Working with non-technical stakeholders

When I work with non-technical stakeholders on a data-driven application, I try to start by understanding the actual decision or task they need help with. I would ask what question the report or application needs to answer, who will use it, and what they will do after they see the result.

To make sure I understood correctly, I would ask for sample files, expected output, and examples of cases that should not be included. If there was a misunderstanding or the requirement changed, I would confirm the new rule in plain language before changing the code. That helps keep the work connected to the user's real need instead of only matching a technical assumption.

### Communicating technical issues

When explaining a technical issue to non-technical users, I try to avoid starting with implementation details. I explain what is happening, how it affects their work, and what options we have.

For example, if an import process could not reliably match organizations because the names were inconsistent, I would explain that some records may match incorrectly unless we add a stronger identifier or a review step. I would recommend the review step first if the data is important, because it protects the users from incorrect results while still allowing the work to move forward.
