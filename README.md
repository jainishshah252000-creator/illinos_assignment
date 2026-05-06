# Subaward Reader

Small .NET console application for the Sponsored Programs Administration programming exercise. It reads SPA-style Excel budget workbooks from a folder, reports subaward recipients under `G. Other Direct Costs`, and totals each distinct subrecipient across all files.

## Requirements

- .NET 9 SDK or .NET 10 SDK

## Run

From the repository root:

```bash
dotnet run --project src/SubawardReader -- .
```

The `.` argument tells the app to read the Excel files in the repository root. You can also pass another folder containing workbooks in the same format:

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

## Approach

- Reads every `.xlsx` file in the requested folder, skipping temporary Excel lock files that start with `~$`.
- Uses the Open XML files inside each workbook directly with built-in .NET ZIP and XML APIs, so the app has no runtime package dependency for Excel parsing.
- Looks for rows in the `G. Other Direct Costs` section whose label starts with `Subaward:`.
- Supports both observed recipient formats:
  - `Subaward: Mayo`
  - `Subaward:` with the recipient name in the next populated cell to the right.
- Uses the worksheet column headed `Total` as the subaward amount. If no `Total` column is found, it falls back to the last numeric value on that subaward row.
- Outputs per-file recipients and a distinct cross-file total grouped by recipient name.
- Keeps the console output simple so non-technical staff can read file-level results and overall totals without inspecting the spreadsheet formulas.

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

When I work with non-technical stakeholders on a data-driven application, I try to start with the decision they need to make rather than the fields or screens they think they need. For example, on a reporting workflow I would ask what question the report must answer, who uses it, what they do after reading it, and what examples represent correct and incorrect output.

To make sure I understand the need, I turn the conversation into concrete examples: sample inputs, expected outputs, edge cases, and a short acceptance checklist. If requirements change, I separate the underlying business rule from the implementation detail, confirm the impact in plain language, and update the acceptance examples before changing code. That keeps misunderstandings visible early, when they are cheaper to correct.

### Communicating technical issues

When explaining a technical limitation to non-technical users, I avoid leading with implementation details. I explain what is happening, how it affects their work, what options exist, and what I recommend.

For example, if an import process cannot reliably match organizations because the source data uses inconsistent names, I would say: "The file has enough variation that automatic matching will produce some wrong results. We can either require a unique organization ID, add a review step for uncertain matches, or accept a higher error rate. I recommend adding the review step first because it protects the data while we learn how common the issue is." That gives users a decision they can evaluate without needing to understand the matching algorithm.
