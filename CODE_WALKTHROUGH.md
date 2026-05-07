# Code Walkthrough

I added this file to make the project easier to review and easier for me to explain in the follow-up interview.

## Main Files

### `Program.cs`

This is the starting point of the console app.

It does a few simple things:

- Gets the folder path from the command line.
- Finds all `.xlsx` files in that folder.
- Skips Excel temporary files that start with `~$`.
- Calls `XlsxSubawardReader` for each workbook.
- Prints each file's subawards.
- Groups all subawards by recipient name and prints the final totals.

The grouping is case-insensitive so names like `Mayo` and `mayo` would be treated as the same recipient.

### `XlsxSubawardReader.cs`

This file has most of the parsing logic.

I learned that an `.xlsx` file is actually a zip file with XML files inside it. Because of that, I used built-in .NET classes like `ZipFile` and `XDocument` instead of adding an Excel library.

The reader:

- Opens the workbook as a zip archive.
- Reads Excel shared strings so text cells can be understood.
- Reads each worksheet XML file.
- Finds the `Total` column.
- Finds the `G.` section, which is `G. Other Direct Costs`.
- Looks for rows starting with `Subaward:`.
- Gets the recipient name.
- Gets the amount from the `Total` column.

There were two subaward formats in the sample files:

```text
Subaward: Mayo
```

and:

```text
Subaward:
```

with the recipient name in the next cell. The code handles both.

### `Models.cs`

This file contains two small record types:

- `SubawardEntry`: one recipient and amount.
- `BudgetFileResult`: one file name and the subawards found in that file.

### `XlsxSubawardReaderTests.cs`

This is the unit test required by the prompt.

It reads `SubawardBudgetExample1.xlsx` and checks that there are exactly four subrecipients:

- Indiana
- Mayo
- Purdue
- Florida

## Calculation Rule

For each subaward row, I use the value in the `Total` column as the subaward amount.

Then the console app adds the amounts together by recipient name.

Example:

```text
Mayo = 20,637 + 19,782 + 25,000 = 65,419
Purdue = 25,000 + 20,000 = 45,000
```

## Why I Kept The Output Simple

The prompt mentioned that the output should be easy for non-technical staff to read. Because of that, I used plain console output with the file name, indented subawards, and then a separate total section.
