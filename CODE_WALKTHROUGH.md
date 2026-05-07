# Code Walkthrough

This document explains the implementation file by file. Line numbers refer to the current source files in this repository.

## `src/SubawardReader/Models.cs`

| Lines | Explanation |
| --- | --- |
| 1 | Declares the `SubawardReader` namespace so the model types share the same namespace as the console app and parser. |
| 3 | Defines `SubawardEntry`, an immutable record for one parsed subaward row. It stores the recipient name and the amount found for that row. |
| 5 | Defines `BudgetFileResult`, an immutable record for one workbook result. It stores the file name and the list of subawards found in that file. |

## `src/SubawardReader/Program.cs`

| Lines | Explanation |
| --- | --- |
| 1 | Imports `System.Globalization` so currency values can be formatted using the current culture. |
| 3 | Declares the `SubawardReader` namespace. |
| 5 | Declares `Program` as a static class because it only contains the console entry point. |
| 7 | Defines `Main`, the method that runs when the console app starts. It returns an exit code. |
| 9 | Uses the first command-line argument as the input folder. If no argument is provided, it uses the current working directory. |
| 11-15 | Checks whether the folder exists. If not, it writes an error and exits with code `1`. |
| 17-21 | Finds all `.xlsx` files directly in the folder, skips temporary Excel lock files that start with `~$`, sorts them by file name, and materializes the list. |
| 23-27 | Handles the empty-folder case with a friendly message and exits successfully. |
| 29 | Creates the reusable Excel parser. |
| 30 | Reads every workbook and stores the parsed results. |
| 32-49 | Prints each file name, then prints that file's subawards. If a file has no subawards, it prints a clear message. |
| 44 | Formats each subaward amount as whole-dollar currency, for example `$25,000`. |
| 51-52 | Prints the heading for the cross-file totals section. |
| 54-55 | Flattens all per-file subaward lists into one sequence. |
| 56 | Groups recipients case-insensitively so `Mayo` and `mayo` would be totaled together. |
| 57-61 | Builds one total row per recipient, preserving the first observed display name and summing amounts. |
| 62 | Sorts the final totals alphabetically by recipient name. |
| 64 | Prints each recipient and total amount as currency. |
| 67 | Returns exit code `0`, meaning the app completed successfully. |

## `src/SubawardReader/XlsxSubawardReader.cs`

| Lines | Explanation |
| --- | --- |
| 1 | Imports culture APIs for invariant numeric parsing. |
| 2 | Imports ZIP APIs because `.xlsx` files are ZIP archives containing XML files. |
| 3 | Imports regular expression APIs for identifying numeric-looking strings. |
| 4 | Imports LINQ-to-XML APIs for reading worksheet XML. |
| 6 | Declares the `SubawardReader` namespace. |
| 8 | Declares the parser class. `partial` is used because the generated regex method also requires a partial type. |
| 10 | Stores the spreadsheet XML namespace used by Open XML workbook files. |
| 12 | Starts the public `Read` method, which parses one workbook path. |
| 14 | Validates that the caller provided a non-empty path. |
| 16 | Opens the `.xlsx` file as a read-only ZIP archive. |
| 17 | Reads the workbook shared string table. Excel often stores text once in `sharedStrings.xml` and references it by index from cells. |
| 18 | Creates the list that will collect all subawards found across worksheets. |
| 20-23 | Selects worksheet XML files from inside the workbook archive and processes them in deterministic name order. |
| 25 | Reads subaward rows from one worksheet and appends them to the workbook-level list. |
| 28 | Returns a result containing the workbook file name and parsed subawards. |
| 31 | Starts `ReadWorksheet`, which parses one worksheet XML file. |
| 33 | Opens the worksheet XML stream from inside the archive. |
| 34 | Loads the worksheet XML into an `XDocument`. |
| 35-39 | Reads every non-empty row into a simpler internal `WorksheetRow` structure, ordered by Excel row number. |
| 41 | Finds the column that contains the worksheet's `Total` heading. That column is used for subaward amounts. |
| 42 | Finds the row range for `G. Other Direct Costs`. |
| 43-45 | Limits scanning to that section when the section markers are present; otherwise, it scans the full sheet as a fallback. |
| 47 | Creates the result list for this worksheet. |
| 49 | Loops through each candidate row. |
| 51 | Orders cells by column so left-to-right parsing is predictable. |
| 53 | Loops through each cell in the row looking for a subaward label. |
| 55 | Trims the current cell text for comparison. |
| 56-59 | Skips cells that do not start with `Subaward:`. |
| 61 | Extracts the recipient name from either the label cell or the next populated text cell. |
| 62-65 | Ignores placeholder subaward rows that do not contain a recipient name. |
| 67 | Reads the subaward amount from the row, preferring the `Total` column. |
| 68 | Adds the parsed recipient and amount to the results. |
| 69 | Stops scanning the current row after the subaward entry has been captured. |
| 73 | Returns all subawards found in this worksheet. |
| 76 | Starts `ReadSharedStrings`, which loads Excel's shared text table. |
| 78 | Looks for `xl/sharedStrings.xml` inside the workbook archive. |
| 79-82 | Returns an empty list if the workbook does not have a shared string table. |
| 84-85 | Opens and parses the shared string XML file. |
| 87-90 | Converts each shared string item into plain text by concatenating all text nodes. This supports rich text strings split across multiple XML runs. |
| 93 | Starts `ReadRow`, which converts one XML row into an internal row object. |
| 95 | Reads the Excel row number from the row's `r` attribute. |
| 96-100 | Converts each XML cell into an internal cell object and removes cells that could not be read. |
| 102 | Returns the simplified row object. |
| 105 | Starts `ReadCell`, which converts one XML cell into an internal cell object. |
| 107 | Reads the Excel cell reference, for example `B64`. |
| 108-111 | Skips cells that have no reference. |
| 113 | Converts the cell reference letters into a numeric column number. |
| 114 | Reads the raw `<v>` value, or an empty string if the cell has no value. |
| 115 | Reads the cell type attribute. |
| 117-124 | Converts the raw value to display text. Shared-string cells are resolved through the shared string table, inline strings are read from text nodes, and other values are used directly. |
| 126 | Returns the simplified cell object. |
| 129 | Starts `FindTotalColumn`, which identifies the amount column. |
| 131-137 | Searches the first 20 rows for cells that equal `Total`, orders matching columns, and uses the rightmost matching column. This handles worksheets with multiple period columns. |
| 140 | Starts `FindOtherDirectCostsSection`, which identifies the budget section to scan. |
| 142 | Finds the row containing `G.`. |
| 143-146 | Returns `null` if the section marker is missing, allowing the caller to scan the full sheet. |
| 148-151 | Finds the next later row containing `H.` and treats it as the end of the `G.` section. If missing, the section runs to the end of the sheet. |
| 153 | Returns the start and end row numbers for the section. |
| 156 | Starts `ExtractRecipientName`, which handles the two observed subaward label formats. |
| 158 | Reads any recipient name that appears in the same cell after `Subaward:`. |
| 159-162 | Returns the inline name when it exists, such as `Subaward: Mayo`. |
| 164-167 | Otherwise, scans cells to the right and returns the first non-empty, non-numeric value. This handles `Subaward:` in one cell and `Indiana` in the next cell. |
| 170 | Starts `ReadAmount`, which finds the amount for one subaward row. |
| 172-179 | If a `Total` column was found, it reads and returns that cell's numeric value. |
| 181-188 | If no `Total` column is available, it falls back to the last numeric value on the row. |
| 191-194 | Parses decimal values using invariant culture and allows scientific notation, which appears in some Excel XML values. |
| 196 | Starts `GetColumnNumber`, which converts Excel column letters to numbers. |
| 198 | Initializes the running column number. |
| 200 | Loops through the letters at the start of the cell reference. |
| 202-205 | Stops when the reference reaches row digits, for example the `64` in `B64`. |
| 207 | Converts letters using base-26 math, so `A` is 1, `B` is 2, and `AA` is 27. |
| 210 | Returns the numeric column. |
| 213-214 | Defines a generated regex for recognizing plain numeric strings. It is used to avoid treating amount cells as recipient names. |
| 216 | Defines the internal row structure used by the parser. |
| 218 | Defines the internal cell structure used by the parser. |

## `tests/SubawardReader.Tests/XlsxSubawardReaderTests.cs`

| Lines | Explanation |
| --- | --- |
| 1 | Imports the application namespace so the test can instantiate `XlsxSubawardReader`. |
| 2 | Imports xUnit test APIs. |
| 4 | Declares the test namespace. |
| 6 | Declares the test class. |
| 8 | Marks the next method as an xUnit test. |
| 9 | Names the behavior under test: example workbook 1 should contain the expected subrecipients. |
| 11 | Locates `SubawardBudgetExample1.xlsx` from the repository tree. |
| 12 | Creates the Excel parser. |
| 14 | Parses the workbook. |
| 16 | Converts parsed names into a case-insensitive set for easy assertions. |
| 17 | Confirms the workbook has exactly four parsed subaward rows. |
| 18-21 | Confirms the four required subrecipients are present: Indiana, Mayo, Purdue, and Florida. |
| 24 | Starts a helper method for finding a repository file from the test output folder. |
| 26 | Begins searching from the test process base directory. |
| 28-37 | Walks upward through parent directories until the requested file is found. |
| 30 | Builds a possible file path in the current directory. |
| 31-34 | Returns the path as soon as it exists. |
| 36 | Moves up one directory level and continues searching. |
| 39 | Throws a clear error if the file cannot be found. |

## Project Files

| File | Explanation |
| --- | --- |
| `src/SubawardReader/SubawardReader.csproj` | Defines the console application project, targets .NET 9, enables nullable reference checking, and enables implicit `using` directives. |
| `tests/SubawardReader.Tests/SubawardReader.Tests.csproj` | Defines the xUnit test project, targets .NET 9, references the app project, and includes test runner packages. |
| `SubawardReader.sln` | Groups the app and test projects into one solution so reviewers can run `dotnet test` from the repository root. |
| `README.md` | Explains how to run the app, how to run tests, assumptions, questions, and the written interview answers. |
