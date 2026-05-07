using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace SubawardReader;

public sealed class XlsxSubawardReader
{
    private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public BudgetFileResult Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        // An xlsx file is really a zip file with XML files inside it.
        // I read those XML files directly so the console app does not need an Excel package.
        using var archive = ZipFile.OpenRead(path);
        var sharedStrings = ReadSharedStrings(archive);
        var subawards = new List<SubawardEntry>();

        foreach (var worksheet in archive.Entries
            .Where(entry => entry.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            subawards.AddRange(ReadWorksheet(worksheet, sharedStrings));
        }

        return new BudgetFileResult(Path.GetFileName(path), subawards);
    }

    private static IReadOnlyList<SubawardEntry> ReadWorksheet(ZipArchiveEntry worksheet, IReadOnlyList<string> sharedStrings)
    {
        using var stream = worksheet.Open();
        var document = XDocument.Load(stream);
        var rows = document.Descendants(SpreadsheetNamespace + "row")
            .Select(row => ReadRow(row, sharedStrings))
            .Where(row => row.Cells.Count > 0)
            .OrderBy(row => row.RowNumber)
            .ToArray();

        var totalColumn = FindTotalColumn(rows);
        var section = FindOtherDirectCostsSection(rows);

        // The prompt says the subaward rows are under "G. Other Direct Costs".
        // If that section is not found, I still scan the sheet so the app gives a useful result.
        var scanRows = section is null
            ? rows
            : rows.Where(row => row.RowNumber > section.Value.StartRow && row.RowNumber < section.Value.EndRow).ToArray();

        var results = new List<SubawardEntry>();

        foreach (var row in scanRows)
        {
            var cells = row.Cells.OrderBy(cell => cell.ColumnNumber).ToArray();

            for (var i = 0; i < cells.Length; i++)
            {
                var label = cells[i].Text.Trim();
                if (!label.StartsWith("Subaward:", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var recipientName = ExtractRecipientName(label, cells, i);
                if (string.IsNullOrWhiteSpace(recipientName))
                {
                    continue;
                }

                var amount = ReadAmount(row, totalColumn);
                results.Add(new SubawardEntry(recipientName, amount));
                break;
            }
        }

        return results;
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return Array.Empty<string>();
        }

        using var stream = entry.Open();
        var document = XDocument.Load(stream);

        return document.Root?
            .Elements(SpreadsheetNamespace + "si")
            .Select(item => string.Concat(item.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value)))
            .ToArray() ?? Array.Empty<string>();
    }

    private static WorksheetRow ReadRow(XElement row, IReadOnlyList<string> sharedStrings)
    {
        var rowNumber = int.Parse(row.Attribute("r")?.Value ?? "0", CultureInfo.InvariantCulture);
        var cells = row.Elements(SpreadsheetNamespace + "c")
            .Select(cell => ReadCell(cell, sharedStrings))
            .Where(cell => cell is not null)
            .Cast<WorksheetCell>()
            .ToArray();

        return new WorksheetRow(rowNumber, cells);
    }

    private static WorksheetCell? ReadCell(XElement cell, IReadOnlyList<string> sharedStrings)
    {
        var reference = cell.Attribute("r")?.Value;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return null;
        }

        var columnNumber = GetColumnNumber(reference);
        var value = cell.Element(SpreadsheetNamespace + "v")?.Value ?? string.Empty;
        var type = cell.Attribute("t")?.Value;

        var text = type switch
        {
            "s" when int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                && index >= 0
                && index < sharedStrings.Count => sharedStrings[index],
            "inlineStr" => string.Concat(cell.Descendants(SpreadsheetNamespace + "t").Select(textNode => textNode.Value)),
            _ => value
        };

        return new WorksheetCell(columnNumber, text);
    }

    private static int? FindTotalColumn(IReadOnlyList<WorksheetRow> rows)
    {
        return rows
            .Where(row => row.RowNumber <= 20)
            .SelectMany(row => row.Cells)
            .Where(cell => string.Equals(cell.Text.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
            .OrderBy(cell => cell.ColumnNumber)
            .Select(cell => (int?)cell.ColumnNumber)
            .LastOrDefault();
    }

    private static (int StartRow, int EndRow)? FindOtherDirectCostsSection(IReadOnlyList<WorksheetRow> rows)
    {
        var startRow = rows.FirstOrDefault(row => row.Cells.Any(cell => string.Equals(cell.Text.Trim(), "G.", StringComparison.OrdinalIgnoreCase)))?.RowNumber;
        if (startRow is null)
        {
            return null;
        }

        var endRow = rows
            .Where(row => row.RowNumber > startRow.Value)
            .FirstOrDefault(row => row.Cells.Any(cell => string.Equals(cell.Text.Trim(), "H.", StringComparison.OrdinalIgnoreCase)))?
            .RowNumber ?? int.MaxValue;

        return (startRow.Value, endRow);
    }

    private static string ExtractRecipientName(string label, IReadOnlyList<WorksheetCell> cells, int labelIndex)
    {
        var inlineName = label[("Subaward:").Length..].Trim();
        if (!string.IsNullOrWhiteSpace(inlineName))
        {
            return inlineName;
        }

        return cells
            .Skip(labelIndex + 1)
            .Select(cell => cell.Text.Trim())
            .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text) && !TryParseDecimal(text, out _)) ?? string.Empty;
    }

    private static decimal ReadAmount(WorksheetRow row, int? totalColumn)
    {
        if (totalColumn is not null)
        {
            var totalCell = row.Cells.FirstOrDefault(cell => cell.ColumnNumber == totalColumn.Value);
            if (totalCell is not null && TryParseDecimal(totalCell.Text, out var amount))
            {
                return amount;
            }
        }

        return row.Cells
            .Where(cell => TryParseDecimal(cell.Text, out _))
            .Select(cell =>
            {
                TryParseDecimal(cell.Text, out var amount);
                return amount;
            })
            .LastOrDefault();
    }

    private static bool TryParseDecimal(string value, out decimal amount)
    {
        return decimal.TryParse(value, NumberStyles.Number | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out amount);
    }

    private static int GetColumnNumber(string cellReference)
    {
        var columnNumber = 0;

        foreach (var character in cellReference)
        {
            if (!char.IsAsciiLetterUpper(character) && !char.IsAsciiLetterLower(character))
            {
                break;
            }

            columnNumber = (columnNumber * 26) + (char.ToUpperInvariant(character) - 'A' + 1);
        }

        return columnNumber;
    }

    private sealed record WorksheetRow(int RowNumber, IReadOnlyList<WorksheetCell> Cells);

    private sealed record WorksheetCell(int ColumnNumber, string Text);
}
