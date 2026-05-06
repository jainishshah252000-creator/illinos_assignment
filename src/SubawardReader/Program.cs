using System.Globalization;

namespace SubawardReader;

public static class Program
{
    public static int Main(string[] args)
    {
        var folder = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

        if (!Directory.Exists(folder))
        {
            Console.Error.WriteLine($"Folder not found: {folder}");
            return 1;
        }

        var files = Directory
            .EnumerateFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (files.Length == 0)
        {
            Console.WriteLine($"No .xlsx files found in {folder}");
            return 0;
        }

        var reader = new XlsxSubawardReader();
        var results = files.Select(reader.Read).ToArray();

        foreach (var result in results)
        {
            Console.WriteLine(result.FileName);

            if (result.Subawards.Count == 0)
            {
                Console.WriteLine("  No subawards found.");
            }
            else
            {
                foreach (var subaward in result.Subawards)
                {
                    Console.WriteLine($"  {subaward.RecipientName}: {subaward.Amount.ToString("C0", CultureInfo.CurrentCulture)}");
                }
            }

            Console.WriteLine();
        }

        Console.WriteLine("Subaward totals across all files");
        Console.WriteLine("--------------------------------");

        foreach (var total in results
            .SelectMany(result => result.Subawards)
            .GroupBy(entry => entry.RecipientName, StringComparer.OrdinalIgnoreCase)
            .Select(group => new
            {
                Name = group.First().RecipientName,
                Amount = group.Sum(entry => entry.Amount)
            })
            .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase))
        {
            Console.WriteLine($"{total.Name}: {total.Amount.ToString("C0", CultureInfo.CurrentCulture)}");
        }

        return 0;
    }
}
