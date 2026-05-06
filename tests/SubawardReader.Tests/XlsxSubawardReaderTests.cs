using SubawardReader;
using Xunit;

namespace SubawardReader.Tests;

public sealed class XlsxSubawardReaderTests
{
    [Fact]
    public void ExampleOneContainsExpectedSubrecipients()
    {
        var workbookPath = FindRepositoryFile("SubawardBudgetExample1.xlsx");
        var reader = new XlsxSubawardReader();

        var result = reader.Read(workbookPath);

        var names = result.Subawards.Select(entry => entry.RecipientName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Equal(4, result.Subawards.Count);
        Assert.Contains("Indiana", names);
        Assert.Contains("Mayo", names);
        Assert.Contains("Purdue", names);
        Assert.Contains("Florida", names);
    }

    private static string FindRepositoryFile(string fileName)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, fileName);
            if (File.Exists(candidate))
            {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Could not find {fileName} from {AppContext.BaseDirectory}");
    }
}
