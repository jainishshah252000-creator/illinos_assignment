namespace SubawardReader;

public sealed record SubawardEntry(string RecipientName, decimal Amount);

public sealed record BudgetFileResult(string FileName, IReadOnlyList<SubawardEntry> Subawards);
