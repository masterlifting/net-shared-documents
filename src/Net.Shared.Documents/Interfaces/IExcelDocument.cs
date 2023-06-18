namespace Net.Shared.Documents.Interfaces
{
    public interface IExcelDocument
    {
        int RowsCount { get; }

        bool TryGetCell(int rowId, int columnId, IEnumerable<string> patterns, out string? value);
        bool TryGetCell(int rowId, int columnId, out string? value);
        bool TryGetCell(int rowId, int columnId, string pattern, out string? value);
    }
}