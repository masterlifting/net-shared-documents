using System.Data;
using Net.Shared.Documents.Interfaces;

namespace Net.Shared.Documents.Excel;

public sealed class ExcelDocument : IExcelDocument
{
    public int RowsCount { get; }
    private readonly DataTable _table;
    public ExcelDocument(DataTable table)
    {
        _table = table;
        RowsCount = table.Rows.Count;
    }

    public bool TryGetCell(int rowId, int columnId, out string? value)
    {
        value = _table.Rows[rowId].ItemArray[columnId]?.ToString();
        return !string.IsNullOrWhiteSpace(value);
    }
    public bool TryGetCell(int rowId, int columnId, string pattern, out string? value) =>
        TryGetCell(rowId, columnId, out value)
        && value!.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) > -1;
    public bool TryGetCell(int rowId, int columnId, IEnumerable<string> patterns, out string? value)
    {
        if (!TryGetCell(rowId, columnId, out value))
            return false;
        else
        {
            string _value = value!;
            return patterns.Any(x => _value.IndexOf(x, StringComparison.OrdinalIgnoreCase) > -1);
        }
    }
}