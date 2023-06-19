using Net.Shared.Documents.Abstractions.Excel;

using System.Data;

namespace Net.Shared.Documents.Excel;

public sealed class ExcelDataReaderDocument : IExcelDocument
{
    private readonly DataTable _table;

    public ExcelDataReaderDocument(DataTable table)
    {
        _table = table;
        RowsCount = _table.Rows.Count;
    }
    public int RowsCount { get; }

    public bool TryGetCell(int rowId, int columnId, out string value)
    {
        value = _table.Rows[rowId].ItemArray[columnId]?.ToString() ?? string.Empty;
        return !string.IsNullOrEmpty(value);
    }
    public bool TryGetCell(int rowId, int columnId, string pattern, out string value) =>
        TryGetCell(rowId, columnId, out value)
        && value!.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) > -1;
    public bool TryGetCell(int rowId, int columnId, IEnumerable<string> patterns, out string value)
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
