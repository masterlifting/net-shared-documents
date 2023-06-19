namespace Net.Shared.Documents.Abstractions.Excel;

public interface IExcelDocumentService
{
    IExcelDocument Load(byte[] document);
}