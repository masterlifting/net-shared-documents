using ExcelDataReader;
using System.Text;
using Net.Shared.Documents.Abstractions.Excel;

namespace Net.Shared.Documents.Excel;

public sealed class ExcelDataReaderService : IExcelDocumentService
{
    public IExcelDocument Load(byte[] document)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        using var stream = new MemoryStream(document);
        using var reader = ExcelReaderFactory.CreateBinaryReader(stream);
        
        var dataSet = reader.AsDataSet();
        var table = dataSet.Tables[0];
        
        stream.Close();

        return new ExcelDataReaderDocument(table);
    }
}