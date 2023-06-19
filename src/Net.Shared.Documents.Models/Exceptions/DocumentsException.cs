using Net.Shared.Exceptions;

namespace Net.Shared.Documents.Models.Exceptions;

public sealed class DocumentsException : NetSharedException
{
    public DocumentsException(string message) : base(message)
    {
    }

    public DocumentsException(Exception exception) : base(exception)
    {
    }
}