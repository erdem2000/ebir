using ClosedXML.Excel;
using NPOI.SS.UserModel;

namespace ExcelBirlestirici;

public sealed class ExcelPasswordCancelledException : Exception
{
    public string FileName { get; }

    public ExcelPasswordCancelledException(string fileName)
        : base($"{fileName}: şifre girilmedi, dosya atlandı.")
    {
        FileName = fileName;
    }
}

internal static class ExcelFileHelper
{
    private static readonly byte[] OleCompoundHeader = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    public static bool IsEncrypted(string path)
    {
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[OleCompoundHeader.Length];
        return stream.Read(header) == header.Length && header.SequenceEqual(OleCompoundHeader);
    }
}

internal sealed class ExcelPasswordSession
{
    private readonly Func<string, bool, string?>? _requestPassword;
    private string? _cachedPassword;

    public ExcelPasswordSession(Func<string, bool, string?>? requestPassword)
    {
        _requestPassword = requestPassword;
    }

    public XLWorkbook Open(string path, string fileName)
    {
        if (!ExcelFileHelper.IsEncrypted(path))
            return new XLWorkbook(path);

        var wrongPassword = false;

        while (true)
        {
            string? password;

            if (_cachedPassword is not null)
            {
                password = _cachedPassword;
            }
            else
            {
                if (_requestPassword is null)
                    throw new ExcelPasswordCancelledException(fileName);

                password = _requestPassword(fileName, wrongPassword);
                if (string.IsNullOrEmpty(password))
                    throw new ExcelPasswordCancelledException(fileName);
            }

            try
            {
                var workbook = OpenEncryptedWorkbook(path, password);
                _cachedPassword = password;
                return workbook;
            }
            catch (Exception ex) when (IsWrongPasswordException(ex))
            {
                if (string.Equals(_cachedPassword, password, StringComparison.Ordinal))
                    _cachedPassword = null;

                wrongPassword = true;
            }
        }
    }

    private static XLWorkbook OpenEncryptedWorkbook(string path, string password)
    {
        using var npoiWorkbook = WorkbookFactory.Create(path, password);
        using var decryptedStream = new MemoryStream();
        npoiWorkbook.Write(decryptedStream);
        return new XLWorkbook(new MemoryStream(decryptedStream.ToArray()));
    }

    private static bool IsWrongPasswordException(Exception ex)
    {
        for (var current = ex; current is not null; current = current.InnerException)
        {
            var typeName = current.GetType().FullName ?? string.Empty;
            if (typeName.Contains("EncryptedDocument", StringComparison.OrdinalIgnoreCase))
                return true;

            var message = current.Message;
            if (message.Contains("password", StringComparison.OrdinalIgnoreCase)
                || message.Contains("şifre", StringComparison.OrdinalIgnoreCase)
                || message.Contains("decrypt", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
