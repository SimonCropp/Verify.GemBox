using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SaveOptions = GemBox.Spreadsheet.SaveOptions;

namespace VerifyTests;

public static partial class VerifyGemBox
{
    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var workbook = ExcelFile.Load(stream);
        return ConvertExcel(workbook, settings);
    }

    static ConversionResult ConvertExcel(ExcelFile book, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetInfo(book);
        return new(info, GetExcelStreams(book).ToList());
    }

    static object GetInfo(ExcelFile book) => new
    {
        book.CodeName,
        book.Use1904DateSystem,
        Author = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Author, out var author) ? author : null,
        Title = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Title, out var title) ? title : null,
        Subject = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Subject, out var subject) ? subject : null,
        Comments = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Comments, out var comments) ? comments : null,
        Category = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Category, out var category) ? category : null,
        Status = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Status, out var status) ? status : null,
        Keywords = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Keywords, out var keywords) ? keywords : null,
        LastSavedBy = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.LastSavedBy, out var lastSavedBy) ? lastSavedBy : null,
        Manager = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Manager, out var manager) ? manager : null,
        Company = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Company, out var company) ? company : null,
        HyperlinkBase = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.HyperlinkBase, out var hyperlinkBase) ? hyperlinkBase : null,
        Application = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.Application, out var application) ? application : null,
        DateContentCreated = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.DateContentCreated, out var dateContentCreated) ? dateContentCreated : null,
        DateLastSaved = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.DateLastSaved, out var dateLastSaved) ? dateLastSaved : null,
        DateLastPrinted = book.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperties.DateLastPrinted, out var dateLastPrinted) ? dateLastPrinted : null,
        IsWindowProtection = book.ProtectionSettings.ProtectWindows,
        IsProtected = book.Protected,
        ActiveSheetIndex = book.Worksheets.ActiveWorksheet?.Index,
        StandardFont = book.Styles.SingleOrDefault(s => s.IsDefault).Font.Name,
        StandardFontSize = book.Styles.SingleOrDefault(s => s.IsDefault).Font.Size
    };

    static IEnumerable<Target> GetExcelStreams(ExcelFile book)
    {
        foreach (var sheet in book.Worksheets)
        {
            using var stream = new MemoryStream();

            var file = new ExcelFile();
            file.Worksheets.AddCopy(sheet.Name, sheet);
            file.Save(stream, SaveOptions.CsvDefault);

            yield return new("csv", ReadNonEmptyLines(stream));
        }
    }

    static string ReadNonEmptyLines(MemoryStream stream)
    {
        stream.Position = 0;
        var builder = new StringBuilder();
        using (var writer = new StringWriter(builder))
        using (var reader = new StreamReader(stream))
        {
            while (reader.ReadLine() is { } line)
            {
                if (!string.IsNullOrWhiteSpace(line) && line.Any(c => c != ','))
                {
                    writer.WriteLine(line);
                }
            }
        }

        return builder.ToString();
    }
}