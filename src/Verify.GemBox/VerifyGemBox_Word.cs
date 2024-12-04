using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using GemBox.Document;
using LoadOptions = GemBox.Document.LoadOptions;
using SaveOptions = GemBox.Document.SaveOptions;

namespace VerifyTests;

public static partial class VerifyGemBox
{
    static ConversionResult ConvertDocx(Stream stream, IReadOnlyDictionary<string, object> settings) =>
        Convert(stream, settings, LoadOptions.DocxDefault);

    static ConversionResult ConvertDoc(Stream stream, IReadOnlyDictionary<string, object> settings) =>
        Convert(stream, settings, LoadOptions.DocDefault);

    static ConversionResult Convert(Stream stream, IReadOnlyDictionary<string, object> settings, LoadOptions loadOptions)
    {
        var document = DocumentModel.Load(stream, loadOptions);
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(DocumentModel document, IReadOnlyDictionary<string, object> settings) =>
        new(GetInfo(document), GetWordStreams(document).ToList());

    static object GetInfo(DocumentModel document) =>
        new
        {
            Author = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Author, out var author) ? author : null,
            Title = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Title, out var title) ? title : null,
            Subject = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Subject, out var subject) ? subject : null,
            Comments = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Comments, out var comments) ? comments : null,
            Category = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Category, out var category) ? category : null,
            Status = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Status, out var status) ? status : null,
            Keywords = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Keywords, out var keywords) ? keywords : null,
            LastSavedBy = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.LastSavedBy, out var lastSavedBy) ? lastSavedBy : null,
            Manager = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Manager, out var manager) ? manager : null,
            Company = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Company, out var company) ? company : null,
            HyperlinkBase = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.HyperlinkBase, out var hyperlinkBase) ? hyperlinkBase : null,
            Application = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.Application, out var application) ? application : null,
            DateContentCreated = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.DateContentCreated, out var dateContentCreated) ? dateContentCreated : null,
            DateLastSaved = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.DateLastSaved, out var dateLastSaved) ? dateLastSaved : null,
            DateLastPrinted = document.DocumentProperties.BuiltIn.TryGetValue(BuiltInDocumentProperty.DateLastPrinted, out var dateLastPrinted) ? dateLastPrinted : null
        };

    static IEnumerable<Target> GetWordStreams(DocumentModel document)
    {
        foreach (var page in document.GetPaginator().Pages)
        {
            var textStream = new MemoryStream();
            page.Save(textStream, SaveOptions.TxtDefault);

            yield return new("txt", ReadLines(textStream));

            var imageStream = new MemoryStream();
            page.Save(imageStream, SaveOptions.ImageDefault);

            imageStream.Position = 0;

            yield return new("png", imageStream);
        }
    }

    static string ReadLines(MemoryStream stream)
    {
        stream.Position = 0;
        var builder = new StringBuilder();
        using var writer = new StringWriter(builder);
        using var reader = new StreamReader(stream);
        while (reader.ReadLine() is { } line)
        {
            writer.WriteLine(line);
        }

        return builder.ToString();
    }
}