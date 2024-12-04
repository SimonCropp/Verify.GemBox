using System.Collections.Generic;
using System.IO;
using System.Linq;
using GemBox.Presentation;
using BuiltInDocumentProperty = GemBox.Presentation.BuiltInDocumentProperty;
using LoadOptions = GemBox.Presentation.LoadOptions;
using SaveOptions = GemBox.Presentation.SaveOptions;

namespace VerifyTests;

public static partial class VerifyGemBox
{
    static ConversionResult ConvertPpt(Stream stream, IReadOnlyDictionary<string, object> settings) =>
        Convert(stream, settings, LoadOptions.Ppt);

    static ConversionResult ConvertPptx(Stream stream, IReadOnlyDictionary<string, object> settings) =>
        Convert(stream, settings, LoadOptions.Pptx);

    static ConversionResult Convert(Stream stream, IReadOnlyDictionary<string, object> settings, LoadOptions loadOptions)
    {
        var document = PresentationDocument.Load(stream, loadOptions);
        return ConvertPowerPoint(document, settings);
    }

    static ConversionResult ConvertPowerPoint(PresentationDocument document, IReadOnlyDictionary<string, object> settings) =>
        new(GetInfo(document), GetPowerPointStreams(document, settings).ToList());

    static object GetInfo(PresentationDocument document) =>
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

    static IEnumerable<Target> GetPowerPointStreams(PresentationDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.Slides.Count);

        for (var index = 0; index < pagesToInclude; index++)
        {
            var slide = document.Slides[index];
            var presentation = new PresentationDocument();
            presentation.Slides.AddClone(slide);

            var image = new MemoryStream();
            presentation.Save(image, SaveOptions.Image);
            yield return new("png", image);
        }
    }
}