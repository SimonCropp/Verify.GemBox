using System.Collections.Generic;
using System.IO;
using System.Linq;
using ImageSaveFormat = GemBox.Pdf.ImageSaveFormat;
using ImageSaveOptions = GemBox.Pdf.ImageSaveOptions;

namespace VerifyTests;

public static partial class VerifyGemBox
{
    static ConversionResult ConvertPdf(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = PdfDocument.Load(stream, PdfLoadOptions.Default);

        return ConvertPdf(document, settings);
    }

    static ConversionResult ConvertPdf(PdfDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetInfo(document, document.Info);
        return new(info, GetPdfStreams(document, settings).ToList());
    }

    static object GetInfo(PdfDocument document, PdfDocumentInformation info) =>
        new
        {
            document.Pages.Count,
            info.Author,
            info.CreationDate,
            info.Creator,
            info.Metadata,
            info.Keywords,
            info.ModificationDate,
            info.Producer,
            info.Subject,
            info.Title
        };

    static IEnumerable<Target> GetPdfStreams(PdfDocument document, IReadOnlyDictionary<string, object> settings)
    {
        var pages = document.Pages;
        return GetPdfStreams(document, settings, pages);
    }

    static IEnumerable<Target> GetPdfStreams(
        PdfDocument document,
        IReadOnlyDictionary<string, object> settings,
        PdfPages pages)
    {
        var pagesToInclude = settings.GetPagesToInclude(pages.Count);
        var pdfStream = new MemoryStream();
        document.Save(pdfStream);
        pdfStream.Position = 0;

        var imageOptions = new ImageSaveOptions(ImageSaveFormat.Png);

        for (var index = 0; index < pagesToInclude; index++)
        {
            var page = pages[index];

            var text = page.Content.ToString();

            yield return new("txt", text);

            var pngStream = new MemoryStream();

            imageOptions.PageNumber = index;
            document.Save(pngStream, imageOptions);

            pngStream.Position = 0;

            yield return new("png", pngStream);
        }
    }
}