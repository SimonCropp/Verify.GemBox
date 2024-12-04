using GemBox.Document;
using GemBox.Presentation;

namespace VerifyTests;

public static partial class VerifyGemBox
{
    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            return;
        }

        Initialized = true;

        VerifierSettings.RegisterFileConverter("xlsx", ConvertExcel);
        VerifierSettings.RegisterFileConverter("xls", ConvertExcel);
        VerifierSettings.RegisterFileConverter<ExcelFile>(ConvertExcel);

        VerifierSettings.RegisterFileConverter("pdf", ConvertPdf);
        VerifierSettings.RegisterFileConverter<PdfDocument>(ConvertPdf);

        VerifierSettings.RegisterFileConverter("docx", ConvertDocx);
        VerifierSettings.RegisterFileConverter("doc", ConvertDoc);
        VerifierSettings.RegisterFileConverter<DocumentModel>(ConvertWord);

        VerifierSettings.RegisterFileConverter("pptx", ConvertPptx);
        VerifierSettings.RegisterFileConverter("ppt", ConvertPpt);
        VerifierSettings.RegisterFileConverter<PresentationDocument>(ConvertPowerPoint);
    }
}