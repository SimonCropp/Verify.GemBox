using GemBox.Pdf;

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

        VerifierSettings.RegisterFileConverter("pdf", ConvertPdf);
        VerifierSettings.RegisterFileConverter<PdfDocument>(ConvertPdf);
    }
}