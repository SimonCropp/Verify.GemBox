using FreeLimitReachedAction = GemBox.Spreadsheet.FreeLimitReachedAction;

public static class ApplyLicense
{
    [ModuleInitializer]
    public static void Initialize()
    {
#pragma warning disable CA1416
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        SpreadsheetInfo.FreeLimitReached += (_, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.Stop;
        GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
#pragma warning restore CA1416
    }
}