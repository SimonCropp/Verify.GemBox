public static class ApplyLicense
{
    [ModuleInitializer]
    public static void Initialize() =>
#pragma warning disable CA1416
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
#pragma warning restore CA1416
}