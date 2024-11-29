using System;
using System.Collections.Generic;

namespace VerifyTests;

public static class VerifyGemBoxSettings
{
    public static void PagesToInclude(this VerifySettings settings, int count) =>
        settings.Context["VerifyGemBoxPagesToInclude"] = count;

    public static SettingsTask PagesToInclude(this SettingsTask settings, int count)
    {
        settings.CurrentSettings.PagesToInclude(count);
        return settings;
    }

    internal static int GetPagesToInclude(this IReadOnlyDictionary<string, object> settings, int count)
    {
        if (!settings.TryGetValue("VerifyGemBoxPagesToInclude", out var value))
        {
            return count;
        }

        return Math.Min(count, (int) value);
    }
}