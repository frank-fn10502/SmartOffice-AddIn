using System;

namespace OutlookAddIn.Clients
{
    /// <summary>
    /// Provides BaseUrl configuration.
    /// All Outlook data push/poll and chat is now done via SignalRClient.
    /// </summary>
    public static class HubClient
    {
        private static string GetBaseUrlFromSettings()
        {
            try
            {
                var obj = global::OutlookAddIn.Properties.Settings.Default["HubBaseUrl"];
                var s = obj as string;
                if (!string.IsNullOrEmpty(s)) return s;
            }
            catch { }
            return "http://localhost:2805";
        }

        public static string BaseUrl { get; set; } = GetBaseUrlFromSettings();
    }

}
