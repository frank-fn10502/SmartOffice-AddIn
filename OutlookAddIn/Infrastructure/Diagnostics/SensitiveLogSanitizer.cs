using System;
using System.Text.RegularExpressions;

namespace OutlookAddIn.Infrastructure.Diagnostics
{
    internal static class SensitiveLogSanitizer
    {
        public static string Sanitize(Exception ex)
        {
            if (ex == null) return "(no exception)";
            // Remove any sequences that look like Outlook folder paths starting with \\\\ (two backslashes)
            var msg = ex.Message ?? "";
            try
            {
                msg = Regex.Replace(msg, "\\\\\\\\[^\r\n\"]*", "[redacted]");
            }
            catch { msg = "(error sanitizing message)"; }
            return ex.GetType().Name + ": " + msg;
        }
    }
}
