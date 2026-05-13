using System;
using System.Globalization;

namespace OutlookAddIn.OutlookServices.Common
{
    internal static class OutlookDateFilter
    {
        public static string FormatItemsDateTime(DateTime value)
        {
            return ToOutlookLocalTime(value).ToString("MM/dd/yyyy HH:mm", CultureInfo.InvariantCulture);
        }

        public static string FormatDaslDateTime(DateTime value)
        {
            return ToOutlookLocalTime(value).ToString("g", CultureInfo.CurrentCulture);
        }

        public static DateTime ToTransportUtc(DateTime value)
        {
            if (value == DateTime.MinValue || value == DateTime.MaxValue)
                return value;

            if (value.Kind == DateTimeKind.Utc)
                return value;

            if (value.Kind == DateTimeKind.Local)
                return value.ToUniversalTime();

            return DateTime.SpecifyKind(value, DateTimeKind.Local).ToUniversalTime();
        }

        public static DateTime? ToTransportUtc(DateTime? value)
        {
            return value.HasValue ? ToTransportUtc(value.Value) : (DateTime?)null;
        }

        public static DateTime ToOutlookLocalDateTime(DateTime value)
        {
            return ToOutlookLocalTime(value);
        }

        public static DateTime? ToOutlookLocalDateTime(DateTime? value)
        {
            return value.HasValue ? ToOutlookLocalDateTime(value.Value) : (DateTime?)null;
        }

        private static DateTime ToOutlookLocalTime(DateTime value)
        {
            if (value.Kind == DateTimeKind.Utc)
                return value.ToLocalTime();

            return value;
        }
    }
}
