using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Calendar
{
    internal sealed class OutlookCalendarReader
    {
        private readonly Outlook.Application _application;

        public OutlookCalendarReader(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public List<CalendarEventDto> ReadCalendarEvents(DateTime start, DateTime end)
        {
            var events = new List<CalendarEventDto>();
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            Outlook.Items restricted = null;

            try
            {
                folder = _application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                if (folder == null)
                    return events;

                items = folder.Items;
                if (items == null)
                    return events;

                items.IncludeRecurrences = true;
                try { items.Sort("[Start]", false); } catch { }

                var filter = string.Format(
                    CultureInfo.InvariantCulture,
                    "[Start] >= '{0}' AND [Start] < '{1}'",
                    OutlookDateFilter.FormatItemsDateTime(start),
                    OutlookDateFilter.FormatItemsDateTime(end));
                restricted = items.Restrict(filter);

                foreach (var obj in restricted)
                {
                    var appointment = obj as Outlook.AppointmentItem;
                    if (appointment == null)
                    {
                        Release(obj);
                        continue;
                    }

                    try
                    {
                        events.Add(ReadAppointment(appointment));
                    }
                    catch { }
                    finally
                    {
                        Release(appointment);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadCalendarEvents error: " + ex);
            }
            finally
            {
                Release(restricted);
                Release(items);
                Release(folder);
            }

            return events;
        }

        private static CalendarEventDto ReadAppointment(Outlook.AppointmentItem appointment)
        {
            var organizerDto = new OutlookRecipientDto
            {
                RecipientKind = "organizer",
                DisplayName = "",
                SmtpAddress = "",
                RawAddress = "",
                AddressType = "",
                EntryUserType = "",
                IsGroup = false,
                IsResolved = false,
                Members = new List<OutlookRecipientDto>()
            };
            try { organizerDto.DisplayName = appointment.Organizer ?? ""; } catch { }

            return new CalendarEventDto
            {
                Id = ReadString(() => appointment.EntryID),
                Subject = ReadString(() => appointment.Subject),
                Start = OutlookDateFilter.ToTransportUtc(ReadDate(() => appointment.Start)),
                End = OutlookDateFilter.ToTransportUtc(ReadDate(() => appointment.End)),
                Location = ReadString(() => appointment.Location),
                Organizer = organizerDto,
                RequiredAttendees = ReadAttendees(appointment),
                IsRecurring = ReadBool(() => appointment.IsRecurring),
                BusyStatus = ReadString(() => appointment.BusyStatus.ToString())
            };
        }

        private static List<OutlookRecipientDto> ReadAttendees(Outlook.AppointmentItem appointment)
        {
            var attendees = new List<OutlookRecipientDto>();
            Outlook.Recipients recipients = null;

            try { recipients = appointment.Recipients; } catch { }
            if (recipients == null)
                return attendees;

            try
            {
                for (int i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        var recipientType = Outlook.OlMeetingRecipientType.olRequired;
                        try { recipientType = (Outlook.OlMeetingRecipientType)recipient.Type; } catch { }

                        if (recipientType == Outlook.OlMeetingRecipientType.olRequired ||
                            recipientType == Outlook.OlMeetingRecipientType.olOptional)
                        {
                            var kind = recipientType == Outlook.OlMeetingRecipientType.olRequired ? "required" : "optional";
                            attendees.Add(OutlookRecipientDtoBuilder.FromRecipient(recipient, kind));
                        }
                    }
                    catch { }
                    finally
                    {
                        Release(recipient);
                    }
                }
            }
            finally
            {
                Release(recipients);
            }

            return attendees;
        }

        private static string ReadString(Func<string> read)
        {
            try { return read() ?? ""; } catch { return ""; }
        }

        private static DateTime ReadDate(Func<DateTime> read)
        {
            try { return read(); } catch { return DateTime.MinValue; }
        }

        private static bool ReadBool(Func<bool> read)
        {
            try { return read(); } catch { return false; }
        }

        private static void Release(object obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
