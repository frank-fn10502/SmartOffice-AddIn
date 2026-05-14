using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Calendar
{
    internal sealed class OutlookCalendarReader
    {
        private const string SmartOfficeOwnedProperty = "SmartOfficeOwned";
        private const string SmartOfficeEventIdProperty = "SmartOfficeEventId";
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

        public List<CalendarRoomDto> ReadCalendarRooms()
        {
            var rooms = new List<CalendarRoomDto>();
            Outlook.AddressLists lists = null;
            try { lists = _application.Session.AddressLists; } catch { }
            if (lists == null) return rooms;

            try
            {
                for (int i = 1; i <= lists.Count; i++)
                {
                    Outlook.AddressList list = null;
                    Outlook.AddressEntries entries = null;
                    try
                    {
                        list = lists[i];
                        var listName = ReadString(() => list.Name);
                        if (!LooksLikeRoomList(listName)) continue;
                        entries = list.AddressEntries;
                        var max = Math.Min(entries.Count, 200);
                        for (int j = 1; j <= max; j++)
                        {
                            Outlook.AddressEntry entry = null;
                            try
                            {
                                entry = entries[j];
                                var name = ReadString(() => entry.Name);
                                var address = ReadString(() => entry.Address);
                                if (string.IsNullOrWhiteSpace(name)) continue;
                                rooms.Add(new CalendarRoomDto
                                {
                                    Id = string.IsNullOrWhiteSpace(address) ? name : address,
                                    DisplayName = name,
                                    RawAddress = address,
                                    SmtpAddress = TryReadSmtpAddress(entry),
                                    Source = listName,
                                });
                            }
                            finally
                            {
                                Release(entry);
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        Release(entries);
                        Release(list);
                    }
                }
            }
            finally
            {
                Release(lists);
            }

            return rooms
                .GroupBy(room => string.IsNullOrWhiteSpace(room.SmtpAddress) ? room.DisplayName : room.SmtpAddress, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(room => room.DisplayName)
                .ToList();
        }

        public List<CalendarEventDto> CreateCalendarEvent(CalendarEventCommandRequest request)
        {
            if (request == null) throw new InvalidOperationException("invalid_calendar_request");
            Outlook.AppointmentItem appointment = null;
            try
            {
                appointment = _application.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                if (appointment == null) throw new InvalidOperationException("calendar_create_failed");
                ApplyAppointmentFields(appointment, request, create: true);
                MarkSmartOfficeOwned(appointment, request.SmartOfficeEventId);
                appointment.Save();
                return ReadMutationWindow(appointment);
            }
            finally
            {
                Release(appointment);
            }
        }

        public List<CalendarEventDto> UpdateCalendarEvent(CalendarEventCommandRequest request)
        {
            var appointment = GetSmartOfficeAppointment(request);
            try
            {
                ApplyAppointmentFields(appointment, request, create: false);
                appointment.Save();
                return ReadMutationWindow(appointment);
            }
            finally
            {
                Release(appointment);
            }
        }

        public List<CalendarEventDto> DeleteCalendarEvent(CalendarEventCommandRequest request)
        {
            var appointment = GetSmartOfficeAppointment(request);
            DateTime start;
            try { start = appointment.Start; } catch { start = DateTime.Now; }
            try
            {
                appointment.Delete();
                return ReadCalendarEvents(start.Date.AddDays(-31), start.Date.AddDays(62));
            }
            finally
            {
                Release(appointment);
            }
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
                BusyStatus = ReadString(() => appointment.BusyStatus.ToString()),
                SmartOfficeOwned = IsSmartOfficeOwned(appointment),
                SmartOfficeEventId = ReadUserPropertyString(appointment, SmartOfficeEventIdProperty)
            };
        }

        private Outlook.AppointmentItem GetSmartOfficeAppointment(CalendarEventCommandRequest request)
        {
            var eventId = request?.EventId;
            if (string.IsNullOrWhiteSpace(eventId))
                throw new InvalidOperationException("missing_event_id");
            if (string.IsNullOrWhiteSpace(request?.SmartOfficeEventId))
                throw new InvalidOperationException("not_smartoffice_owned");

            object item = null;
            try
            {
                item = _application.Session.GetItemFromID(eventId);
                var appointment = item as Outlook.AppointmentItem;
                if (appointment == null)
                    throw new InvalidOperationException("calendar_event_not_found");
                if (!IsSmartOfficeOwned(appointment))
                    throw new InvalidOperationException("not_smartoffice_owned");
                var storedSmartOfficeEventId = ReadUserPropertyString(appointment, SmartOfficeEventIdProperty);
                if (!string.Equals(storedSmartOfficeEventId, request.SmartOfficeEventId.Trim(), StringComparison.Ordinal))
                    throw new InvalidOperationException("not_smartoffice_owned");
                item = null;
                return appointment;
            }
            finally
            {
                Release(item);
            }
        }

        private static void ApplyAppointmentFields(Outlook.AppointmentItem appointment, CalendarEventCommandRequest request, bool create)
        {
            if (string.IsNullOrWhiteSpace(request.Subject))
                throw new InvalidOperationException("missing_subject");
            if (request.Start == null || request.End == null || request.Start >= request.End)
                throw new InvalidOperationException("invalid_calendar_range");

            appointment.Subject = request.Subject.Trim();
            appointment.Start = OutlookDateFilter.ToOutlookLocalDateTime(request.Start.Value);
            appointment.End = OutlookDateFilter.ToOutlookLocalDateTime(request.End.Value);
            appointment.Location = request.Location ?? "";
            appointment.Body = request.Body ?? "";
            appointment.BusyStatus = ToBusyStatus(request.BusyStatus);
            var hasRequiredAttendees = request.RequiredAttendees != null && request.RequiredAttendees.Count > 0;
            var hasResources = request.Resources != null && request.Resources.Count > 0;
            if (hasRequiredAttendees || hasResources)
                appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            ClearRecipients(appointment);
            ApplyRequiredAttendees(appointment, request.RequiredAttendees);
            ApplyResourceRecipients(appointment, request.Resources);
        }

        private static void ClearRecipients(Outlook.AppointmentItem appointment)
        {
            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                for (int i = recipients.Count; i >= 1; i--)
                {
                    Outlook.Recipient recipient = null;
                    try
                    {
                        recipient = recipients[i];
                        recipient.Delete();
                    }
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
        }

        private static void ApplyRequiredAttendees(Outlook.AppointmentItem appointment, List<OutlookRecipientDto> attendees)
        {
            ApplyRecipients(appointment, attendees, Outlook.OlMeetingRecipientType.olRequired);
        }

        private static void ApplyResourceRecipients(Outlook.AppointmentItem appointment, List<OutlookRecipientDto> resources)
        {
            ApplyRecipients(appointment, resources, Outlook.OlMeetingRecipientType.olResource);
        }

        private static void ApplyRecipients(Outlook.AppointmentItem appointment, List<OutlookRecipientDto> recipientsToAdd, Outlook.OlMeetingRecipientType type)
        {
            if (recipientsToAdd == null) return;
            Outlook.Recipients recipients = null;
            try
            {
                recipients = appointment.Recipients;
                foreach (var attendee in recipientsToAdd)
                {
                    var address = !string.IsNullOrWhiteSpace(attendee.SmtpAddress)
                        ? attendee.SmtpAddress
                        : attendee.RawAddress;
                    if (string.IsNullOrWhiteSpace(address)) address = attendee.DisplayName;
                    if (string.IsNullOrWhiteSpace(address)) continue;
                    var recipient = recipients.Add(address);
                    recipient.Type = (int)type;
                    try { recipient.Resolve(); } catch { }
                    Release(recipient);
                }
            }
            finally
            {
                Release(recipients);
            }
        }

        private static Outlook.OlBusyStatus ToBusyStatus(string value)
        {
            switch ((value ?? "").Trim().ToLowerInvariant())
            {
                case "free": return Outlook.OlBusyStatus.olFree;
                case "tentative": return Outlook.OlBusyStatus.olTentative;
                case "outofoffice":
                case "out_of_office": return Outlook.OlBusyStatus.olOutOfOffice;
                default: return Outlook.OlBusyStatus.olBusy;
            }
        }

        private List<CalendarEventDto> ReadMutationWindow(Outlook.AppointmentItem appointment)
        {
            DateTime start;
            try { start = appointment.Start; } catch { start = DateTime.Now; }
            return ReadCalendarEvents(start.Date.AddDays(-31), start.Date.AddDays(62));
        }

        private static void MarkSmartOfficeOwned(Outlook.AppointmentItem appointment, string smartOfficeEventId)
        {
            SetUserProperty(appointment, SmartOfficeOwnedProperty, true, Outlook.OlUserPropertyType.olYesNo);
            SetUserProperty(
                appointment,
                SmartOfficeEventIdProperty,
                string.IsNullOrWhiteSpace(smartOfficeEventId) ? Guid.NewGuid().ToString() : smartOfficeEventId,
                Outlook.OlUserPropertyType.olText);
        }

        private static void SetUserProperty(Outlook.AppointmentItem appointment, string name, object value, Outlook.OlUserPropertyType type)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = appointment.UserProperties;
                property = properties.Find(name) ?? properties.Add(name, type, true);
                property.Value = value;
            }
            finally
            {
                Release(property);
                Release(properties);
            }
        }

        private static bool IsSmartOfficeOwned(Outlook.AppointmentItem appointment)
        {
            return ReadUserPropertyBool(appointment, SmartOfficeOwnedProperty);
        }

        private static bool ReadUserPropertyBool(Outlook.AppointmentItem appointment, string name)
        {
            var value = ReadUserPropertyValue(appointment, name);
            return value is bool b && b;
        }

        private static string ReadUserPropertyString(Outlook.AppointmentItem appointment, string name)
        {
            return Convert.ToString(ReadUserPropertyValue(appointment, name)) ?? "";
        }

        private static bool LooksLikeRoomList(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            return name.IndexOf("room", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("rooms", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("會議室", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("資源", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string TryReadSmtpAddress(Outlook.AddressEntry entry)
        {
            try
            {
                var user = entry.GetExchangeUser();
                if (user != null)
                    return user.PrimarySmtpAddress ?? "";
            }
            catch { }

            try { return entry.Address ?? ""; } catch { return ""; }
        }

        private static object ReadUserPropertyValue(Outlook.AppointmentItem appointment, string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = appointment.UserProperties;
                property = properties.Find(name);
                return property == null ? null : property.Value;
            }
            catch
            {
                return null;
            }
            finally
            {
                Release(property);
                Release(properties);
            }
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
