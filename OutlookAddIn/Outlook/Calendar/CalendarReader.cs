using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        // Read calendar events from default calendar for a date range
        public List<CalendarEventDto> ReadCalendarEvents(DateTime start, DateTime end)
        {
            var events = new List<CalendarEventDto>();
            try
            {
                var folder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                if (folder == null) return events;

                var items = folder.Items;
                items.IncludeRecurrences = true;
                // Sort by start
                try { items.Sort("[Start]", false); } catch { }

                // Outlook DASL filter expects MM/dd/yyyy format in invariant culture
                var filter = string.Format("[Start] >= '{0}' AND [Start] < '{1}'", start.ToString("MM/dd/yyyy HH:mm"), end.ToString("MM/dd/yyyy HH:mm"));
                var restricted = items.Restrict(filter);

                foreach (var obj in restricted)
                {
                    var appt = obj as Outlook.AppointmentItem;
                    if (appt == null) { if (obj != null) try { Marshal.ReleaseComObject(obj); } catch { } continue; }
                    try
                    {
                        // Build structured organizer DTO
                        OutlookRecipientDto organizerDto = new OutlookRecipientDto
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
                        try { organizerDto.DisplayName = appt.Organizer ?? ""; } catch { }

                        // Build structured required attendees list
                        var requiredAttendeeDtos = new List<OutlookRecipientDto>();
                        Outlook.Recipients apptRecipients = null;
                        try { apptRecipients = appt.Recipients; } catch { }
                        if (apptRecipients != null)
                        {
                            try
                            {
                                for (int i = 1; i <= apptRecipients.Count; i++)
                                {
                                    Outlook.Recipient r = null;
                                    try
                                    {
                                        r = apptRecipients[i];
                                        Outlook.OlMeetingRecipientType rt = Outlook.OlMeetingRecipientType.olRequired;
                                        try { rt = (Outlook.OlMeetingRecipientType)r.Type; } catch { }
                                        if (rt == Outlook.OlMeetingRecipientType.olRequired ||
                                            rt == Outlook.OlMeetingRecipientType.olOptional)
                                        {
                                            string kind = rt == Outlook.OlMeetingRecipientType.olRequired ? "required" : "optional";
                                            requiredAttendeeDtos.Add(BuildRecipientDto(r, kind));
                                        }
                                    }
                                    catch { }
                                    finally { if (r != null) try { Marshal.ReleaseComObject(r); } catch { } }
                                }
                            }
                            finally { try { Marshal.ReleaseComObject(apptRecipients); } catch { } }
                        }

                        var dto = new CalendarEventDto
                        {
                            Id = appt.EntryID ?? "",
                            Subject = appt.Subject ?? "",
                            Start = appt.Start,
                            End = appt.End,
                            Location = appt.Location ?? "",
                            Organizer = organizerDto,
                            RequiredAttendees = requiredAttendeeDtos,
                            IsRecurring = appt.IsRecurring,
                            BusyStatus = appt.BusyStatus.ToString()
                        };
                        events.Add(dto);
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(appt); } catch { } }
                }

                try { Marshal.ReleaseComObject(restricted); } catch { }
                try { Marshal.ReleaseComObject(items); } catch { }
                try { Marshal.ReleaseComObject(folder); } catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadCalendarEvents error: " + ex);
            }
            return events;
        }
    }
}
