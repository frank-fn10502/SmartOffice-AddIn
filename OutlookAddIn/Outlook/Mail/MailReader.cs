using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private const int FetchMailsMaxCount = 100;
        private static readonly TimeSpan FetchMailsTableBudget = TimeSpan.FromSeconds(8);

        /// <summary>
        /// Builds an OutlookRecipientDto from a resolved Outlook Recipient COM object.
        /// Caller is responsible for releasing the COM object.
        /// </summary>
        private static OutlookRecipientDto BuildRecipientDto(Outlook.Recipient r, string kind)
        {
            var dto = new OutlookRecipientDto
            {
                RecipientKind = kind,
                Members = new List<OutlookRecipientDto>()
            };
            try { dto.DisplayName = r.Name ?? ""; } catch { dto.DisplayName = ""; }
            try { dto.RawAddress = r.Address ?? ""; } catch { dto.RawAddress = ""; }

            Outlook.AddressEntry ae = null;
            try { ae = r.AddressEntry; } catch { }
            if (ae != null)
            {
                try
                {
                    try { dto.SmtpAddress = ResolveSmtpFromAddressEntry(ae); } catch { }
                    try { dto.AddressType = ae.AddressEntryUserType.ToString(); } catch { dto.AddressType = ""; }
                    try { dto.EntryUserType = ae.AddressEntryUserType.ToString(); } catch { dto.EntryUserType = ""; }
                    try { dto.IsGroup = ae.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry; } catch { }
                    dto.IsResolved = true;
                }
                finally { try { Marshal.ReleaseComObject(ae); } catch { } }
            }
            if (string.IsNullOrEmpty(dto.SmtpAddress)) dto.SmtpAddress = "";
            if (string.IsNullOrEmpty(dto.AddressType)) dto.AddressType = "";
            if (string.IsNullOrEmpty(dto.EntryUserType)) dto.EntryUserType = "";
            return dto;
        }

        private static string ResolveSmtpFromAddressEntry(Outlook.AddressEntry ae)
        {
            if (ae == null) return "";
            try
            {
                // For Exchange users, GetExchangeUser resolves the SMTP address.
                if (ae.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                {
                    var eu = ae.GetExchangeUser();
                    if (eu != null)
                    {
                        string smtp = "";
                        try { smtp = eu.PrimarySmtpAddress ?? ""; } finally { try { Marshal.ReleaseComObject(eu); } catch { } }
                        if (!string.IsNullOrEmpty(smtp)) return smtp;
                    }
                }
            }
            catch { }
            try { return ae.Address ?? ""; } catch { return ""; }
        }

        private static OutlookRecipientDto BuildSenderDto(Outlook.MailItem mail)
        {
            var dto = new OutlookRecipientDto
            {
                RecipientKind = "sender",
                Members = new List<OutlookRecipientDto>()
            };
            try { dto.DisplayName = mail.SenderName ?? ""; } catch { dto.DisplayName = ""; }
            try { dto.RawAddress = mail.SenderEmailAddress ?? ""; } catch { dto.RawAddress = ""; }

            Outlook.AddressEntry ae = null;
            try { ae = mail.Sender; } catch { }
            if (ae != null)
            {
                try
                {
                    try { dto.SmtpAddress = ResolveSmtpFromAddressEntry(ae); } catch { }
                    try { dto.AddressType = ae.AddressEntryUserType.ToString(); } catch { dto.AddressType = ""; }
                    try { dto.EntryUserType = ae.AddressEntryUserType.ToString(); } catch { dto.EntryUserType = ""; }
                    try { dto.IsGroup = ae.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry; } catch { }
                    dto.IsResolved = true;
                }
                finally { try { Marshal.ReleaseComObject(ae); } catch { } }
            }
            if (string.IsNullOrEmpty(dto.SmtpAddress)) dto.SmtpAddress = "";
            if (string.IsNullOrEmpty(dto.AddressType)) dto.AddressType = "";
            if (string.IsNullOrEmpty(dto.EntryUserType)) dto.EntryUserType = "";
            return dto;
        }

        private static OutlookRecipientDto BuildSenderDto(Outlook.MeetingItem meeting)
        {
            var dto = new OutlookRecipientDto
            {
                RecipientKind = "sender",
                Members = new List<OutlookRecipientDto>()
            };
            try { dto.DisplayName = meeting.SenderName ?? ""; } catch { dto.DisplayName = ""; }
            try { dto.RawAddress = meeting.SenderEmailAddress ?? ""; } catch { dto.RawAddress = ""; }
            dto.SmtpAddress = dto.RawAddress ?? "";
            dto.AddressType = "";
            dto.EntryUserType = "";
            return dto;
        }

        private static List<OutlookRecipientDto> BuildRecipientsDto(Outlook.Recipients recipients, string kind)
        {
            var list = new List<OutlookRecipientDto>();
            if (recipients == null) return list;
            try
            {
                for (int i = 1; i <= recipients.Count; i++)
                {
                    Outlook.Recipient r = null;
                    try
                    {
                        r = recipients[i];
                        list.Add(BuildRecipientDto(r, kind));
                    }
                    catch { }
                    finally { if (r != null) try { Marshal.ReleaseComObject(r); } catch { } }
                }
            }
            catch { }
            return list;
        }

        /// <summary>
        /// Reads a single MailItem into a MailItemDto.
        /// When includeBody is false, body/bodyHtml are left empty (used for list/metadata push).
        /// </summary>
        public MailItemDto ReadSingleMailDto(Outlook.MailItem mail, string folderPath, bool includeBody = false)
        {
            if (mail == null) return null;
            try
            {
                string entryId = "";
                try { entryId = mail.EntryID ?? ""; } catch { }

                string subject = "";
                try { subject = mail.Subject ?? ""; } catch { }

                OutlookRecipientDto sender = BuildSenderDto(mail);

                List<OutlookRecipientDto> toRecipients = new List<OutlookRecipientDto>();
                List<OutlookRecipientDto> ccRecipients = new List<OutlookRecipientDto>();
                List<OutlookRecipientDto> bccRecipients = new List<OutlookRecipientDto>();
                try
                {
                    Outlook.Recipients allRecipients = null;
                    try { allRecipients = mail.Recipients; } catch { }
                    if (allRecipients != null)
                    {
                        try
                        {
                            for (int i = 1; i <= allRecipients.Count; i++)
                            {
                                Outlook.Recipient r = null;
                                try
                                {
                                    r = allRecipients[i];
                                    Outlook.OlMailRecipientType rt = Outlook.OlMailRecipientType.olTo;
                                    try { rt = (Outlook.OlMailRecipientType)r.Type; } catch { }
                                    switch (rt)
                                    {
                                        case Outlook.OlMailRecipientType.olCC:
                                            ccRecipients.Add(BuildRecipientDto(r, "cc")); break;
                                        case Outlook.OlMailRecipientType.olBCC:
                                            bccRecipients.Add(BuildRecipientDto(r, "bcc")); break;
                                        default:
                                            toRecipients.Add(BuildRecipientDto(r, "to")); break;
                                    }
                                }
                                catch { }
                                finally { if (r != null) try { Marshal.ReleaseComObject(r); } catch { } }
                            }
                        }
                        finally { try { Marshal.ReleaseComObject(allRecipients); } catch { } }
                    }
                }
                catch { }

                DateTime receivedTime = DateTime.MinValue;
                try { receivedTime = mail.ReceivedTime; } catch { }

                string body = "";
                string bodyHtml = "";
                if (includeBody)
                {
                    try { body = mail.Body ?? ""; } catch { }
                    try { bodyHtml = mail.HTMLBody ?? ""; } catch { }
                }

                string categories = "";
                try { categories = mail.Categories ?? ""; } catch { }

                string conversationId = "";
                try { conversationId = mail.ConversationID ?? ""; } catch { }
                string conversationTopic = "";
                try { conversationTopic = mail.ConversationTopic ?? ""; } catch { }
                string conversationIndex = "";
                try { conversationIndex = mail.ConversationIndex ?? ""; } catch { }

                string messageClass = "";
                try { messageClass = mail.MessageClass ?? ""; } catch { }

                bool isRead = false;
                try { isRead = !mail.UnRead; } catch { }

                bool isMarkedAsTask = false;
                try { isMarkedAsTask = mail.IsMarkedAsTask; } catch { }

                string flagRequest = "";
                try { flagRequest = mail.FlagRequest ?? ""; } catch { }

                string flagInterval = "none";
                try
                {
                    var fs = mail.FlagStatus;
                    if (fs == Outlook.OlFlagStatus.olFlagMarked) flagInterval = "custom";
                    else if (fs == Outlook.OlFlagStatus.olFlagComplete) flagInterval = "complete";
                    else flagInterval = "none";
                }
                catch { }

                DateTime? taskStartDate = null;
                try { if (isMarkedAsTask) taskStartDate = mail.TaskStartDate; } catch { }

                DateTime? taskDueDate = null;
                try { if (isMarkedAsTask) taskDueDate = mail.TaskDueDate; } catch { }

                DateTime? taskCompletedDate = null;
                try { if (isMarkedAsTask) taskCompletedDate = mail.TaskCompletedDate; } catch { }

                string importance = "normal";
                try
                {
                    var imp = mail.Importance;
                    if (imp == Outlook.OlImportance.olImportanceLow) importance = "low";
                    else if (imp == Outlook.OlImportance.olImportanceHigh) importance = "high";
                }
                catch { }

                string sensitivity = "normal";
                try
                {
                    var s = mail.Sensitivity;
                    if (s == Outlook.OlSensitivity.olPersonal) sensitivity = "personal";
                    else if (s == Outlook.OlSensitivity.olPrivate) sensitivity = "private";
                    else if (s == Outlook.OlSensitivity.olConfidential) sensitivity = "confidential";
                }
                catch { }

                if (string.IsNullOrEmpty(folderPath))
                {
                    try { folderPath = ((Outlook.MAPIFolder)mail.Parent)?.FolderPath ?? ""; } catch { folderPath = ""; }
                }

                int attachmentCount = 0;
                string attachmentNames = "";
                try
                {
                    var atts = mail.Attachments;
                    if (atts != null)
                    {
                        attachmentCount = atts.Count;
                        var names = new List<string>();
                        for (int i = 1; i <= atts.Count; i++)
                        {
                            Outlook.Attachment a = null;
                            try
                            {
                                a = atts[i];
                                string fn = "";
                                try { fn = a.FileName ?? ""; } catch { }
                                if (!string.IsNullOrEmpty(fn)) names.Add(fn);
                            }
                            catch { }
                            finally { if (a != null) try { Marshal.ReleaseComObject(a); } catch { } }
                        }
                        attachmentNames = string.Join(", ", names);
                        try { Marshal.ReleaseComObject(atts); } catch { }
                    }
                }
                catch { }

                return new MailItemDto
                {
                    Id = entryId,
                    Subject = subject,
                    Sender = sender,
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipients,
                    BccRecipients = bccRecipients,
                    ReceivedTime = OutlookDateFilter.ToTransportUtc(receivedTime == DateTime.MinValue ? DateTime.Now : receivedTime),
                    Body = body,
                    BodyHtml = bodyHtml,
                    FolderPath = folderPath,
                    MessageClass = messageClass,
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    ConversationIndex = conversationIndex,
                    Categories = categories,
                    IsRead = isRead,
                    IsMarkedAsTask = isMarkedAsTask,
                    AttachmentCount = attachmentCount,
                    AttachmentNames = attachmentNames,
                    FlagRequest = flagRequest,
                    FlagInterval = flagInterval,
                    TaskStartDate = OutlookDateFilter.ToTransportUtc(taskStartDate),
                    TaskDueDate = OutlookDateFilter.ToTransportUtc(taskDueDate),
                    TaskCompletedDate = OutlookDateFilter.ToTransportUtc(taskCompletedDate),
                    Importance = importance,
                    Sensitivity = sensitivity
                };
            }
            catch
            {
                return null;
            }
        }

        public MailItemDto ReadSingleMailDto(Outlook.MeetingItem meeting, string folderPath, bool includeBody = false)
        {
            if (meeting == null) return null;
            try
            {
                string entryId = "";
                try { entryId = meeting.EntryID ?? ""; } catch { }

                string subject = "";
                try { subject = meeting.Subject ?? ""; } catch { }

                DateTime receivedTime = DateTime.MinValue;
                try { receivedTime = meeting.ReceivedTime; } catch { }

                string body = "";
                string bodyHtml = "";
                if (includeBody)
                {
                    try { body = meeting.Body ?? ""; } catch { }
                    bodyHtml = TryReadComStringProperty(meeting, "HTMLBody");
                }

                string categories = "";
                try { categories = meeting.Categories ?? ""; } catch { }

                string conversationId = "";
                try { conversationId = meeting.ConversationID ?? ""; } catch { }
                string conversationTopic = "";
                try { conversationTopic = meeting.ConversationTopic ?? ""; } catch { }
                string conversationIndex = "";
                try { conversationIndex = meeting.ConversationIndex ?? ""; } catch { }

                string messageClass = "";
                try { messageClass = meeting.MessageClass ?? ""; } catch { }

                bool isRead = false;
                try { isRead = !meeting.UnRead; } catch { }

                string importance = "normal";
                try
                {
                    var imp = meeting.Importance;
                    if (imp == Outlook.OlImportance.olImportanceLow) importance = "low";
                    else if (imp == Outlook.OlImportance.olImportanceHigh) importance = "high";
                }
                catch { }

                string sensitivity = "normal";
                try
                {
                    var s = meeting.Sensitivity;
                    if (s == Outlook.OlSensitivity.olPersonal) sensitivity = "personal";
                    else if (s == Outlook.OlSensitivity.olPrivate) sensitivity = "private";
                    else if (s == Outlook.OlSensitivity.olConfidential) sensitivity = "confidential";
                }
                catch { }

                if (string.IsNullOrEmpty(folderPath))
                    folderPath = GetOutlookItemFolderPath(meeting);

                int attachmentCount = 0;
                string attachmentNames = "";
                Outlook.Attachments atts = null;
                try
                {
                    atts = meeting.Attachments;
                    if (atts != null)
                    {
                        attachmentCount = atts.Count;
                        var names = new List<string>();
                        for (int i = 1; i <= atts.Count; i++)
                        {
                            Outlook.Attachment a = null;
                            try
                            {
                                a = atts[i];
                                string fn = "";
                                try { fn = a.FileName ?? ""; } catch { }
                                if (!string.IsNullOrEmpty(fn)) names.Add(fn);
                            }
                            catch { }
                            finally { if (a != null) try { Marshal.ReleaseComObject(a); } catch { } }
                        }
                        attachmentNames = string.Join(", ", names);
                    }
                }
                finally { if (atts != null) try { Marshal.ReleaseComObject(atts); } catch { } }

                return new MailItemDto
                {
                    Id = entryId,
                    Subject = subject,
                    Sender = BuildSenderDto(meeting),
                    ToRecipients = new List<OutlookRecipientDto>(),
                    CcRecipients = new List<OutlookRecipientDto>(),
                    BccRecipients = new List<OutlookRecipientDto>(),
                    ReceivedTime = OutlookDateFilter.ToTransportUtc(receivedTime == DateTime.MinValue ? DateTime.Now : receivedTime),
                    Body = body,
                    BodyHtml = bodyHtml,
                    FolderPath = folderPath,
                    MessageClass = messageClass,
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    ConversationIndex = conversationIndex,
                    Categories = categories,
                    IsRead = isRead,
                    IsMarkedAsTask = false,
                    AttachmentCount = attachmentCount,
                    AttachmentNames = attachmentNames,
                    FlagRequest = "",
                    FlagInterval = "none",
                    TaskStartDate = null,
                    TaskDueDate = null,
                    TaskCompletedDate = null,
                    Importance = importance,
                    Sensitivity = sensitivity
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Builds the lean metadata used by list/search/slice commands.
        /// Avoids expensive AddressEntry, Recipients and attachment-name enumeration.
        /// Must be called on the Outlook STA thread.
        /// </summary>
        private MailItemDto ReadMailListMetadataDto(Outlook.MailItem mail, string folderPath)
        {
            if (mail == null) return null;
            try
            {
                string entryId = "";
                try { entryId = mail.EntryID ?? ""; } catch { }

                string subject = "";
                try { subject = mail.Subject ?? ""; } catch { }

                var sender = new OutlookRecipientDto
                {
                    RecipientKind = "sender",
                    Members = new List<OutlookRecipientDto>()
                };
                try { sender.DisplayName = mail.SenderName ?? ""; } catch { sender.DisplayName = ""; }
                try { sender.RawAddress = mail.SenderEmailAddress ?? ""; } catch { sender.RawAddress = ""; }
                sender.SmtpAddress = sender.RawAddress ?? "";
                sender.AddressType = "";
                sender.EntryUserType = "";

                DateTime receivedTime = DateTime.MinValue;
                try { receivedTime = mail.ReceivedTime; } catch { }

                string categories = "";
                try { categories = mail.Categories ?? ""; } catch { }

                string conversationId = "";
                try { conversationId = mail.ConversationID ?? ""; } catch { }
                string conversationTopic = "";
                try { conversationTopic = mail.ConversationTopic ?? ""; } catch { }
                string conversationIndex = "";
                try { conversationIndex = mail.ConversationIndex ?? ""; } catch { }

                string messageClass = "";
                try { messageClass = mail.MessageClass ?? ""; } catch { }

                bool isRead = false;
                try { isRead = !mail.UnRead; } catch { }

                bool isMarkedAsTask = false;
                try { isMarkedAsTask = mail.IsMarkedAsTask; } catch { }

                string flagRequest = "";
                try { flagRequest = mail.FlagRequest ?? ""; } catch { }

                string flagInterval = "none";
                try
                {
                    var fs = mail.FlagStatus;
                    if (fs == Outlook.OlFlagStatus.olFlagMarked) flagInterval = "custom";
                    else if (fs == Outlook.OlFlagStatus.olFlagComplete) flagInterval = "complete";
                }
                catch { }

                DateTime? taskStartDate = null;
                try { if (isMarkedAsTask) taskStartDate = mail.TaskStartDate; } catch { }

                DateTime? taskDueDate = null;
                try { if (isMarkedAsTask) taskDueDate = mail.TaskDueDate; } catch { }

                DateTime? taskCompletedDate = null;
                try { if (isMarkedAsTask) taskCompletedDate = mail.TaskCompletedDate; } catch { }

                string importance = "normal";
                try
                {
                    var imp = mail.Importance;
                    if (imp == Outlook.OlImportance.olImportanceLow) importance = "low";
                    else if (imp == Outlook.OlImportance.olImportanceHigh) importance = "high";
                }
                catch { }

                string sensitivity = "normal";
                try
                {
                    var s = mail.Sensitivity;
                    if (s == Outlook.OlSensitivity.olPersonal) sensitivity = "personal";
                    else if (s == Outlook.OlSensitivity.olPrivate) sensitivity = "private";
                    else if (s == Outlook.OlSensitivity.olConfidential) sensitivity = "confidential";
                }
                catch { }

                if (string.IsNullOrEmpty(folderPath))
                {
                    Outlook.MAPIFolder parent = null;
                    try
                    {
                        parent = mail.Parent as Outlook.MAPIFolder;
                        folderPath = parent?.FolderPath ?? "";
                    }
                    catch { folderPath = ""; }
                    finally { if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { } }
                }

                int attachmentCount = 0;
                Outlook.Attachments atts = null;
                try
                {
                    atts = mail.Attachments;
                    if (atts != null) attachmentCount = atts.Count;
                }
                catch { }
                finally { if (atts != null) try { Marshal.ReleaseComObject(atts); } catch { } }

                return new MailItemDto
                {
                    Id = entryId,
                    Subject = subject,
                    Sender = sender,
                    ToRecipients = new List<OutlookRecipientDto>(),
                    CcRecipients = new List<OutlookRecipientDto>(),
                    BccRecipients = new List<OutlookRecipientDto>(),
                    ReceivedTime = OutlookDateFilter.ToTransportUtc(receivedTime == DateTime.MinValue ? DateTime.Now : receivedTime),
                    Body = "",
                    BodyHtml = "",
                    FolderPath = folderPath,
                    MessageClass = messageClass,
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    ConversationIndex = conversationIndex,
                    Categories = categories,
                    IsRead = isRead,
                    IsMarkedAsTask = isMarkedAsTask,
                    AttachmentCount = attachmentCount,
                    AttachmentNames = "",
                    FlagRequest = flagRequest,
                    FlagInterval = flagInterval,
                    TaskStartDate = OutlookDateFilter.ToTransportUtc(taskStartDate),
                    TaskDueDate = OutlookDateFilter.ToTransportUtc(taskDueDate),
                    TaskCompletedDate = OutlookDateFilter.ToTransportUtc(taskCompletedDate),
                    Importance = importance,
                    Sensitivity = sensitivity
                };
            }
            catch
            {
                return null;
            }
        }

        private MailItemDto ReadMailListMetadataDto(Outlook.MeetingItem meeting, string folderPath)
        {
            if (meeting == null) return null;
            try
            {
                string entryId = "";
                try { entryId = meeting.EntryID ?? ""; } catch { }

                string subject = "";
                try { subject = meeting.Subject ?? ""; } catch { }

                DateTime receivedTime = DateTime.MinValue;
                try { receivedTime = meeting.ReceivedTime; } catch { }

                string categories = "";
                try { categories = meeting.Categories ?? ""; } catch { }

                string conversationId = "";
                try { conversationId = meeting.ConversationID ?? ""; } catch { }
                string conversationTopic = "";
                try { conversationTopic = meeting.ConversationTopic ?? ""; } catch { }
                string conversationIndex = "";
                try { conversationIndex = meeting.ConversationIndex ?? ""; } catch { }

                string messageClass = "";
                try { messageClass = meeting.MessageClass ?? ""; } catch { }

                bool isRead = false;
                try { isRead = !meeting.UnRead; } catch { }

                string importance = "normal";
                try
                {
                    var imp = meeting.Importance;
                    if (imp == Outlook.OlImportance.olImportanceLow) importance = "low";
                    else if (imp == Outlook.OlImportance.olImportanceHigh) importance = "high";
                }
                catch { }

                string sensitivity = "normal";
                try
                {
                    var s = meeting.Sensitivity;
                    if (s == Outlook.OlSensitivity.olPersonal) sensitivity = "personal";
                    else if (s == Outlook.OlSensitivity.olPrivate) sensitivity = "private";
                    else if (s == Outlook.OlSensitivity.olConfidential) sensitivity = "confidential";
                }
                catch { }

                if (string.IsNullOrEmpty(folderPath))
                    folderPath = GetOutlookItemFolderPath(meeting);

                int attachmentCount = 0;
                Outlook.Attachments atts = null;
                try
                {
                    atts = meeting.Attachments;
                    if (atts != null) attachmentCount = atts.Count;
                }
                catch { }
                finally { if (atts != null) try { Marshal.ReleaseComObject(atts); } catch { } }

                return new MailItemDto
                {
                    Id = entryId,
                    Subject = subject,
                    Sender = BuildSenderDto(meeting),
                    ToRecipients = new List<OutlookRecipientDto>(),
                    CcRecipients = new List<OutlookRecipientDto>(),
                    BccRecipients = new List<OutlookRecipientDto>(),
                    ReceivedTime = OutlookDateFilter.ToTransportUtc(receivedTime == DateTime.MinValue ? DateTime.Now : receivedTime),
                    Body = "",
                    BodyHtml = "",
                    FolderPath = folderPath,
                    MessageClass = messageClass,
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    ConversationIndex = conversationIndex,
                    Categories = categories,
                    IsRead = isRead,
                    IsMarkedAsTask = false,
                    AttachmentCount = attachmentCount,
                    AttachmentNames = "",
                    FlagRequest = "",
                    FlagInterval = "none",
                    TaskStartDate = null,
                    TaskDueDate = null,
                    TaskCompletedDate = null,
                    Importance = importance,
                    Sensitivity = sensitivity
                };
            }
            catch
            {
                return null;
            }
        }

        private MailItemDto ReadMailListMetadataDto(object item, string folderPath)
        {
            var mail = item as Outlook.MailItem;
            if (mail != null) return ReadMailListMetadataDto(mail, folderPath);

            var meeting = item as Outlook.MeetingItem;
            if (meeting != null) return ReadMailListMetadataDto(meeting, folderPath);

            return null;
        }

        /// <summary>
        /// Parses a date range string in the form "yyyy/MM/dd ~ yyyy/MM/dd" or
        /// "yyyy-MM-dd HH:mm ~ yyyy-MM-dd HH:mm" into start/end DateTimes.
        /// Returns false if the string is not a range expression.
        /// </summary>
        private static bool TryParseDateRangeString(string range, out DateTime from, out DateTime to)
        {
            from = DateTime.MinValue;
            to = DateTime.MaxValue;
            if (string.IsNullOrWhiteSpace(range)) return false;

            int tilde = range.IndexOf('~');
            if (tilde < 0) return false;

            string left = range.Substring(0, tilde).Trim();
            string right = range.Substring(tilde + 1).Trim();

            string[] formats = new[]
            {
                "yyyy/MM/dd HH:mm", "yyyy/MM/dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd",
                "yyyy/M/d HH:mm", "yyyy/M/d",
            };

            if (!DateTime.TryParseExact(left, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out from) &&
                !DateTime.TryParse(left, out from))
                return false;

            if (!DateTime.TryParseExact(right, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out to) &&
                !DateTime.TryParse(right, out to))
                return false;

            // date-only end point means end of that day
            if (!right.Contains(":"))
                to = to.Date.AddDays(1).AddSeconds(-1);

            return true;
        }

        public List<MailItemDto> ReadMails(FetchMailsRequest req)
        {
            List<MailItemDto> mails;
            string error;
            if (TryReadMailsFast(req, out mails, out error)) return mails;

            System.Diagnostics.Debug.WriteLine("ReadMails error: " + error);
            return mails ?? new List<MailItemDto>();
        }

        public bool TryReadMailsFast(FetchMailsRequest req, out List<MailItemDto> mails, out string error)
        {
            mails = new List<MailItemDto>();
            error = "";
            Outlook.MAPIFolder folder = null;
            try
            {
                int maxCount = req.MaxCount > 0 ? req.MaxCount : 100;
                if (maxCount > FetchMailsMaxCount) maxCount = FetchMailsMaxCount;

                if (string.IsNullOrEmpty(req.FolderPath))
                {
                    folder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                }
                else
                {
                    folder = GetFolderByPath(req.FolderPath);
                }

                if (folder == null)
                {
                    error = "folder_not_found";
                    System.Diagnostics.Debug.WriteLine("ReadMails: folder is null for path: " + req.FolderPath);
                    return false;
                }

                // Determine date range: Hub sends receivedFrom/receivedTo directly.
                DateTime since = DateTime.MinValue;
                DateTime until = DateTime.MaxValue;

                if (req.ReceivedFrom.HasValue || req.ReceivedTo.HasValue)
                {
                    if (req.ReceivedFrom.HasValue) since = req.ReceivedFrom.Value;
                    if (req.ReceivedTo.HasValue) until = req.ReceivedTo.Value;
                }
                else if (!string.IsNullOrEmpty(req.Range))
                {
                    // Legacy fallback: only used if Hub doesn't send receivedFrom/receivedTo
                    if (TryParseDateRangeString(req.Range, out DateTime rangeFrom, out DateTime rangeTo))
                    {
                        since = rangeFrom;
                        until = rangeTo;
                    }
                    else
                    {
                        switch (req.Range)
                        {
                            case "1d": since = DateTime.Now.AddDays(-1); break;
                            case "1w": since = DateTime.Now.AddDays(-7); break;
                            case "60d": since = DateTime.Now.AddDays(-60); break;
                            case "90d": since = DateTime.Now.AddDays(-90); break;
                            default: since = DateTime.Now.AddDays(-30); break;
                        }
                    }
                }
                else
                {
                    since = DateTime.Now.AddDays(-30);
                }

                string filterExpr;
                if (since > DateTime.MinValue && until < DateTime.MaxValue)
                    filterExpr = string.Format("[ReceivedTime] >= '{0}' AND [ReceivedTime] <= '{1}'",
                        OutlookDateFilter.FormatItemsDateTime(since), OutlookDateFilter.FormatItemsDateTime(until));
                else if (since > DateTime.MinValue)
                    filterExpr = string.Format("[ReceivedTime] >= '{0}'", OutlookDateFilter.FormatItemsDateTime(since));
                else if (until < DateTime.MaxValue)
                    filterExpr = string.Format("[ReceivedTime] <= '{0}'", OutlookDateFilter.FormatItemsDateTime(until));
                else
                    filterExpr = null;

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                if (!TryReadMailsFromTable(folder, currentFolderPath, filterExpr, maxCount, mails, out error))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                error = OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex);
                System.Diagnostics.Debug.WriteLine("ReadMails error: " + ex);
            }
            finally
            {
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }
            return false;
        }

        private bool TryReadMailsFromTable(
            Outlook.MAPIFolder folder,
            string folderPath,
            string filterExpr,
            int maxCount,
            List<MailItemDto> mails,
            out string error)
        {
            error = "";
            Outlook.Table table = null;
            try
            {
                var stopwatch = Stopwatch.StartNew();
                table = folder.GetTable(filterExpr ?? "", Outlook.OlTableContents.olUserItems);
                TryResetTableColumns(table,
                    "EntryID",
                    "MessageClass",
                    "Subject",
                    "SenderName",
                    "SenderEmailAddress",
                    "ReceivedTime",
                    "Categories",
                    "ConversationID",
                    "ConversationTopic",
                    "ConversationIndex",
                    "UnRead",
                    "IsMarkedAsTask",
                    "FlagRequest",
                    "FlagStatus",
                    "Importance",
                    "Sensitivity");
                table.Sort("[ReceivedTime]", true);

                int count = 0;
                while (!table.EndOfTable && count < maxCount)
                {
                    if (stopwatch.Elapsed > FetchMailsTableBudget)
                    {
                        error = "fetch_mails timed out while reading Outlook table metadata";
                        return false;
                    }

                    Outlook.Row row = null;
                    try
                    {
                        row = table.GetNextRow();
                        var dto = BuildMailListMetadataDtoFromTableRow(row, folderPath);
                        if (dto == null || string.IsNullOrEmpty(dto.Id)) continue;
                        mails.Add(dto);
                        count++;
                    }
                    catch { }
                    finally { if (row != null) try { Marshal.ReleaseComObject(row); } catch { } }
                }
                return true;
            }
            catch (Exception ex)
            {
                error = OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex);
                System.Diagnostics.Debug.WriteLine("TryReadMailsFromTable failed: " + ex.Message);
                mails.Clear();
                return false;
            }
            finally
            {
                if (table != null) try { Marshal.ReleaseComObject(table); } catch { }
            }
        }

        private static void TryResetTableColumns(Outlook.Table table, params string[] columns)
        {
            Outlook.Columns tableColumns = null;
            try
            {
                tableColumns = table.Columns;
                try { tableColumns.RemoveAll(); } catch { }
                foreach (var column in columns)
                {
                    try { tableColumns.Add(column); } catch { }
                }
            }
            finally
            {
                if (tableColumns != null) try { Marshal.ReleaseComObject(tableColumns); } catch { }
            }
        }

        private static MailItemDto BuildMailListMetadataDtoFromTableRow(Outlook.Row row, string folderPath)
        {
            if (row == null) return null;

            var senderAddress = TableString(row, "SenderEmailAddress");
            var receivedTime = TableDate(row, "ReceivedTime") ?? DateTime.Now;
            var isMarkedAsTask = TableBool(row, "IsMarkedAsTask");
            var flagStatus = TableInt(row, "FlagStatus");
            var importanceValue = TableInt(row, "Importance");
            var sensitivityValue = TableInt(row, "Sensitivity");

            return new MailItemDto
            {
                Id = TableString(row, "EntryID"),
                Subject = TableString(row, "Subject"),
                Sender = new OutlookRecipientDto
                {
                    RecipientKind = "sender",
                    DisplayName = TableString(row, "SenderName"),
                    RawAddress = senderAddress,
                    SmtpAddress = senderAddress,
                    AddressType = "",
                    EntryUserType = "",
                    Members = new List<OutlookRecipientDto>()
                },
                ToRecipients = new List<OutlookRecipientDto>(),
                CcRecipients = new List<OutlookRecipientDto>(),
                BccRecipients = new List<OutlookRecipientDto>(),
                ReceivedTime = OutlookDateFilter.ToTransportUtc(receivedTime),
                Body = "",
                BodyHtml = "",
                FolderPath = folderPath,
                MessageClass = TableString(row, "MessageClass"),
                ConversationId = TableString(row, "ConversationID"),
                ConversationTopic = TableString(row, "ConversationTopic"),
                ConversationIndex = TableString(row, "ConversationIndex"),
                Categories = TableString(row, "Categories"),
                IsRead = !TableBool(row, "UnRead"),
                IsMarkedAsTask = isMarkedAsTask,
                AttachmentCount = 0,
                AttachmentNames = "",
                FlagRequest = TableString(row, "FlagRequest"),
                FlagInterval = flagStatus == (int)Outlook.OlFlagStatus.olFlagMarked
                    ? "custom"
                    : flagStatus == (int)Outlook.OlFlagStatus.olFlagComplete ? "complete" : "none",
                Importance = importanceValue == (int)Outlook.OlImportance.olImportanceHigh
                    ? "high"
                    : importanceValue == (int)Outlook.OlImportance.olImportanceLow ? "low" : "normal",
                Sensitivity = sensitivityValue == (int)Outlook.OlSensitivity.olPersonal
                    ? "personal"
                    : sensitivityValue == (int)Outlook.OlSensitivity.olPrivate
                        ? "private"
                        : sensitivityValue == (int)Outlook.OlSensitivity.olConfidential ? "confidential" : "normal"
            };
        }

        private static string TableString(Outlook.Row row, string name)
        {
            try { return Convert.ToString(row[name]) ?? ""; } catch { return ""; }
        }

        private static bool TableBool(Outlook.Row row, string name)
        {
            try
            {
                var value = row[name];
                if (value is bool b) return b;
                bool parsed;
                return bool.TryParse(Convert.ToString(value), out parsed) && parsed;
            }
            catch { return false; }
        }

        private static int TableInt(Outlook.Row row, string name)
        {
            try
            {
                var value = row[name];
                if (value is int i) return i;
                int parsed;
                return int.TryParse(Convert.ToString(value), out parsed) ? parsed : 0;
            }
            catch { return 0; }
        }

        private static DateTime? TableDate(Outlook.Row row, string name)
        {
            try
            {
                var value = row[name];
                if (value is DateTime dt) return dt;
                DateTime parsed;
                return DateTime.TryParse(Convert.ToString(value), out parsed) ? parsed : (DateTime?)null;
            }
            catch { return null; }
        }

        private Outlook.MAPIFolder GetFolderByPath(string path)
        {
            try
            {
                var stores = this.Application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    try
                    {
                        var root = store.GetRootFolder();
                        var found = NavigateToFolder(root, path);
                        if (found != null)
                        {
                            Marshal.ReleaseComObject(store);
                            Marshal.ReleaseComObject(stores);
                            return found;
                        }
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(store); } catch { } }
                }
                try { Marshal.ReleaseComObject(stores); } catch { }
            }
            catch { }
            return null;
        }

        private Outlook.MAPIFolder NavigateToFolder(Outlook.MAPIFolder current, string targetPath)
        {
            if (current.FolderPath == targetPath)
                return current;

            var subFolders = current.Folders;
            Outlook.MAPIFolder result = null;
            foreach (Outlook.MAPIFolder sub in subFolders)
            {
                if (result == null && targetPath.StartsWith(sub.FolderPath))
                {
                    result = NavigateToFolder(sub, targetPath);
                    if (result == null) Marshal.ReleaseComObject(sub);
                }
                else
                {
                    Marshal.ReleaseComObject(sub);
                }
            }
            Marshal.ReleaseComObject(subFolders);
            if (result == null) Marshal.ReleaseComObject(current);
            return result;
        }

        /// <summary>
        /// Reads the body and HTML body of a single mail by EntryID.
        /// </summary>
        public MailBodyDto ReadMailBody(string mailId, string folderPath)
        {
            object item = null;
            try
            {
                item = FindOutlookItemByEntryId(mailId);
                if (item == null) return null;

                string body = "";
                string bodyHtml = "";
                var mail = item as Outlook.MailItem;
                var meeting = item as Outlook.MeetingItem;
                if (mail != null)
                {
                    try { body = mail.Body ?? ""; } catch { }
                    try { bodyHtml = mail.HTMLBody ?? ""; } catch { }
                }
                else if (meeting != null)
                {
                    try { body = meeting.Body ?? ""; } catch { }
                    bodyHtml = TryReadComStringProperty(meeting, "HTMLBody");
                }
                else
                {
                    return null;
                }

                if (string.IsNullOrEmpty(folderPath))
                    folderPath = GetOutlookItemFolderPath(item);

                return new MailBodyDto { MailId = mailId, FolderPath = folderPath, Body = body, BodyHtml = bodyHtml };
            }
            catch { return null; }
            finally { if (item != null) try { Marshal.ReleaseComObject(item); } catch { } }
        }

        /// <summary>
        /// Reads attachment metadata for a single mail by EntryID.
        /// AttachmentId is set to Outlook Attachment.Index.ToString() for stable round-trip.
        /// </summary>
        public MailAttachmentsDto ReadMailAttachments(string mailId, string folderPath)
        {
            object item = null;
            try
            {
                item = FindOutlookItemByEntryId(mailId);
                if (item == null) return null;

                if (string.IsNullOrEmpty(folderPath))
                    folderPath = GetOutlookItemFolderPath(item);

                var attachments = new List<MailAttachmentDto>();
                Outlook.Attachments outlookAttachments = null;
                var mail = item as Outlook.MailItem;
                var meeting = item as Outlook.MeetingItem;
                if (mail != null)
                    outlookAttachments = mail.Attachments;
                else if (meeting != null)
                    outlookAttachments = meeting.Attachments;
                else
                    return null;

                if (outlookAttachments != null)
                {
                    for (int i = 1; i <= outlookAttachments.Count; i++)
                    {
                        Outlook.Attachment att = null;
                        try
                        {
                            att = outlookAttachments[i];
                            string fileName = ""; try { fileName = att.FileName ?? ""; } catch { }
                            string displayName = ""; try { displayName = att.DisplayName ?? ""; } catch { }
                            string name = !string.IsNullOrEmpty(fileName) ? fileName : displayName;
                            long size = 0; try { size = att.Size; } catch { }

                            attachments.Add(new MailAttachmentDto
                            {
                                MailId = mailId,
                                AttachmentId = i.ToString(),   // 1-based index as stable id
                                Index = i,
                                Name = name,
                                FileName = fileName,
                                DisplayName = displayName,
                                ContentType = "",              // Outlook OM doesn't expose MIME directly
                                Size = size,
                                IsExported = false,
                                ExportedAttachmentId = "",
                                ExportedPath = ""
                            });
                        }
                        catch { }
                        finally { if (att != null) try { Marshal.ReleaseComObject(att); } catch { } }
                    }
                    try { Marshal.ReleaseComObject(outlookAttachments); } catch { }
                }

                return new MailAttachmentsDto { MailId = mailId, FolderPath = folderPath, Attachments = attachments };
            }
            catch { return null; }
            finally { if (item != null) try { Marshal.ReleaseComObject(item); } catch { } }
        }

        /// <summary>
        /// Reads the Outlook conversation for a single mail. Requires Windows Outlook COM/VSTO runtime validation.
        /// </summary>
        public MailConversationDto ReadMailConversation(string mailId, string folderPath, int maxCount, bool includeBody)
        {
            Outlook.MailItem mail = null;
            Outlook.Conversation conversation = null;
            try
            {
                mail = FindMailByEntryId(mailId);
                if (mail == null) return null;

                if (maxCount <= 0) maxCount = 100;
                if (maxCount > 300) maxCount = 300;
                if (string.IsNullOrEmpty(folderPath)) folderPath = GetMailFolderPath(mail);

                string conversationId = "";
                try { conversationId = mail.ConversationID ?? ""; } catch { }
                string conversationTopic = "";
                try { conversationTopic = mail.ConversationTopic ?? ""; } catch { }

                var mails = new List<MailItemDto>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (IsConversationEnabled(mail))
                {
                    try { conversation = mail.GetConversation(); } catch { conversation = null; }
                    if (conversation != null)
                    {
                        Outlook.SimpleItems roots = null;
                        try
                        {
                            roots = conversation.GetRootItems();
                            if (roots != null)
                            {
                                foreach (object root in roots)
                                    AddConversationItem(root, conversation, mails, seen, includeBody, maxCount);
                            }
                        }
                        finally { if (roots != null) try { Marshal.ReleaseComObject(roots); } catch { } }
                    }
                }

                if (mails.Count == 0)
                {
                    var single = ReadSingleMailDto(mail, folderPath, includeBody);
                    if (single != null) mails.Add(single);
                }

                mails.Sort((left, right) => left.ReceivedTime.CompareTo(right.ReceivedTime));

                return new MailConversationDto
                {
                    MailId = mailId,
                    FolderPath = folderPath,
                    ConversationId = conversationId,
                    ConversationTopic = conversationTopic,
                    Mails = mails
                };
            }
            catch { return null; }
            finally
            {
                if (conversation != null) try { Marshal.ReleaseComObject(conversation); } catch { }
                if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { }
            }
        }

        private bool IsConversationEnabled(Outlook.MailItem mail)
        {
            Outlook.MAPIFolder folder = null;
            Outlook.Store store = null;
            try
            {
                folder = mail.Parent as Outlook.MAPIFolder;
                if (folder == null) return false;
                store = folder.Store;
                return store != null && store.IsConversationEnabled;
            }
            catch { return false; }
            finally
            {
                if (store != null) try { Marshal.ReleaseComObject(store); } catch { }
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }
        }

        private void AddConversationItem(object item, Outlook.Conversation conversation, List<MailItemDto> mails, HashSet<string> seen, bool includeBody, int maxCount)
        {
            try
            {
                if (mails.Count >= maxCount) return;
                var mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    string entryId = "";
                    try { entryId = mail.EntryID ?? ""; } catch { }
                    if (!string.IsNullOrEmpty(entryId) && !seen.Contains(entryId))
                    {
                        var dto = ReadSingleMailDto(mail, GetMailFolderPath(mail), includeBody);
                        if (dto != null)
                        {
                            mails.Add(dto);
                            seen.Add(entryId);
                        }
                    }
                }

                Outlook.SimpleItems children = null;
                try
                {
                    children = conversation.GetChildren(item);
                    if (children != null)
                    {
                        foreach (object child in children)
                            AddConversationItem(child, conversation, mails, seen, includeBody, maxCount);
                    }
                }
                finally { if (children != null) try { Marshal.ReleaseComObject(children); } catch { } }
            }
            finally
            {
                if (item != null) try { Marshal.ReleaseComObject(item); } catch { }
            }
        }

        private static string GetMailFolderPath(Outlook.MailItem mail)
        {
            Outlook.MAPIFolder parent = null;
            try
            {
                parent = mail.Parent as Outlook.MAPIFolder;
                return parent?.FolderPath ?? "";
            }
            catch { return ""; }
            finally { if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { } }
        }

        private static string GetOutlookItemFolderPath(object item)
        {
            Outlook.MAPIFolder parent = null;
            try
            {
                var mail = item as Outlook.MailItem;
                if (mail != null) parent = mail.Parent as Outlook.MAPIFolder;

                var meeting = item as Outlook.MeetingItem;
                if (meeting != null) parent = meeting.Parent as Outlook.MAPIFolder;

                return parent?.FolderPath ?? "";
            }
            catch { return ""; }
            finally { if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { } }
        }

        private object FindOutlookItemByEntryId(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return null;
            try
            {
                return this.Application.Session.GetItemFromID(entryId);
            }
            catch { return null; }
        }

        private static string TryReadComStringProperty(object item, string propertyName)
        {
            if (item == null || string.IsNullOrEmpty(propertyName)) return "";
            try
            {
                return Convert.ToString(item.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    item,
                    null)) ?? "";
            }
            catch { return ""; }
        }

        /// <summary>
        /// Exports a single attachment to the Hub attachment root directory.
        /// Resolves the Outlook 1-based index from: req.Index > req.AttachmentId (parsed) > req.AttachmentIndex (legacy).
        /// Uses req.ExportRootPath when provided by Hub, otherwise falls back to AppData.
        /// </summary>
        public ExportedMailAttachmentDto ExportMailAttachment(OutlookCommandExportMailAttachmentRequest req)
        {
            object item = null;
            try
            {
                item = FindOutlookItemByEntryId(req.MailId);
                if (item == null) return null;

                Outlook.Attachments outlookAttachments = null;
                var mail = item as Outlook.MailItem;
                var meeting = item as Outlook.MeetingItem;
                if (mail != null)
                    outlookAttachments = mail.Attachments;
                else if (meeting != null)
                    outlookAttachments = meeting.Attachments;
                else
                    return null;

                if (outlookAttachments == null || outlookAttachments.Count == 0)
                {
                    if (outlookAttachments != null) try { Marshal.ReleaseComObject(outlookAttachments); } catch { }
                    return null;
                }

                // Resolve 1-based Outlook attachment index
                int resolvedIndex = 0;
                if (req.Index >= 1)
                {
                    resolvedIndex = req.Index;
                }
                else if (!string.IsNullOrEmpty(req.AttachmentId) && int.TryParse(req.AttachmentId, out int parsedId) && parsedId >= 1)
                {
                    resolvedIndex = parsedId;
                }

                if (resolvedIndex < 1 || resolvedIndex > outlookAttachments.Count)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"ExportMailAttachment: invalid index {resolvedIndex}, count={outlookAttachments.Count}, AttachmentId={req.AttachmentId}, Index={req.Index}");
                    try { Marshal.ReleaseComObject(outlookAttachments); } catch { }
                    return null;
                }

                Outlook.Attachment att = null;
                try
                {
                    att = outlookAttachments[resolvedIndex];
                    string fileName = ""; try { fileName = att.FileName ?? ""; } catch { }
                    string displayName = ""; try { displayName = att.DisplayName ?? ""; } catch { }
                    string name = !string.IsNullOrEmpty(fileName) ? fileName
                        : !string.IsNullOrEmpty(displayName) ? displayName
                        : !string.IsNullOrEmpty(req.Name) ? req.Name
                        : !string.IsNullOrEmpty(req.FileName) ? req.FileName
                        : "attachment";

                    // Sanitize filename for filesystem
                    string safeFileName = name;
                    foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                        safeFileName = safeFileName.Replace(c, '_');
                    if (string.IsNullOrEmpty(safeFileName)) safeFileName = "attachment";

                    // Determine export root: prefer Hub-provided path, else AppData
                    string attachmentRoot;
                    if (!string.IsNullOrEmpty(req.ExportRootPath))
                    {
                        attachmentRoot = req.ExportRootPath;
                    }
                    else
                    {
                        attachmentRoot = System.IO.Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "SmartOffice", "Attachments");
                    }
                    System.IO.Directory.CreateDirectory(attachmentRoot);

                    string subDir = System.IO.Path.Combine(attachmentRoot, Guid.NewGuid().ToString("N").Substring(0, 8));
                    System.IO.Directory.CreateDirectory(subDir);
                    string exportPath = System.IO.Path.Combine(subDir, safeFileName);
                    att.SaveAsFile(exportPath);

                    long exportedSize = 0;
                    try { exportedSize = new System.IO.FileInfo(exportPath).Length; } catch { try { exportedSize = att.Size; } catch { } }

                    string folderPath = req.FolderPath ?? "";
                    if (string.IsNullOrEmpty(folderPath))
                        folderPath = GetOutlookItemFolderPath(item);

                    return new ExportedMailAttachmentDto
                    {
                        MailId = req.MailId,
                        FolderPath = folderPath,
                        AttachmentId = req.AttachmentId ?? resolvedIndex.ToString(),
                        ExportedAttachmentId = "",
                        Name = name,
                        FileName = fileName,
                        DisplayName = displayName,
                        ContentType = "",
                        Size = exportedSize,
                        ExportedPath = exportPath,
                        ExportedAt = DateTime.UtcNow
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("ExportMailAttachment inner error: " + ex.Message);
                    return null;
                }
                finally
                {
                    if (att != null) try { Marshal.ReleaseComObject(att); } catch { }
                    try { Marshal.ReleaseComObject(outlookAttachments); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ExportMailAttachment error: " + ex.Message);
                return null;
            }
            finally
            {
                if (item != null) try { Marshal.ReleaseComObject(item); } catch { }
            }
        }
    }
}
