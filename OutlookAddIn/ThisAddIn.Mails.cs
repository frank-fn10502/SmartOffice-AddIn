using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
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
                    ReceivedTime = receivedTime == DateTime.MinValue ? DateTime.Now : receivedTime,
                    Body = body,
                    BodyHtml = bodyHtml,
                    FolderPath = folderPath,
                    Categories = categories,
                    IsRead = isRead,
                    IsMarkedAsTask = isMarkedAsTask,
                    AttachmentCount = attachmentCount,
                    AttachmentNames = attachmentNames,
                    FlagRequest = flagRequest,
                    FlagInterval = flagInterval,
                    TaskStartDate = taskStartDate,
                    TaskDueDate = taskDueDate,
                    TaskCompletedDate = taskCompletedDate,
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
            var mails = new List<MailItemDto>();
            try
            {
                int maxCount = req.MaxCount > 0 ? req.MaxCount : 100;
                if (maxCount > 500) maxCount = 500;

                Outlook.MAPIFolder folder = null;
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
                    System.Diagnostics.Debug.WriteLine("ReadMails: folder is null for path: " + req.FolderPath);
                    return mails;
                }

                var items = folder.Items;

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
                        since.ToString("MM/dd/yyyy HH:mm"), until.ToString("MM/dd/yyyy HH:mm"));
                else if (since > DateTime.MinValue)
                    filterExpr = string.Format("[ReceivedTime] >= '{0}'", since.ToString("MM/dd/yyyy HH:mm"));
                else if (until < DateTime.MaxValue)
                    filterExpr = string.Format("[ReceivedTime] <= '{0}'", until.ToString("MM/dd/yyyy HH:mm"));
                else
                    filterExpr = null;

                // Apply Restrict first, then Sort the filtered collection.
                // Sorting before Restrict can cause COM exceptions and the restricted
                // collection does not inherit the sort order.
                Outlook.Items filtered = filterExpr != null ? items.Restrict(filterExpr) : items;
                filtered.Sort("[ReceivedTime]", true);

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                int count = 0;
                foreach (var obj in filtered)
                {
                    if (count >= maxCount) break;

                    var mail = obj as Outlook.MailItem;
                    if (mail == null)
                    {
                        if (obj != null) try { Marshal.ReleaseComObject(obj); } catch { }
                        continue;
                    }

                    try
                    {
                        var dto = ReadSingleMailDto(mail, currentFolderPath, false);
                        if (dto != null)
                        {
                            mails.Add(dto);
                            count++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("ReadMails: failed to convert one mail: " + ex.Message);
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(mail); } catch { }
                    }
                }

                if (filtered != items) try { Marshal.ReleaseComObject(filtered); } catch { }
                try { Marshal.ReleaseComObject(items); } catch { }
                try { Marshal.ReleaseComObject(folder); } catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadMails error: " + ex);
            }
            return mails;
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
            Outlook.MailItem mail = null;
            try
            {
                mail = FindMailByEntryId(mailId);
                if (mail == null) return null;

                string body = ""; try { body = mail.Body ?? ""; } catch { }
                string bodyHtml = ""; try { bodyHtml = mail.HTMLBody ?? ""; } catch { }

                if (string.IsNullOrEmpty(folderPath))
                    try { folderPath = ((Outlook.MAPIFolder)mail.Parent)?.FolderPath ?? ""; } catch { folderPath = ""; }

                return new MailBodyDto { MailId = mailId, FolderPath = folderPath, Body = body, BodyHtml = bodyHtml };
            }
            catch { return null; }
            finally { if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { } }
        }

        /// <summary>
        /// Reads attachment metadata for a single mail by EntryID.
        /// AttachmentId is set to Outlook Attachment.Index.ToString() for stable round-trip.
        /// </summary>
        public MailAttachmentsDto ReadMailAttachments(string mailId, string folderPath)
        {
            Outlook.MailItem mail = null;
            try
            {
                mail = FindMailByEntryId(mailId);
                if (mail == null) return null;

                if (string.IsNullOrEmpty(folderPath))
                    try { folderPath = ((Outlook.MAPIFolder)mail.Parent)?.FolderPath ?? ""; } catch { folderPath = ""; }

                var attachments = new List<MailAttachmentDto>();
                var outlookAttachments = mail.Attachments;
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
            finally { if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { } }
        }

        /// <summary>
        /// Exports a single attachment to the Hub attachment root directory.
        /// Resolves the Outlook 1-based index from: req.Index > req.AttachmentId (parsed) > req.AttachmentIndex (legacy).
        /// Uses req.ExportRootPath when provided by Hub, otherwise falls back to AppData.
        /// </summary>
        public ExportedMailAttachmentDto ExportMailAttachment(OutlookCommandExportMailAttachmentRequest req)
        {
            Outlook.MailItem mail = null;
            try
            {
                mail = FindMailByEntryId(req.MailId);
                if (mail == null) return null;

                var outlookAttachments = mail.Attachments;
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
                        try { folderPath = ((Outlook.MAPIFolder)mail.Parent)?.FolderPath ?? ""; } catch { }

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
                        ExportedAt = DateTime.Now
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
                if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { }
            }
        }
    }
}
