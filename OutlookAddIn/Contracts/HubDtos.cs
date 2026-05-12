using System;
using System.Collections.Generic;

namespace OutlookAddIn.Contracts
{
    public class FetchMailsRequest
    {
        public string FolderPath { get; set; }
        public string Range { get; set; }
        public int MaxCount { get; set; }
        /// <summary>Overrides Range when set; date-only means start of that day.</summary>
        public DateTime? ReceivedFrom { get; set; }
        /// <summary>Overrides Range when set; date-only means end of that day (inclusive).</summary>
        public DateTime? ReceivedTo { get; set; }
    }

    public class FetchCalendarRequest
    {
        public int DaysForward { get; set; }
    }

    public class MailItemDto
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        /// <summary>Structured sender; AddIn uses Outlook resolved address book name.</summary>
        public OutlookRecipientDto Sender { get; set; }
        public List<OutlookRecipientDto> ToRecipients { get; set; }
        public List<OutlookRecipientDto> CcRecipients { get; set; }
        public List<OutlookRecipientDto> BccRecipients { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string Body { get; set; }
        public string BodyHtml { get; set; }
        public string FolderPath { get; set; }
        public string ConversationId { get; set; }
        public string ConversationTopic { get; set; }
        public string ConversationIndex { get; set; }
        public string Categories { get; set; }
        public bool IsRead { get; set; }
        public bool IsMarkedAsTask { get; set; }
        // Mail list header ��ܥΡF�����ɬ� 0�A���� metadata �H fetch_mail_attachments ���ǡC
        public int AttachmentCount { get; set; }
        // Mail list header ��ܥΡA�h�Ӫ���W�٥H ", " �걵�F�קK��J�ɮפ��e�Υ������|�C
        public string AttachmentNames { get; set; }
        public string FlagRequest { get; set; }
        public string FlagInterval { get; set; }
        public DateTime? TaskStartDate { get; set; }
        public DateTime? TaskDueDate { get; set; }
        public DateTime? TaskCompletedDate { get; set; }
        public string Importance { get; set; }
        public string Sensitivity { get; set; }
    }

    public class FolderDto
    {
        public string Name { get; set; }
        /// <summary>Outlook MAPIFolder.EntryID; used by Hub to locate parent folder in subsequent fetch_folder_children commands.</summary>
        public string EntryId { get; set; }
        public string FolderPath { get; set; }
        /// <summary>Parent folder's EntryID; empty for store root.</summary>
        public string ParentEntryId { get; set; }
        public string ParentFolderPath { get; set; }
        public int ItemCount { get; set; }
        public string StoreId { get; set; }
        public bool IsStoreRoot { get; set; }
        /// <summary>Outlook Folder.DefaultItemType cast to int. 0 = olMailItem; -1 = unknown / store root.</summary>
        public int DefaultItemType { get; set; }
        /// <summary>From MAPI PR_ATTR_HIDDEN (0x10F4000B). Must not be inferred from folder name.</summary>
        public bool IsHidden { get; set; }
        /// <summary>From MAPI PR_ATTR_SYSTEM (0x10F5000B). Must not be inferred from folder name.</summary>
        public bool IsSystem { get; set; }
        /// <summary>OutlookFolderType enum string; e.g. "Inbox", "Sent", "StoreRoot". See Hub OutlookFolderType enum for all values.</summary>
        public string FolderType { get; set; }
        /// <summary>Whether this folder may have direct children.</summary>
        public bool HasChildren { get; set; }
        /// <summary>Whether Hub has already loaded direct children via fetch_folder_children.</summary>
        public bool ChildrenLoaded { get; set; }
        /// <summary>Expected values: "partial", "loaded", "failed".</summary>
        public string DiscoveryState { get; set; }
    }

    public class OutlookStoreDto
    {
        public string StoreId { get; set; }
        public string DisplayName { get; set; }
        public string StoreKind { get; set; }
        public string StoreFilePath { get; set; }
        public string RootFolderPath { get; set; }
    }

    /// <summary>
    /// Structured Outlook recipient / address entry.
    /// Mirrors OutlookRecipientDto in SmartOffice.Hub.
    /// </summary>
    public class OutlookRecipientDto
    {
        /// <summary>sender / to / cc / bcc / organizer / required / member</summary>
        public string RecipientKind { get; set; }
        public string DisplayName { get; set; }
        public string SmtpAddress { get; set; }
        /// <summary>Raw Outlook address; may be Exchange legacyDN (/O=.../CN=...).</summary>
        public string RawAddress { get; set; }
        /// <summary>Outlook addressType; common values: "SMTP", "EX".</summary>
        public string AddressType { get; set; }
        /// <summary>Outlook AddressEntryUserType name, e.g. olExchangeUserAddressEntry.</summary>
        public string EntryUserType { get; set; }
        /// <summary>True for distribution lists / groups.</summary>
        public bool IsGroup { get; set; }
        public bool IsResolved { get; set; }
        /// <summary>Expanded members for groups; empty when not expanded or not a group.</summary>
        public List<OutlookRecipientDto> Members { get; set; }
    }

    public class FolderSyncBeginDto
    {
        public string SyncId { get; set; }
    }

    public class FolderSyncBatchDto
    {
        public string SyncId { get; set; }
        public int Sequence { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        public List<OutlookStoreDto> Stores { get; set; }
        public List<FolderDto> Folders { get; set; }
    }

    public class FolderSyncCompleteDto
    {
        public string SyncId { get; set; }
    }

    public class ChatMessageDto
    {
        public string Id { get; set; }
        public string Source { get; set; }
        public string Text { get; set; }
        public DateTime Timestamp { get; set; }
    }

    public class OutlookRuleDto
    {
        public string StoreId { get; set; }
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public int ExecutionOrder { get; set; }
        public string RuleType { get; set; }
        public bool IsLocalRule { get; set; }
        public List<string> Conditions { get; set; }
        public List<string> Actions { get; set; }
        public List<string> Exceptions { get; set; }
        /// <summary>
        /// False when the rule contains conditions/actions that the Outlook Rules object model
        /// cannot programmatically create or modify. AddIn must NOT attempt full definition
        /// changes on such rules.
        /// </summary>
        public bool CanModifyDefinition { get; set; }
    }

    public class CalendarEventDto
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Location { get; set; }
        /// <summary>Structured organizer; recipientKind = "organizer".</summary>
        public OutlookRecipientDto Organizer { get; set; }
        /// <summary>Structured attendees; recipientKind = "required".</summary>
        public List<OutlookRecipientDto> RequiredAttendees { get; set; }
        public bool IsRecurring { get; set; }
        public string BusyStatus { get; set; }
    }

    public class MailBodyDto
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public string Body { get; set; }
        public string BodyHtml { get; set; }
    }

    public class MailAttachmentDto
    {
        /// <summary>Must equal the request mailId for round-trip.</summary>
        public string MailId { get; set; }
        /// <summary>Stable attachment id; use Outlook Attachment.Index.ToString() (1-based).</summary>
        public string AttachmentId { get; set; }
        /// <summary>Outlook Attachment.Index (1-based).</summary>
        public int Index { get; set; }
        /// <summary>Web UI display name; prefer FileName, fall back to DisplayName.</summary>
        public string Name { get; set; }
        /// <summary>Outlook Attachment.FileName</summary>
        public string FileName { get; set; }
        /// <summary>Outlook Attachment.DisplayName</summary>
        public string DisplayName { get; set; }
        /// <summary>MIME type; may be empty when Outlook doesn't expose it directly.</summary>
        public string ContentType { get; set; }
        public long Size { get; set; }
        public bool IsExported { get; set; }
        public string ExportedAttachmentId { get; set; }
        public string ExportedPath { get; set; }
    }

    public class MailAttachmentsDto
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public List<MailAttachmentDto> Attachments { get; set; }
    }

    public class MailConversationDto
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public string ConversationId { get; set; }
        public string ConversationTopic { get; set; }
        public List<MailItemDto> Mails { get; set; }
    }

    public class ExportedMailAttachmentDto
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        /// <summary>Must equal the request attachmentId.</summary>
        public string AttachmentId { get; set; }
        /// <summary>AddIn-generated id; Hub fills a GUID when empty.</summary>
        public string ExportedAttachmentId { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public string DisplayName { get; set; }
        public string ContentType { get; set; }
        public long Size { get; set; }
        /// <summary>Full local path output by SaveAsFile; Hub uses this to open the file.</summary>
        public string ExportedPath { get; set; }
        public DateTime ExportedAt { get; set; }
    }

    /// <summary>
    /// Payload for BeginMailSearch and PushMailSearchSliceResult.
    /// Mirrors the slice-based search contract.
    /// </summary>
    public class MailSearchSliceResultDto
    {
        public string SearchId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        public int Sequence { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        /// <summary>Always true when AddIn finishes processing a single folder slice.</summary>
        public bool IsSliceComplete { get; set; }
        public List<MailItemDto> Mails { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Payload for CompleteMailSearchSlice.
    /// </summary>
    public class MailSearchCompleteDto
    {
        public string SearchId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Payload for BeginFolderMails and PushFolderMailsSliceResult.
    /// Mirrors the slice-based folder mails contract.
    /// </summary>
    public class FolderMailsSliceResultDto
    {
        public string FolderMailsId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        public int Sequence { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        public bool Reset { get; set; }
        public bool IsFinal { get; set; }
        /// <summary>True when AddIn finishes processing all items for this single folder slice.</summary>
        public bool IsSliceComplete { get; set; }
        public List<MailItemDto> Mails { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Payload for CompleteFolderMailsSlice.
    /// </summary>
    public class FolderMailsCompleteDto
    {
        public string FolderMailsId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
    }
}
