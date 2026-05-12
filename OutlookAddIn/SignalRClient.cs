using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR.Client;

namespace OutlookAddIn
{
    /// <summary>
    /// SignalR client connecting to /hub/outlook-addin on SmartOffice.Hub.
    /// </summary>
    public class SignalRClient : IDisposable
    {
        private HubConnection _connection;
        private readonly string _hubUrl;
        private bool _disposed;

        public event Func<OutlookCommand, Task> CommandReceived;

        public bool IsConnected => _connection?.State == HubConnectionState.Connected;

        public SignalRClient(string baseUrl)
        {
            _hubUrl = baseUrl.TrimEnd('/') + "/hub/outlook-addin";
        }

        public async Task StartAsync(CancellationToken ct = default)
        {
            if (_connection != null) return;

            _connection = new HubConnectionBuilder()
                .WithUrl(_hubUrl)
                .WithAutomaticReconnect(new[] { TimeSpan.Zero, TimeSpan.FromSeconds(2), TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(10), TimeSpan.FromSeconds(30) })
                .Build();

            _connection.On<OutlookCommand>("OutlookCommand", async cmd =>
            {
                if (CommandReceived != null)
                    await CommandReceived.Invoke(cmd);
            });

            _connection.Reconnected += async (connectionId) =>
            {
                await RegisterAsync(ct);
            };

            await _connection.StartAsync(ct).ConfigureAwait(false);
            await RegisterAsync(ct);
        }

        private async Task RegisterAsync(CancellationToken ct = default)
        {
            try
            {
                await _connection.InvokeAsync("RegisterOutlookAddin", new
                {
                    ClientName = "Outlook VSTO AddIn",
                    Workstation = Environment.MachineName,
                    Version = "1.0.0"
                }, ct).ConfigureAwait(false);
            }
            catch { }
        }

        public async Task StopAsync()
        {
            if (_connection != null)
            {
                try { await _connection.StopAsync().ConfigureAwait(false); } catch { }
                try { await _connection.DisposeAsync().ConfigureAwait(false); } catch { }
                _connection = null;
            }
        }

        // --- Push methods ---

        public async Task BeginFolderSyncAsync(FolderSyncBeginDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("BeginFolderSync", dto).ConfigureAwait(false);
        }

        public async Task PushFolderBatchAsync(FolderSyncBatchDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushFolderBatch", dto).ConfigureAwait(false);
        }

        public async Task CompleteFolderSyncAsync(FolderSyncCompleteDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("CompleteFolderSync", dto).ConfigureAwait(false);
        }

        public async Task PushMailsAsync(List<MailItemDto> mails)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMails", mails).ConfigureAwait(false);
        }

        public async Task PushMailAsync(MailItemDto mail)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMail", mail).ConfigureAwait(false);
        }

        public async Task BeginMailSearchAsync(MailSearchSliceResultDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("BeginMailSearch", dto).ConfigureAwait(false);
        }

        public async Task PushMailSearchSliceResultAsync(MailSearchSliceResultDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMailSearchSliceResult", dto).ConfigureAwait(false);
        }

        public async Task CompleteMailSearchSliceAsync(MailSearchCompleteDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("CompleteMailSearchSlice", dto).ConfigureAwait(false);
        }

        public async Task BeginFolderMailsAsync(FolderMailsSliceResultDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("BeginFolderMails", dto).ConfigureAwait(false);
        }

        public async Task PushFolderMailsSliceResultAsync(FolderMailsSliceResultDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushFolderMailsSliceResult", dto).ConfigureAwait(false);
        }

        public async Task CompleteFolderMailsSliceAsync(FolderMailsCompleteDto dto)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("CompleteFolderMailsSlice", dto).ConfigureAwait(false);
        }

        public async Task PushMailBodyAsync(MailBodyDto body)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMailBody", body).ConfigureAwait(false);
        }

        public async Task PushMailAttachmentsAsync(MailAttachmentsDto attachments)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMailAttachments", attachments).ConfigureAwait(false);
        }

        public async Task PushMailConversationAsync(MailConversationDto conversation)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushMailConversation", conversation).ConfigureAwait(false);
        }

        public async Task PushExportedMailAttachmentAsync(ExportedMailAttachmentDto exported)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushExportedMailAttachment", exported).ConfigureAwait(false);
        }

        public async Task PushRulesAsync(List<OutlookRuleDto> rules)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushRules", rules).ConfigureAwait(false);
        }

        public async Task PushCategoriesAsync(List<OutlookCategoryDto> categories)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushCategories", categories).ConfigureAwait(false);
        }

        public async Task PushCalendarAsync(List<CalendarEventDto> events)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("PushCalendar", events).ConfigureAwait(false);
        }

        public async Task SendChatMessageAsync(ChatMessageDto message)
        {
            if (!IsConnected) return;
            await _connection.InvokeAsync("SendChatMessage", message).ConfigureAwait(false);
        }

        public async Task ReportLogAsync(string level, string message)
        {
            if (!IsConnected) return;
            try
            {
                await _connection.InvokeAsync("ReportAddinLog", new
                {
                    Level = level,
                    Message = message,
                    Timestamp = DateTime.Now
                }).ConfigureAwait(false);
            }
            catch { }
        }

        public async Task ReportCommandResultAsync(string commandId, bool success, string message, string payload = "")
        {
            if (!IsConnected) return;
            try
            {
                await _connection.InvokeAsync("ReportCommandResult", new
                {
                    CommandId = commandId,
                    Success = success,
                    Message = message,
                    Payload = payload,
                    Timestamp = DateTime.Now
                }).ConfigureAwait(false);
            }
            catch { }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { StopAsync().GetAwaiter().GetResult(); } catch { }
        }
    }

    /// <summary>
    /// Mirrors PendingCommand from SmartOffice.Hub.
    /// </summary>
    public class OutlookCommand
    {
        public string Id { get; set; }
        public string Type { get; set; }
        public OutlookCommandFolderDiscoveryRequest FolderDiscoveryRequest { get; set; }
        public OutlookCommandMailsRequest MailsRequest { get; set; }
        public OutlookCommandMailSearchSliceRequest MailSearchSliceRequest { get; set; }
        public OutlookCommandMailBodyRequest MailBodyRequest { get; set; }
        public OutlookCommandMailAttachmentsRequest MailAttachmentsRequest { get; set; }
        public OutlookCommandMailConversationRequest MailConversationRequest { get; set; }
        public OutlookCommandExportMailAttachmentRequest ExportMailAttachmentRequest { get; set; }
        public OutlookCommandCalendarRequest CalendarRequest { get; set; }
        public OutlookCommandMailPropertiesRequest MailPropertiesRequest { get; set; }
        public OutlookCommandCategoryRequest CategoryRequest { get; set; }
        public OutlookCommandCreateFolderRequest CreateFolderRequest { get; set; }
        public OutlookCommandDeleteFolderRequest DeleteFolderRequest { get; set; }
        public OutlookCommandMoveMailRequest MoveMailRequest { get; set; }
        public OutlookCommandMoveMailsRequest MoveMailsRequest { get; set; }
        public OutlookCommandDeleteMailRequest DeleteMailRequest { get; set; }
        public OutlookCommandFolderMailsSliceRequest FolderMailsSliceRequest { get; set; }
        public OutlookCommandRuleRequest RuleRequest { get; set; }
    }

    public class OutlookCommandMailsRequest
    {
        public string FolderPath { get; set; }
        public string Range { get; set; }
        public int MaxCount { get; set; }
        /// <summary>Takes priority over Range when set.</summary>
        public DateTime? ReceivedFrom { get; set; }
        /// <summary>Takes priority over Range when set.</summary>
        public DateTime? ReceivedTo { get; set; }
    }

    public class OutlookCommandCalendarRequest
    {
        public int DaysForward { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
    }

    public class OutlookCommandMailPropertiesRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public bool? IsRead { get; set; }
        public string FlagInterval { get; set; }
        public string FlagRequest { get; set; }
        public DateTime? TaskStartDate { get; set; }
        public DateTime? TaskDueDate { get; set; }
        public DateTime? TaskCompletedDate { get; set; }
        public List<string> Categories { get; set; }
        public List<OutlookCommandNewCategory> NewCategories { get; set; }
    }

    public class OutlookCommandNewCategory
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; }
    }

    public class OutlookCommandCategoryRequest
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; }
    }

    public class OutlookCommandCreateFolderRequest
    {
        public string ParentFolderPath { get; set; }
        public string Name { get; set; }
    }

    public class OutlookCommandDeleteFolderRequest
    {
        public string FolderPath { get; set; }
    }

    public class OutlookCommandMoveMailRequest
    {
        public string MailId { get; set; }
        public string SourceFolderPath { get; set; }
        public string DestinationFolderPath { get; set; }
    }

    public class OutlookCommandDeleteMailRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
    }

    public class OutlookCommandMailBodyRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
    }

    public class OutlookCommandMailAttachmentsRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
    }

    public class OutlookCommandMailConversationRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public int MaxCount { get; set; }
        public bool IncludeBody { get; set; }
    }

    public class OutlookCommandExportMailAttachmentRequest
    {
        public string MailId { get; set; }
        public string FolderPath { get; set; }
        public string AttachmentId { get; set; }
        public int Index { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public string DisplayName { get; set; }
        public string ExportRootPath { get; set; }
    }

    /// <summary>
    /// Slice-based mail search request dispatched by Hub.
    /// AddIn must only read a single folder; all filtering is done via Outlook DASL.
    /// </summary>
    public class OutlookCommandMailSearchSliceRequest
    {
        public string SearchId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        public string StoreId { get; set; }
        /// <summary>Outlook MAPIFolder.EntryID; AddIn uses this as primary folder identity. Must be non-empty.</summary>
        public string FolderEntryId { get; set; }
        /// <summary>Fallback / display path; used only when FolderEntryId cannot be resolved.</summary>
        public string FolderPath { get; set; }
        /// <summary>Text keyword; empty means keyword-less (filter-only) search.</summary>
        public string Keyword { get; set; }
        /// <summary>Fields to search for the keyword: "subject", "sender", "body". Default: ["subject"].</summary>
        public List<string> TextFields { get; set; }
        /// <summary>Category filter; any matching category qualifies the mail.</summary>
        public List<string> CategoryNames { get; set; }
        /// <summary>null = no filter; true = has attachments; false = no attachments.</summary>
        public bool? HasAttachments { get; set; }
        /// <summary>"any" | "flagged" | "unflagged"</summary>
        public string FlagState { get; set; }
        /// <summary>"any" | "unread" | "read"</summary>
        public string ReadState { get; set; }
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        /// <summary>
        /// Number of mails per PushMailSearchSliceResult batch.
        /// AddIn clamps to 3–5; default 5.
        /// </summary>
        public int ResultBatchSize { get; set; }
        public bool ResetSearchResults { get; set; }
        public bool CompleteSearchOnSlice { get; set; }
        /// <summary>
        /// "items_filter": use folder Items/Items.Restrict metadata filter only (no AdvancedSearch).
        /// "outlook_search": use Outlook AdvancedSearch for content (body keyword) search.
        /// Default: "outlook_search" for backward compatibility.
        /// </summary>
        public string ExecutionMode { get; set; }
    }

    public class OutlookCommandFolderDiscoveryRequest
    {
        public string SyncId { get; set; }
        /// <summary>Only relevant for fetch_folder_children; empty for fetch_folder_roots.</summary>
        public string StoreId { get; set; }
        /// <summary>Outlook MAPIFolder.EntryID of the parent; used for fetch_folder_children. Empty for roots.</summary>
        public string ParentEntryId { get; set; }
        /// <summary>Fallback parent path when ParentEntryId is empty.</summary>
        public string ParentFolderPath { get; set; }
        public int MaxDepth { get; set; }
        public int MaxChildren { get; set; }
        public bool Reset { get; set; }
    }

    public class OutlookCommandMoveMailsRequest
    {
        /// <summary>Ordered list of mail EntryIDs to move. Max 500 per call.</summary>
        public List<string> MailIds { get; set; }
        /// <summary>Primary source folder path (single-source shorthand).</summary>
        public string SourceFolderPath { get; set; }
        /// <summary>All source folder paths (multi-source batch move from search results).</summary>
        public List<string> SourceFolderPaths { get; set; }
        public string DestinationFolderPath { get; set; }
        /// <summary>When true, skip individual failures and report stats instead of aborting.</summary>
        public bool ContinueOnError { get; set; }
    }

    public class OutlookCategoryDto
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public int ColorValue { get; set; }
        public string ShortcutKey { get; set; }
    }

    /// <summary>
    /// Slice-based folder mails request dispatched by Hub.
    /// AddIn must only enumerate Items/Items.Restrict on a single folder; must NOT call AdvancedSearch.
    /// </summary>
    public class OutlookCommandFolderMailsSliceRequest
    {
        /// <summary>Folder mails correlation id; must be echoed back in every result DTO.</summary>
        public string FolderMailsId { get; set; }
        public string CommandId { get; set; }
        public string ParentCommandId { get; set; }
        /// <summary>Outlook Store.StoreID; must be non-empty.</summary>
        public string StoreId { get; set; }
        /// <summary>Outlook MAPIFolder.EntryID; primary folder identity. Must be non-empty.</summary>
        public string FolderEntryId { get; set; }
        /// <summary>Display / fallback path; used only when FolderEntryId cannot be resolved.</summary>
        public string FolderPath { get; set; }
        public DateTime? ReceivedFrom { get; set; }
        public DateTime? ReceivedTo { get; set; }
        /// <summary>Maximum mails to return for this folder slice. AddIn clamps to 1-500; default 30.</summary>
        public int MaxCount { get; set; }
        public int SliceIndex { get; set; }
        public int SliceCount { get; set; }
        /// <summary>Number of mails per PushFolderMailsSliceResult batch. AddIn clamps to 3-5; default 5.</summary>
        public int ResultBatchSize { get; set; }
        /// <summary>True only on the first slice of this folder mails request.</summary>
        public bool ResetResults { get; set; }
        /// <summary>True only on the last slice; AddIn should call CompleteFolderMailsSlice after this batch.</summary>
        public bool CompleteOnSlice { get; set; }
    }

    public class OutlookCommandRuleConditions
    {
        public List<string> SubjectContains { get; set; }
        public List<string> BodyContains { get; set; }
        public List<string> SenderAddressContains { get; set; }
        public List<string> Categories { get; set; }
        public bool? HasAttachment { get; set; }
    }

    public class OutlookCommandRuleActions
    {
        public string MoveToFolderPath { get; set; }
        public List<string> AssignCategories { get; set; }
        public bool MarkAsTask { get; set; }
        public bool StopProcessingMoreRules { get; set; }
    }

    /// <summary>
    /// Payload for the manage_rule command.
    /// Only conditions/actions supported by the Outlook Rules object model are handled.
    /// </summary>
    public class OutlookCommandRuleRequest
    {
        /// <summary>"upsert", "delete", or "set_enabled".</summary>
        public string Operation { get; set; }
        public string StoreId { get; set; }
        public string RuleName { get; set; }
        public string OriginalRuleName { get; set; }
        public int? OriginalExecutionOrder { get; set; }
        /// <summary>"receive" or "send".</summary>
        public string RuleType { get; set; }
        public bool Enabled { get; set; }
        public int? ExecutionOrder { get; set; }
        public OutlookCommandRuleConditions Conditions { get; set; }
        public OutlookCommandRuleActions Actions { get; set; }
    }
}
