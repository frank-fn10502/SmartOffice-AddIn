using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR.Client;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.Clients
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
                    Timestamp = DateTime.UtcNow
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
                    Timestamp = DateTime.UtcNow
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

}
