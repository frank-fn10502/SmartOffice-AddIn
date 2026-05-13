using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAddIn.Clients;
using OutlookAddIn.Infrastructure.Diagnostics;
using OutlookAddIn.Infrastructure.Threading;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.Application
{
    internal sealed class OutlookCommandDispatcher
    {
        private const int FetchMailsMaxCount = 100;

        private readonly SignalRClient _signalRClient;
        private readonly OutlookThreadInvoker _outlookThread;
        private readonly IOutlookCommandAutomation _automation;
        private readonly SemaphoreSlim _commandGate = new SemaphoreSlim(1, 1);

        public OutlookCommandDispatcher(
            SignalRClient signalRClient,
            OutlookThreadInvoker outlookThread,
            IOutlookCommandAutomation automation)
        {
            _signalRClient = signalRClient ?? throw new ArgumentNullException(nameof(signalRClient));
            _outlookThread = outlookThread ?? throw new ArgumentNullException(nameof(outlookThread));
            _automation = automation ?? throw new ArgumentNullException(nameof(automation));
        }

        public async Task OnCommandReceivedAsync(OutlookCommand cmd)
        {
            if (!_commandGate.Wait(0))
            {
                try
                {
                    await _signalRClient.ReportCommandResultAsync(
                        cmd.Id,
                        false,
                        "addin_busy: previous Outlook command is still running").ConfigureAwait(false);
                }
                catch { }
                return;
            }

            try
            {
                await HandleCommandAsync(cmd).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                try
                {
                    await _signalRClient.ReportCommandResultAsync(
                        cmd.Id,
                        false,
                        "Error: " + SensitiveLogSanitizer.Sanitize(ex)).ConfigureAwait(false);
                }
                catch { }
            }
            finally
            {
                _commandGate.Release();
            }
        }

        private async Task HandleCommandAsync(OutlookCommand cmd)
        {
            try { await _signalRClient.ReportLogAsync("info", "Received command: " + cmd.Type).ConfigureAwait(false); } catch { }

            switch (cmd.Type)
            {
                case "ping":
                    await HandlePingAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_folder_roots":
                    await _automation.HandleFetchFolderRootsAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_folder_children":
                    await _automation.HandleFetchFolderChildrenAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_mails":
                    await HandleFetchMailsAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_mail_search_slice":
                    await _automation.HandleMailSearchSliceAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_folder_mails_slice":
                    await _automation.HandleFolderMailsSliceAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_mail_body":
                    await HandleFetchMailBodyAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_mail_attachments":
                    await HandleFetchMailAttachmentsAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_mail_conversation":
                    await HandleFetchMailConversationAsync(cmd).ConfigureAwait(false);
                    break;
                case "export_mail_attachment":
                    await HandleExportMailAttachmentAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_rules":
                    await HandleFetchRulesAsync(cmd).ConfigureAwait(false);
                    break;
                case "manage_rule":
                    await _automation.HandleManageRuleAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_categories":
                    await HandleFetchCategoriesAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_calendar":
                    await HandleFetchCalendarAsync(cmd).ConfigureAwait(false);
                    break;
                case "update_mail_properties":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleUpdateMailPropertiesAsync(cmd)).ConfigureAwait(false);
                    break;
                case "move_mail":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleMoveMailAsync(cmd)).ConfigureAwait(false);
                    break;
                case "move_mails":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleMoveMailsAsync(cmd)).ConfigureAwait(false);
                    break;
                case "delete_mail":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleDeleteMailAsync(cmd)).ConfigureAwait(false);
                    break;
                case "create_folder":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleCreateFolderAsync(cmd)).ConfigureAwait(false);
                    break;
                case "delete_folder":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleDeleteFolderAsync(cmd)).ConfigureAwait(false);
                    break;
                case "upsert_category":
                    await _outlookThread.InvokeLegacyAsyncCommand(() => _automation.HandleUpsertCategoryAsync(cmd)).ConfigureAwait(false);
                    break;
                default:
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "Unknown command type: " + cmd.Type).ConfigureAwait(false);
                    break;
            }
        }

        private async Task HandlePingAsync(OutlookCommand cmd)
        {
            bool outlookReady = await _outlookThread.InvokeAsync(() => _automation.IsOutlookReady()).ConfigureAwait(false);

            await _signalRClient.ReportCommandResultAsync(cmd.Id, outlookReady, outlookReady ? "pong" : "Outlook session not ready").ConfigureAwait(false);
        }

        private async Task HandleFetchMailsAsync(OutlookCommand cmd)
        {
            try
            {
                var mr = cmd.MailsRequest;
                var mailReq = new FetchMailsRequest
                {
                    FolderPath = mr?.FolderPath ?? "",
                    Range = mr?.Range ?? "30d",
                    MaxCount = (mr?.MaxCount > 0 ? mr.MaxCount : 100),
                    ReceivedFrom = mr?.ReceivedFrom,
                    ReceivedTo = mr?.ReceivedTo
                };
                if (mailReq.MaxCount > FetchMailsMaxCount) mailReq.MaxCount = FetchMailsMaxCount;

                string readError = null;
                List<MailItemDto> mails = await _outlookThread.InvokeAsync(() =>
                {
                    List<MailItemDto> readMails;
                    if (!_automation.TryReadMailsFast(mailReq, out readMails, out readError))
                        return null;
                    return readMails;
                }).ConfigureAwait(false);

                if (mails == null && !string.IsNullOrEmpty(readError))
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mails error: " + readError).ConfigureAwait(false);
                    return;
                }

                await _signalRClient.PushMailsAsync(mails ?? new List<MailItemDto>()).ConfigureAwait(false);
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mails completed. Items: " + (mails?.Count ?? 0)).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mails error: " + SensitiveLogSanitizer.Sanitize(ex)).ConfigureAwait(false);
            }
        }

        private async Task HandleFetchMailBodyAsync(OutlookCommand cmd)
        {
            var req = cmd.MailBodyRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_body failed: missing mail id").ConfigureAwait(false);
                return;
            }

            MailBodyDto dto = await _outlookThread.InvokeAsync(() => _automation.ReadMailBody(req.MailId, req.FolderPath)).ConfigureAwait(false);
            if (dto == null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_body failed: mail not found").ConfigureAwait(false);
                return;
            }

            await _signalRClient.PushMailBodyAsync(dto).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mail_body completed.").ConfigureAwait(false);
        }

        private async Task HandleFetchMailAttachmentsAsync(OutlookCommand cmd)
        {
            var req = cmd.MailAttachmentsRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_attachments failed: missing mail id").ConfigureAwait(false);
                return;
            }

            MailAttachmentsDto dto = await _outlookThread.InvokeAsync(() => _automation.ReadMailAttachments(req.MailId, req.FolderPath)).ConfigureAwait(false);
            if (dto == null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_attachments failed: mail not found").ConfigureAwait(false);
                return;
            }

            await _signalRClient.PushMailAttachmentsAsync(dto).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mail_attachments completed. Items: " + (dto.Attachments?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task HandleFetchMailConversationAsync(OutlookCommand cmd)
        {
            var req = cmd.MailConversationRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_conversation failed: missing mail id").ConfigureAwait(false);
                return;
            }

            MailConversationDto dto = await _outlookThread.InvokeAsync(() =>
                _automation.ReadMailConversation(req.MailId, req.FolderPath, req.MaxCount, req.IncludeBody)).ConfigureAwait(false);
            if (dto == null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "fetch_mail_conversation failed: mail not found or conversation unavailable").ConfigureAwait(false);
                return;
            }

            await _signalRClient.PushMailConversationAsync(dto).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_mail_conversation completed. Items: " + (dto.Mails?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task HandleExportMailAttachmentAsync(OutlookCommand cmd)
        {
            var req = cmd.ExportMailAttachmentRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "export_mail_attachment failed: missing mail id").ConfigureAwait(false);
                return;
            }

            ExportedMailAttachmentDto dto = await _outlookThread.InvokeAsync(() => _automation.ExportMailAttachment(req)).ConfigureAwait(false);
            if (dto == null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "export_mail_attachment failed: attachment not found or export error").ConfigureAwait(false);
                return;
            }

            await _signalRClient.PushExportedMailAttachmentAsync(dto).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "export_mail_attachment completed.").ConfigureAwait(false);
        }

        private async Task HandleFetchRulesAsync(OutlookCommand cmd)
        {
            await _signalRClient.ReportLogAsync("info", "fetch_rules: starting").ConfigureAwait(false);
            List<OutlookRuleDto> rules = await _outlookThread.InvokeAsync(() => _automation.ReadRules()).ConfigureAwait(false);
            await _signalRClient.PushRulesAsync(rules ?? new List<OutlookRuleDto>()).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_rules completed. Items: " + (rules?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task HandleFetchCategoriesAsync(OutlookCommand cmd)
        {
            await _signalRClient.ReportLogAsync("info", "fetch_categories: starting").ConfigureAwait(false);
            List<OutlookCategoryDto> categories = await _outlookThread.InvokeAsync(() => _automation.ReadCategories()).ConfigureAwait(false);
            await _signalRClient.PushCategoriesAsync(categories ?? new List<OutlookCategoryDto>()).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_categories completed. Items: " + (categories?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task HandleFetchCalendarAsync(OutlookCommand cmd)
        {
            DateTime calStart;
            DateTime calEnd;
            if (cmd.CalendarRequest != null
                && !string.IsNullOrEmpty(cmd.CalendarRequest.StartDate)
                && !string.IsNullOrEmpty(cmd.CalendarRequest.EndDate))
            {
                if (!DateTime.TryParse(cmd.CalendarRequest.StartDate, out calStart))
                    calStart = DateTime.Now.Date;
                if (!DateTime.TryParse(cmd.CalendarRequest.EndDate, out calEnd))
                    calEnd = calStart.AddMonths(1);
            }
            else
            {
                int days = cmd.CalendarRequest?.DaysForward ?? 14;
                if (days <= 0) days = 14;
                calStart = DateTime.Now.Date;
                calEnd = calStart.AddDays(days);
            }

            await _signalRClient.ReportLogAsync("info", $"fetch_calendar: {calStart:yyyy-MM-dd} to {calEnd:yyyy-MM-dd}").ConfigureAwait(false);
            List<CalendarEventDto> events = await _outlookThread.InvokeAsync(() => _automation.ReadCalendarEvents(calStart, calEnd)).ConfigureAwait(false);
            await _signalRClient.PushCalendarAsync(events ?? new List<CalendarEventDto>()).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_calendar completed. Items: " + (events?.Count ?? 0)).ConfigureAwait(false);
        }
    }
}
