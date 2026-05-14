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
        private const int AddressBookBatchSize = 25;

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
                case "fetch_calendar_rooms":
                    await HandleFetchCalendarRoomsAsync(cmd).ConfigureAwait(false);
                    break;
                case "create_calendar_event":
                    await HandleCalendarMutationAsync(cmd, "create").ConfigureAwait(false);
                    break;
                case "update_calendar_event":
                    await HandleCalendarMutationAsync(cmd, "update").ConfigureAwait(false);
                    break;
                case "delete_calendar_event":
                    await HandleCalendarMutationAsync(cmd, "delete").ConfigureAwait(false);
                    break;
                case "fetch_address_book":
                    await HandleFetchAddressBookAsync(cmd).ConfigureAwait(false);
                    break;
                case "fetch_address_book_group_members":
                    await HandleFetchAddressBookGroupMembersAsync(cmd).ConfigureAwait(false);
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

        private async Task HandleCalendarMutationAsync(OutlookCommand cmd, string operation)
        {
            try
            {
                var request = cmd.CalendarEventRequest ?? new CalendarEventCommandRequest();
                List<CalendarEventDto> events = await _outlookThread.InvokeAsync(() =>
                {
                    if (operation == "create") return _automation.CreateCalendarEvent(request);
                    if (operation == "update") return _automation.UpdateCalendarEvent(request);
                    return _automation.DeleteCalendarEvent(request);
                }).ConfigureAwait(false);
                await _signalRClient.PushCalendarAsync(events ?? new List<CalendarEventDto>()).ConfigureAwait(false);
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, cmd.Type + " completed.").ConfigureAwait(false);
            }
            catch (InvalidOperationException ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, ex.Message).ConfigureAwait(false);
            }
        }

        private async Task HandleFetchCalendarRoomsAsync(OutlookCommand cmd)
        {
            await _signalRClient.ReportLogAsync("info", "fetch_calendar_rooms: starting").ConfigureAwait(false);
            List<CalendarRoomDto> rooms = await _outlookThread.InvokeAsync(() => _automation.ReadCalendarRooms()).ConfigureAwait(false);
            await _signalRClient.PushCalendarRoomsAsync(rooms ?? new List<CalendarRoomDto>()).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_calendar_rooms completed. Items: " + (rooms?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task HandleFetchAddressBookAsync(OutlookCommand cmd)
        {
            var request = new AddressBookSyncRequest
            {
                IncludeOutlookContacts = cmd.AddressBookRequest == null || cmd.AddressBookRequest.IncludeOutlookContacts,
                IncludeAddressLists = cmd.AddressBookRequest == null || cmd.AddressBookRequest.IncludeAddressLists,
                MaxContacts = cmd.AddressBookRequest != null && cmd.AddressBookRequest.MaxContacts > 0 ? cmd.AddressBookRequest.MaxContacts : 0,
                MaxAddressEntriesPerList = cmd.AddressBookRequest != null && cmd.AddressBookRequest.MaxAddressEntriesPerList > 0 ? cmd.AddressBookRequest.MaxAddressEntriesPerList : 0,
                MaxGroupMembers = cmd.AddressBookRequest != null && cmd.AddressBookRequest.MaxGroupMembers >= 0 ? cmd.AddressBookRequest.MaxGroupMembers : 0,
                MaxGroupDepth = cmd.AddressBookRequest != null && cmd.AddressBookRequest.MaxGroupDepth >= 0 ? cmd.AddressBookRequest.MaxGroupDepth : 1
            };

            await _signalRClient.ReportLogAsync("info", "fetch_address_book: starting").ConfigureAwait(false);
            var batchId = Guid.NewGuid().ToString("N");
            var sequence = 1;
            var sentPartialKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var partialReset = true;
            var pushChain = Task.CompletedTask;
            Action<List<AddressBookContactDto>> publishSnapshot = snapshot =>
            {
                var safeSnapshot = snapshot ?? new List<AddressBookContactDto>();
                var newContacts = NewAddressBookContacts(safeSnapshot, sentPartialKeys);
                if (newContacts.Count == 0) return;
                var batches = BuildAddressBookBatches(newContacts, batchId, ref sequence, partialReset, false, safeSnapshot.Count);
                partialReset = false;
                pushChain = pushChain.ContinueWith(_ => PushAddressBookBatchesAsync(batches, true), TaskScheduler.Default).Unwrap();
            };
            List<AddressBookContactDto> contacts = null;
            await _outlookThread.InvokeLegacyAsyncCommand(async () =>
            {
                contacts = await _automation.ReadAddressBookAsync(request, publishSnapshot).ConfigureAwait(true);
            }).ConfigureAwait(false);
            await pushChain.ConfigureAwait(false);
            if (!partialReset && contacts != null)
            {
                var finalContacts = NewAddressBookContacts(contacts, sentPartialKeys);
                var finalBatches = BuildAddressBookBatches(finalContacts, batchId, ref sequence, false, true, contacts.Count);
                await PushAddressBookBatchesAsync(finalBatches, true).ConfigureAwait(false);
                if (finalBatches.Count == 0)
                    await PushAddressBookFinalMarkerAsync(batchId, sequence, contacts.Count).ConfigureAwait(false);
            }
            if (partialReset)
                await PushEmptyAddressBookSnapshotAsync(batchId, sequence).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_address_book completed. Items: " + (contacts?.Count ?? 0)).ConfigureAwait(false);
        }

        private async Task PushEmptyAddressBookSnapshotAsync(string batchId, int sequence)
        {
            await PushAddressBookBatchesAsync(
                new List<AddressBookBatchDto>
                {
                    new AddressBookBatchDto
                    {
                        BatchId = batchId,
                        Sequence = sequence,
                        Reset = true,
                        IsFinal = true,
                        TotalCount = 0,
                        Contacts = new List<AddressBookContactDto>()
                    }
                },
                true).ConfigureAwait(false);
        }

        private async Task PushAddressBookFinalMarkerAsync(string batchId, int sequence, int totalCount)
        {
            await PushAddressBookBatchesAsync(
                new List<AddressBookBatchDto>
                {
                    new AddressBookBatchDto
                    {
                        BatchId = batchId,
                        Sequence = sequence,
                        Reset = false,
                        IsFinal = true,
                        TotalCount = totalCount,
                        Contacts = new List<AddressBookContactDto>()
                    }
                },
                true).ConfigureAwait(false);
        }

        private async Task HandleFetchAddressBookGroupMembersAsync(OutlookCommand cmd)
        {
            var request = new AddressBookGroupMembersRequest
            {
                GroupId = cmd.AddressBookGroupMembersRequest?.GroupId ?? string.Empty,
                GroupSmtpAddress = cmd.AddressBookGroupMembersRequest?.GroupSmtpAddress ?? string.Empty,
                MaxMembers = cmd.AddressBookGroupMembersRequest != null && cmd.AddressBookGroupMembersRequest.MaxMembers > 0
                    ? cmd.AddressBookGroupMembersRequest.MaxMembers
                    : 0,
                ForceRefresh = cmd.AddressBookGroupMembersRequest?.ForceRefresh ?? false,
            };
            await _signalRClient.ReportLogAsync("info", "fetch_address_book_group_members: starting").ConfigureAwait(false);

            List<AddressBookContactDto> members = null;
            await _outlookThread.InvokeLegacyAsyncCommand(async () =>
            {
                members = await _automation.ReadAddressBookGroupMembersAsync(request).ConfigureAwait(true);
            }).ConfigureAwait(false);

            var batch = new AddressBookGroupMembersBatchDto
            {
                GroupId = request.GroupId,
                GroupSmtpAddress = request.GroupSmtpAddress,
                BatchId = Guid.NewGuid().ToString("N"),
                Sequence = 1,
                Reset = true,
                IsFinal = true,
                TotalCount = members?.Count ?? 0,
                Members = members ?? new List<AddressBookContactDto>(),
            };
            await _signalRClient.PushAddressBookGroupMembersBatchAsync(batch).ConfigureAwait(false);
            await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "fetch_address_book_group_members completed. Items: " + batch.Members.Count).ConfigureAwait(false);
        }

        private async Task PushAddressBookBatchesAsync(List<AddressBookBatchDto> batches, bool throwOnError)
        {
            foreach (var batch in batches)
            {
                try
                {
                    await _signalRClient.PushAddressBookBatchAsync(batch).ConfigureAwait(false);
                    await _signalRClient.ReportLogAsync(
                        "info",
                        "fetch_address_book: batch " + batch.Sequence + " items: " + batch.Contacts.Count + "/" + batch.TotalCount).ConfigureAwait(false);
                }
                catch
                {
                    if (throwOnError) throw;
                }
            }
        }

        private static List<AddressBookBatchDto> BuildAddressBookBatches(
            List<AddressBookContactDto> contacts,
            string batchId,
            ref int sequence,
            bool resetFirst,
            bool finalSnapshot,
            int totalCount)
        {
            var batches = new List<AddressBookBatchDto>();
            for (var offset = 0; offset < contacts.Count; offset += AddressBookBatchSize)
            {
                var take = Math.Min(AddressBookBatchSize, contacts.Count - offset);
                batches.Add(new AddressBookBatchDto
                {
                    BatchId = batchId,
                    Sequence = sequence++,
                    Reset = resetFirst && offset == 0,
                    IsFinal = finalSnapshot && offset + take >= contacts.Count,
                    TotalCount = totalCount,
                    Contacts = contacts.GetRange(offset, take)
                });
            }
            return batches;
        }

        private static List<AddressBookContactDto> NewAddressBookContacts(List<AddressBookContactDto> snapshot, HashSet<string> sentKeys)
        {
            var contacts = new List<AddressBookContactDto>();
            foreach (var contact in snapshot)
            {
                var key = AddressBookContactKey(contact);
                if (string.IsNullOrWhiteSpace(key) || !sentKeys.Add(key)) continue;
                contacts.Add(contact);
            }
            return contacts;
        }

        private static string AddressBookContactKey(AddressBookContactDto contact)
        {
            if (!string.IsNullOrWhiteSpace(contact.SmtpAddress)) return contact.SmtpAddress.Trim();
            if (!string.IsNullOrWhiteSpace(contact.RawAddress)) return contact.RawAddress.Trim();
            if (!string.IsNullOrWhiteSpace(contact.DisplayName)) return contact.DisplayName.Trim();
            return contact.Id?.Trim() ?? string.Empty;
        }
    }
}
