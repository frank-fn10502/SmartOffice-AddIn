using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OutlookAddIn.Clients;
using OutlookAddIn.Infrastructure.Threading;
using OutlookAddIn.OutlookServices.Categories;
using OutlookAddIn.OutlookServices.Calendar;
using OutlookAddIn.OutlookServices.Rules;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.Application
{
    internal sealed class ThisAddInAutomationAdapter : IOutlookCommandAutomation
    {
        private readonly ThisAddIn _addin;
        private readonly OutlookCategoryReader _categoryReader;
        private readonly OutlookCalendarReader _calendarReader;
        private readonly OutlookRuleReader _ruleReader;
        private readonly OutlookRuleCommandHandler _ruleCommandHandler;

        public ThisAddInAutomationAdapter(
            ThisAddIn addin,
            SignalRClient signalRClient,
            OutlookThreadInvoker outlookThread)
        {
            _addin = addin ?? throw new ArgumentNullException(nameof(addin));
            _categoryReader = new OutlookCategoryReader(addin.Application);
            _calendarReader = new OutlookCalendarReader(addin.Application);
            _ruleReader = new OutlookRuleReader(addin.Application);
            _ruleCommandHandler = new OutlookRuleCommandHandler(signalRClient, outlookThread, addin.Application);
        }

        public bool IsOutlookReady()
        {
            try
            {
                var session = _addin.Application.Session;
                return session != null && session.Stores != null;
            }
            catch
            {
                return false;
            }
        }

        public bool TryReadMailsFast(FetchMailsRequest request, out List<MailItemDto> mails, out string error)
        {
            return _addin.TryReadMailsFast(request, out mails, out error);
        }

        public MailBodyDto ReadMailBody(string mailId, string folderPath) =>
            _addin.ReadMailBody(mailId, folderPath);

        public MailAttachmentsDto ReadMailAttachments(string mailId, string folderPath) =>
            _addin.ReadMailAttachments(mailId, folderPath);

        public MailConversationDto ReadMailConversation(string mailId, string folderPath, int maxCount, bool includeBody) =>
            _addin.ReadMailConversation(mailId, folderPath, maxCount, includeBody);

        public ExportedMailAttachmentDto ExportMailAttachment(OutlookCommandExportMailAttachmentRequest request) =>
            _addin.ExportMailAttachment(request);

        public List<OutlookRuleDto> ReadRules() => _ruleReader.ReadRules();

        public List<OutlookCategoryDto> ReadCategories() => _categoryReader.ReadCategories();

        public List<CalendarEventDto> ReadCalendarEvents(DateTime start, DateTime end) =>
            _calendarReader.ReadCalendarEvents(start, end);

        public Task HandleFetchFolderRootsAsync(OutlookCommand command) =>
            _addin.HandleFetchFolderRootsAsync(command);

        public Task HandleFetchFolderChildrenAsync(OutlookCommand command) =>
            _addin.HandleFetchFolderChildrenAsync(command);

        public Task HandleMailSearchSliceAsync(OutlookCommand command) =>
            _addin.HandleMailSearchSliceAsync(command);

        public Task HandleFolderMailsSliceAsync(OutlookCommand command) =>
            _addin.HandleFolderMailsSliceAsync(command);

        public Task HandleManageRuleAsync(OutlookCommand command) =>
            _ruleCommandHandler.HandleManageRuleAsync(command);

        public Task HandleUpdateMailPropertiesAsync(OutlookCommand command) =>
            _addin.HandleUpdateMailPropertiesAsync(command);

        public Task HandleMoveMailAsync(OutlookCommand command) =>
            _addin.HandleMoveMailAsync(command);

        public Task HandleMoveMailsAsync(OutlookCommand command) =>
            _addin.HandleMoveMailsAsync(command);

        public Task HandleDeleteMailAsync(OutlookCommand command) =>
            _addin.HandleDeleteMailAsync(command);

        public Task HandleCreateFolderAsync(OutlookCommand command) =>
            _addin.HandleCreateFolderAsync(command);

        public Task HandleDeleteFolderAsync(OutlookCommand command) =>
            _addin.HandleDeleteFolderAsync(command);

        public Task HandleUpsertCategoryAsync(OutlookCommand command) =>
            _addin.HandleUpsertCategoryAsync(command);
    }
}
