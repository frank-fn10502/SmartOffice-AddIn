using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.Application
{
    internal interface IOutlookCommandAutomation
    {
        bool IsOutlookReady();
        bool TryReadMailsFast(FetchMailsRequest request, out List<MailItemDto> mails, out string error);
        MailBodyDto ReadMailBody(string mailId, string folderPath);
        MailAttachmentsDto ReadMailAttachments(string mailId, string folderPath);
        MailConversationDto ReadMailConversation(string mailId, string folderPath, int maxCount, bool includeBody);
        ExportedMailAttachmentDto ExportMailAttachment(OutlookCommandExportMailAttachmentRequest request);
        List<OutlookRuleDto> ReadRules();
        List<OutlookCategoryDto> ReadCategories();
        List<CalendarEventDto> ReadCalendarEvents(DateTime start, DateTime end);
        List<CalendarRoomDto> ReadCalendarRooms();
        List<CalendarEventDto> CreateCalendarEvent(CalendarEventCommandRequest request);
        List<CalendarEventDto> UpdateCalendarEvent(CalendarEventCommandRequest request);
        List<CalendarEventDto> DeleteCalendarEvent(CalendarEventCommandRequest request);
        List<AddressBookContactDto> ReadAddressBook(AddressBookSyncRequest request, Action<List<AddressBookContactDto>> publishSnapshot = null);

        Task HandleFetchFolderRootsAsync(OutlookCommand command);
        Task HandleFetchFolderChildrenAsync(OutlookCommand command);
        Task HandleMailSearchSliceAsync(OutlookCommand command);
        Task HandleFolderMailsSliceAsync(OutlookCommand command);
        Task HandleManageRuleAsync(OutlookCommand command);
        Task HandleUpdateMailPropertiesAsync(OutlookCommand command);
        Task HandleMoveMailAsync(OutlookCommand command);
        Task HandleMoveMailsAsync(OutlookCommand command);
        Task HandleDeleteMailAsync(OutlookCommand command);
        Task HandleCreateFolderAsync(OutlookCommand command);
        Task HandleDeleteFolderAsync(OutlookCommand command);
        Task HandleUpsertCategoryAsync(OutlookCommand command);
    }
}
