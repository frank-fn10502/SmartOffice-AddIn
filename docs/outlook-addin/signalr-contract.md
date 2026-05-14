# Outlook AddIn SignalR 溝通介面

本文件整理工作機 Outlook AddIn 的正式 SignalR-only contract。AddIn 實作順序與驗收請先看 `features-checklist.md`；本文件只保留 SignalR method、request object、DTO 與 payload 細節。

## 適用範圍

- Outlook AddIn 實作應在工作機完整 SmartOffice solution 中完成。
- AddIn 只負責 Outlook object model / Office automation，並把結果轉成本文件的 DTO。
- AddIn 不負責 request endpoint、快取、跨 folder 搜尋排程、資料 merge 或外部 client workflow。
- 不相容舊版 AddIn channel；不要實作或保留 `/api/outlook/poll`、`/api/outlook/push-*`、HTTP chat 或未列於本文件的 legacy command。
- 不維護未使用功能。若 command、欄位或 handler 沒被本 contract 使用，請刪除或不要新增。
- Mail body、folder name、category name、chat message 都可能含有敏感 business data；測試回報請匿名化。

## 通訊模型

1. Outlook AddIn 連線到 `/hub/outlook-addin`。
2. AddIn invoke `RegisterOutlookAddin(info)` 完成註冊。
3. AddIn 透過 SignalR client event `OutlookCommand` 收到 command。
4. AddIn 執行 Outlook automation。
5. AddIn 透過 SignalR server method `Push*`、`ReportAddinLog` 或 `ReportCommandResult` 回報結果。

正式 AddIn endpoint：

```text
/hub/outlook-addin
```

舊的 `/api/outlook/poll` 與 `/api/outlook/push-*` 不再是 AddIn contract，也不需要任何 fallback。

## AddIn 連線註冊

AddIn 連到 `/hub/outlook-addin` 後，先 invoke：

```text
RegisterOutlookAddin(info)
```

```json
{
  "clientName": "Outlook VSTO AddIn",
  "workstation": "WORKSTATION-01",
  "version": "0.1.0"
}
```

連線註冊完成後，後續 command 會送到這個 connection 所屬的 Outlook AddIn group。

## 時間規範

Hub、Web UI、AddIn 之間所有 JSON / SignalR / HTTP transport 的 date-time 欄位一律使用 UTC，例如 `2026-05-01T01:30:00Z`。Web UI 顯示時可轉成本機時間；AddIn 呼叫 Outlook COM 查詢、寫入 task due date 或 calendar filter 時，才在 Outlook automation 邊界轉成 Outlook 本機時間。AddIn 從 Outlook 讀出的 `ReceivedTime`、calendar `Start` / `End`、task date 與 `timestamp` 類欄位，回推 Hub 前必須轉回 UTC。

## AddIn 接收 Command

AddIn 需要 listen：

```text
OutlookCommand
```

Payload：

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "receivedFrom": "2026-05-01T01:30:00Z",
    "receivedTo": "2026-05-08T01:30:00Z",
    "maxCount": 30
  },
  "folderDiscoveryRequest": null,
  "mailSearchSliceRequest": null,
  "mailBodyRequest": null,
  "mailAttachmentsRequest": null,
  "exportMailAttachmentRequest": null,
  "calendarRequest": null,
  "calendarEventRequest": null,
  "mailPropertiesRequest": null,
  "categoryRequest": null,
  "ruleRequest": null,
  "createFolderRequest": null,
  "deleteFolderRequest": null,
  "moveMailRequest": null,
  "moveMailsRequest": null,
  "deleteMailRequest": null
}
```

上方是 `PendingCommand` 外層 shape 範例；實際 command 只會填入對應 request object，其餘欄位會是 `null` 或未被 AddIn 使用。AddIn 的 command model 應包含表格中列出的所有 request property，避免處理 `fetch_folder_roots`、`fetch_folder_children`、`fetch_mail_attachments`、`export_mail_attachment` 或 `move_mails` 時因 model 漏欄位而無法讀取 request。

目前 command type：

| Command type | 對應 request object | 說明 |
| --- | --- | --- |
| `fetch_folder_roots` | `folderDiscoveryRequest` | 只讀取 Outlook stores 與各 store root folder；不得遞迴 subfolders。 |
| `fetch_folder_children` | `folderDiscoveryRequest` | 只讀取指定 parent folder 的直接 children。 |
| `fetch_mails` | `mailsRequest` | 讀取指定 folder 的 mail metadata，不應包含完整 body |
| `fetch_folder_mails_slice` | `folderMailsSliceRequest` | 直接列出指定單一 folder 的 mail metadata；不得使用 Outlook search |
| `fetch_mail_search_slice` | `mailSearchSliceRequest` | 讀取指定單一 folder 的 mail search slice |
| `fetch_mail_body` | `mailBodyRequest` | 使用者點開單封 mail 後，讀取該 mail body |
| `fetch_mail_attachments` | `mailAttachmentsRequest` | 讀取單封 mail 的附件 metadata |
| `fetch_mail_conversation` | `mailConversationRequest` | 讀取單封 mail 所屬 Outlook conversation；`includeBody=true` 時可回推每封 mail body |
| `export_mail_attachment` | `exportMailAttachmentRequest` | 將指定 attachment 匯出到約定的本機 attachment root |
| `fetch_rules` | 無 | 讀取 Outlook rules |
| `manage_rule` | `ruleRequest` | 新增、刪除、啟用/停用或修改 Outlook rules 中 Rules object model 可建立的條件與動作 |
| `fetch_categories` | 無 | 讀取 Outlook master category list |
| `ping` | 無 | readiness probe；只有 Outlook object model 可正常呼叫時才回成功 |
| `fetch_calendar` | `calendarRequest` | 讀取 calendar events |
| `fetch_calendar_rooms` | 無 | 讀取 Outlook room/resource 清單 |
| `create_calendar_event` | `calendarEventRequest` | 建立 SmartOffice-owned calendar event |
| `update_calendar_event` | `calendarEventRequest` | 更新 SmartOffice-owned calendar event |
| `delete_calendar_event` | `calendarEventRequest` | 刪除 SmartOffice-owned calendar event |
| `fetch_address_book` | `addressBookRequest` | 讀取 Outlook Contacts folder 與 AddressLists / GAL 的通訊錄 metadata |
| `update_mail_properties` | `mailPropertiesRequest` | 一次更新已讀、flag、category 與新 category |
| `upsert_category` | `categoryRequest` | 新增或更新 master category |
| `create_folder` | `createFolderRequest` | 建立 folder |
| `delete_folder` | `deleteFolderRequest` | 將 folder 移到 Outlook default Deleted Items folder |
| `move_mail` | `moveMailRequest` | 移動單封 mail |
| `move_mails` | `moveMailsRequest` | 移動多封 mail |
| `delete_mail` | `deleteMailRequest` | 將單封 mail 移到 Outlook default Deleted Items folder |

`delete_mail` 是獨立 command；但它的唯一允許實作仍是將 mail 移到同一個 Outlook store / mailbox 的 default Deleted Items folder。AddIn 必須用 Outlook object model 的 default folder identity 定位目的地，例如 `Store.GetDefaultFolder(olFolderDeletedItems)` 或等效流程；不得用顯示名稱、本地化名稱、`folderPath` 字串包含 `Deleted Items` / `刪除的郵件` 等方式猜測目的 folder。AddIn 不得呼叫 Outlook `MailItem.Delete()` 或永久刪除郵件。

所有 mail / folder delete 類 command 都是 soft delete。`delete_mail` 只允許把 mail 移到 Outlook default Deleted Items folder，不得永久刪除。`delete_folder` 也只允許把 folder 移到 Outlook default Deleted Items folder；AddIn 收到 command 時永遠只執行 move，不呼叫永久刪除 API。

Calendar mutation 只能作用於 SmartOffice 建立的 event。AddIn 建立 event 時必須寫入 ownership marker，例如 Outlook `UserProperties`；收到 `update_calendar_event` 或 `delete_calendar_event` 時必須重新讀取 Outlook item 並確認 marker 存在。若 marker 不存在，回報 `ReportCommandResult(success=false, message="not_smartoffice_owned")`，不得用 subject、時間或 organizer 猜測 ownership。

Calendar `resources` 是 Outlook meeting room / equipment resource recipients。AddIn 應把每個 resource 加到 `AppointmentItem.Recipients`，並設定 recipient type 為 `olResource`；不要只把會議室名稱寫進 `Location`。

`ping` 不是單純 SignalR echo。收到 `ping` 時，AddIn 必須確認 Outlook object model 可正常呼叫。若 Outlook 剛啟動、profile 尚未 ready、COM object 暫時 busy，AddIn 應回 `ReportCommandResult(success=false)` 或等到可判斷後再回覆；不要在 Outlook 尚不可操作時回成功。

## AddIn 回報結果

AddIn 可 invoke 下列 server method：

| Method | Payload | 用途 |
| --- | --- | --- |
| `BeginFolderSync` | `FolderSyncBeginDto` | 開始 folder 增量同步 |
| `PushFolderBatch` | `FolderSyncBatchDto` | 推送一批 stores / folders |
| `CompleteFolderSync` | `FolderSyncCompleteDto` | 結束 folder 增量同步 |
| `PushMails` | `MailItemDto[]` | 回推目前 mail snapshot；`fetch_mails` 的回傳應只含 metadata，`body` / `bodyHtml` 留空 |
| `PushMail` | `MailItemDto` | 回推同 id 的單封 mail；用於 `update_mail_properties` 這類不應重抓列表的單封 mutation |
| `BeginFolderMails` | `FolderMailsSliceResultDto` | 開始直接列出 folder mails |
| `PushFolderMailsSliceResult` | `FolderMailsSliceResultDto` | 推送直接列出的 folder mail metadata |
| `CompleteFolderMailsSlice` | `FolderMailsCompleteDto` | 結束 folder mails slice |
| `BeginMailSearch` | `MailSearchSliceResultDto` | 開始 mail search slice |
| `PushMailSearchSliceResult` | `MailSearchSliceResultDto` | 推送 Outlook 內建搜尋結果 |
| `CompleteMailSearchSlice` | `MailSearchCompleteDto` | 結束 mail search slice |
| `PushMailBody` | `MailBodyDto` | 回推同 id 的 body；用於 `fetch_mail_body` |
| `PushMailAttachments` | `MailAttachmentsDto` | 回推 attachment metadata；用於 `fetch_mail_attachments` |
| `PushMailConversation` | `MailConversationDto` | 回推單封 mail 所屬討論串；用於 `fetch_mail_conversation` |
| `PushExportedMailAttachment` | `ExportedMailAttachmentDto` | 回推已匯出的 attachment path；用於 `export_mail_attachment` |
| `PushRules` | `OutlookRuleDto[]` | 回推 Outlook rules snapshot |
| `PushCategories` | `OutlookCategoryDto[]` | 回推 Outlook master category snapshot |
| `PushCalendar` | `CalendarEventDto[]` | 回推 calendar events snapshot |
| `PushAddressBook` | `AddressBookContactDto[]` | 回推 Outlook address book metadata snapshot |
| `SendChatMessage` | `ChatMessageDto` | AddIn 透過 SignalR 送出 chat message |
| `ReportAddinLog` | `AddinLogEntry` | 回報診斷 log |
| `ReportCommandResult` | `OutlookCommandResult` | 回報 command 成敗 |

每個 command 完成後，建議至少 invoke `ReportCommandResult`。如果 command 會改變 Outlook snapshot，請同時 invoke 對應 `Push*` method。Folder discovery 只使用 `fetch_folder_roots` 與 `fetch_folder_children`，並透過 `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync` 回推增量結果；AddIn 不得實作一次遞迴整棵樹的 command。Folder mails 只使用 `BeginFolderMails`、`PushFolderMailsSliceResult`、`CompleteFolderMailsSlice`，且不得走 Outlook search。Mail search slice 只使用 `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice`，不要用 `PushMails` 覆蓋目前 folder list。單封屬性更新請使用 `PushMail`，不要為了更新一封 mail 重新 `PushMails`。郵件 body 請只在 `fetch_mail_body` 後以 `PushMailBody` 回推。郵件討論串請只在 `fetch_mail_conversation` 後以 `PushMailConversation` 回推。附件採 `fetch_mail_attachments` 先看 metadata、有需要才 `export_mail_attachment` 匯出到本機路徑；AddIn 不負責開檔。

`fetch_folder_roots` 與 `fetch_folder_children` 不應只回 `ReportCommandResult(success=true)`。AddIn 必須在成功結果前至少完成 `PushFolderBatch` 或 `CompleteFolderSync`；如果 Outlook store 尚未 ready、parent folder 無效或列舉結果異常，請回報 `success=false` 與可診斷訊息。

AddIn 不應使用 HTTP `/api/outlook/chat` 送 chat；請改用 `/hub/outlook-addin` 上的 `SendChatMessage(message)`。

`OutlookCommandResult` sample：

```json
{
  "commandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "success": true,
  "message": "fetch_mails completed",
  "payload": "",
  "timestamp": "2026-05-04T01:30:06Z"
}
```

## Request Object 格式

### FetchMailsRequest

```json
{
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "receivedFrom": "2026-05-01T01:30:00Z",
  "receivedTo": "2026-05-08T01:30:00Z",
  "maxCount": 30
}
```

AddIn 只需要依 `receivedFrom` / `receivedTo` date-time 邊界篩選。

### FetchCalendarRequest

```json
{
  "daysForward": 31,
  "startDate": "2026-05-01",
  "endDate": "2026-06-01"
}
```

`startDate` 含當日，`endDate` 不含當日。`daysForward` 不再作為舊 AddIn fallback；若同時收到 date range 與 `daysForward`，以 `startDate` / `endDate` 為準。

### OutlookRuleCommandRequest

AddIn 只處理 Microsoft Outlook Rules object model 明確支援的 rule 管理面向。AddIn 不得承諾或自行模擬 Rules and Alerts Wizard 中無法 programmatically create 的特殊條件與動作。

`manage_rule` sample：

```json
{
  "operation": "upsert",
  "storeId": "",
  "ruleName": "客戶郵件標記",
  "originalRuleName": "",
  "originalExecutionOrder": null,
  "ruleType": "receive",
  "enabled": true,
  "executionOrder": null,
  "conditions": {
    "subjectContains": ["報價"],
    "bodyContains": [],
    "bodyOrSubjectContains": [],
    "messageHeaderContains": [],
    "senderAddressContains": ["example.com"],
    "recipientAddressContains": [],
    "categories": ["客戶"],
    "hasAttachment": true,
    "importance": "high",
    "toMe": false,
    "toOrCcMe": false,
    "onlyToMe": false,
    "meetingInviteOrUpdate": false
  },
  "actions": {
    "moveToFolderPath": "\\\\主要信箱 - User\\Inbox\\客戶",
    "copyToFolderPath": "",
    "assignCategories": ["客戶"],
    "clearCategories": false,
    "markAsTask": true,
    "markAsTaskInterval": "this_week",
    "delete": false,
    "desktopAlert": true,
    "stopProcessingMoreRules": true
  }
}
```

- `operation`: `upsert`、`delete` 或 `set_enabled`。
- `ruleName`: 新增或更新後的 rule name；Outlook rules collection 的 rule name 不保證唯一，因此更新與刪除時也應帶 `originalExecutionOrder`。
- `originalRuleName` / `originalExecutionOrder`: 更新、刪除或啟用/停用既有 rule 時用於定位原 rule；AddIn 可用 `Rules.Item(index)` 優先定位，必要時用 name fallback。
- `ruleType`: `receive` 或 `send`，對應 Outlook `OlRuleType`。
- `conditions`: 只包含 AddIn 必須支援建立的條件：subject contains、body contains、subject or body contains、message header contains、sender address contains、recipient address contains、category、has attachment、importance、to me、to or cc me、only to me、meeting invite/update。`hasAttachment` 只支援 `true` 或省略；Outlook Rules object model 不提供可建立的「無附件」條件。
- `actions`: 只包含 AddIn 必須支援建立的動作：move to folder、copy to folder、assign categories、clear categories、mark as task、delete、desktop alert、stop processing more rules。`delete` 對應 Outlook rule action，不是 Hub mail delete command；不得擴充成永久刪除 API。

AddIn 實作時應使用 `Store.GetRules()` 取得 rules collection。新增 rule 使用 `Rules.Create`；更新既有 rule 可修改 `Rule.Enabled`、`Rule.Name`、`Rule.ExecutionOrder`、支援的 `Rule.Conditions` 與 `Rule.Actions`；刪除 rule 使用 `Rules.Remove`；任何變更都必須呼叫 `Rules.Save(false)` 或等效流程保存。`Rules.Save` 可能因 Exchange 規則限制、空白條件/動作或使用者同時開啟 Rules and Alerts Wizard 而失敗；失敗時回 `ReportCommandResult(success=false)`，message 不得包含敏感資料。

若既有 rule 含 Rules object model 無法建立的特殊條件或動作，AddIn 仍可列舉並回推 snapshot；但 `OutlookRuleDto.canModifyDefinition` 應回 `false`，表示該 rule 不支援完整修改 definition。

### FolderDiscoveryRequest

`fetch_folder_roots` sample：

```json
{
  "syncId": "folder-sync-001",
  "storeId": "",
  "parentEntryId": "",
  "parentFolderPath": "",
  "maxDepth": 0,
  "maxChildren": 50,
  "reset": true
}
```

`fetch_folder_children` sample：

```json
{
  "syncId": "folder-sync-001",
  "storeId": "[redacted store id]",
  "parentEntryId": "[redacted folder entry id]",
  "parentFolderPath": "\\\\主要信箱 - User\\Inbox",
  "maxDepth": 1,
  "maxChildren": 50,
  "reset": false
}
```

- `fetch_folder_roots` 只允許列出 `Application.Session.Stores`、每個 `Store.GetRootFolder()`，以及必要的 root metadata。
- `fetch_folder_children` 必須使用 `storeId` + `parentEntryId` 優先定位 parent folder；若 `parentEntryId` 空白才可用 `parentFolderPath` fallback。
- `maxDepth` 預設與正式值都是 `1`；AddIn 不得自行遞迴超過 request 指定深度。
- `maxChildren` 是單次 command 上限，AddIn 必須 clamp 到合理值。
- 每個 children command 應回推 parent folder 本身，並將該 parent 的 `childrenLoaded=true`、`discoveryState="loaded"`。

### FolderMailsSliceRequest

`fetch_folder_mails_slice` 是直接用 folder 查 mails 的 command，不是 search。AddIn 不得為此 command 呼叫 `Application.AdvancedSearch`、DASL content search 或全域搜尋。

```json
{
  "folderMailsId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "storeId": "[redacted store id]",
  "folderEntryId": "[redacted folder entry id]",
  "folderPath": "\\\\主要信箱 - User\\Inbox\\folderA",
  "receivedFrom": "2026-04-30T16:00:00Z",
  "receivedTo": "2026-05-04T15:59:59Z",
  "maxCount": 30,
  "sliceIndex": 0,
  "sliceCount": 30,
  "resultBatchSize": 5,
  "resetResults": true,
  "completeOnSlice": false
}
```

- `folderMailsId`: folder mails correlation id；AddIn 回推 result 時必須沿用。
- `commandId`: 此 slice command id；AddIn 回推 result 時必須沿用。
- `parentCommandId`: 原始 `request-folder-mails` 的 command id。
- `storeId`: 指定單一 Outlook Store，必須非空。
- `folderEntryId`: Outlook folder EntryID，必須非空；AddIn 應優先使用 `storeId` + `folderEntryId` 定位 folder。
- `folderPath`: 指定單一 Outlook folder，必須非空；只作為顯示與 `folderEntryId` 無法定位時的 fallback。
- `receivedFrom` / `receivedTo`: 收到時間區段，兩者可獨立使用；transport 必須是 UTC，AddIn 組 Outlook `Items.Restrict` / `GetTable` filter 前再轉 Outlook 本機時間。
- `maxCount`: 此 folder slice 最多回傳幾封 mail；AddIn 應 clamp 到 `1` 到 `500`，預設 `30`。
- `sliceIndex` / `sliceCount`: folder slice 序號與總數，可用於 progress message。
- `resultBatchSize`: 結果每批回推筆數；AddIn 應 clamp 在 `3` 到 `5` 之間，預設 `5`。
- `resetResults`: 只有第一個 slice 是 `true`。
- `completeOnSlice`: 只有最後一個 slice 是 `true`。

AddIn 收到 `fetch_folder_mails_slice` 時，必須定位指定單一 folder，使用該 folder 的 `Items` / `Items.Restrict` 或 `Folder.GetTable` 逐步讀取 mail metadata；若有 `receivedFrom` / `receivedTo`，可用 Outlook filter 限縮時間。AddIn 必須遵守 `maxCount`，避免單一 command 長時間占用 Outlook UI thread。回傳使用 `BeginFolderMails`、`PushFolderMailsSliceResult`、`CompleteFolderMailsSlice`，只回 metadata，`body` / `bodyHtml` 應留空。

Folder mails slice result sample：

```json
{
  "folderMailsId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "sequence": 1,
  "sliceIndex": 0,
  "sliceCount": 30,
  "reset": true,
  "isFinal": false,
  "isSliceComplete": false,
  "mails": [
    {
      "id": "[redacted Outlook EntryID or stable id]",
      "subject": "Sample mail",
      "sender": {
        "recipientKind": "sender",
        "displayName": "Sender",
        "smtpAddress": "sender@example.invalid",
        "rawAddress": "",
        "addressType": "SMTP",
        "entryUserType": "",
        "isGroup": false,
        "isResolved": true,
        "members": []
      },
      "toRecipients": [],
      "ccRecipients": [],
      "bccRecipients": [],
      "receivedTime": "2026-05-04T01:30:00Z",
      "body": "",
      "bodyHtml": "",
      "folderPath": "\\\\主要信箱 - User\\Inbox\\folderA",
      "attachmentCount": 0,
      "attachmentNames": ""
    }
  ],
  "message": ""
}
```

### MailSearchSliceRequest

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "storeId": "[redacted store id]",
  "folderEntryId": "[redacted folder entry id]",
  "folderPath": "\\\\主要信箱 - User\\Inbox",
  "executionMode": "items_filter",
  "keyword": "客戶",
  "textFields": ["subject"],
  "categoryNames": ["Customer"],
  "hasAttachments": true,
  "flagState": "any",
  "readState": "unread",
  "receivedFrom": "2026-04-30T16:00:00Z",
  "receivedTo": "2026-05-04T15:59:59Z",
  "sliceIndex": 0,
  "sliceCount": 30,
  "resultBatchSize": 5,
  "resetSearchResults": true,
  "completeSearchOnSlice": false
}
```

- `searchId`: search correlation id；AddIn 回推 slice result 時必須沿用。
- `commandId`: 此 slice command id；AddIn 回推 slice result 時必須沿用。
- `parentCommandId`: 原始 `request-mail-search` 的 command id。
- `storeId`: 指定單一 Outlook Store，必須非空。
- `folderEntryId`: Outlook folder EntryID，必須非空；AddIn 應優先使用 `storeId` + `folderEntryId` 定位 folder。
- `folderPath`: 指定單一 Outlook folder，必須非空；只作為顯示、search scope 組合與 `folderEntryId` 無法定位時的 fallback，不得取代 `folderEntryId` 作為主要 identity。
- `executionMode`: `items_filter` 或 `outlook_search`。`items_filter` 代表只做 folder `Items` / `Items.Restrict` metadata filter；`outlook_search` 才能使用 Outlook 內容搜尋。
- `keyword`: 文字搜尋關鍵字；空白時只套用篩選條件。
- `textFields`: keyword 文字搜尋欄位；目前正式值為 `subject`、`sender`、`body`。預設只有 `subject`。
- `categoryNames`: 分類篩選；任一分類符合即可。
- `hasAttachments`: 附件篩選；`true` 表示包含附件，`false` 表示不含附件，省略表示不限。
- `flagState`: 旗標篩選；`any`、`flagged` 或 `unflagged`。
- `readState`: 已讀篩選；`any`、`unread` 或 `read`。
- `receivedFrom` / `receivedTo`: 收到時間區段，兩者可獨立使用；transport 必須是 UTC。
- `sliceIndex` / `sliceCount`: folder slice 序號與總數，可用於 progress message。
- `resultBatchSize`: 搜尋結果每批回推筆數；AddIn 應 clamp 在 `3` 到 `5` 之間，預設 `5`。
- `resetSearchResults`: 只有第一個 slice 是 `true`；AddIn 呼叫 `BeginMailSearch` 或第一批 `PushMailSearchSliceResult` 時應沿用。
- `completeSearchOnSlice`: 只有最後一個 slice 是 `true`；AddIn 只有最後一個 slice 才應呼叫 `CompleteMailSearchSlice` 或送 `isFinal=true`。

AddIn 若收到空 `storeId`、空 `folderEntryId` 或空 `folderPath`，應使用 `CompleteMailSearchSlice(success=false)` 結束該 slice，不得自行展開整個 store 或全域搜尋。AddIn 定位 folder 時必須優先使用 `storeId` + `folderEntryId`；只有 `folderEntryId` 無法在目前 Outlook profile 中解析時，才可用 `folderPath` fallback 並回報匿名化 warning。

AddIn 必須在指定單一 folder 內依 `executionMode` 執行。`items_filter` 用 folder `Items` / `Items.Restrict` 或等效逐項 metadata filter 套用 subject、sender、category、attachment、flag、read state 與 received time，不得呼叫 `Application.AdvancedSearch`。`outlook_search` 才能使用 Microsoft Outlook 內容搜尋流程處理 body keyword。這不是 typo-tolerant fuzzy search。AddIn 只回傳 metadata，`body` / `bodyHtml` 應留空。

Mail search slice result sample：

```json
{
  "searchId": "6fb66d3a-7f4f-4a6d-9b3f-7e1e8c2f2d84",
  "commandId": "slice-command-id",
  "parentCommandId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "sequence": 1,
  "sliceIndex": 0,
  "sliceCount": 30,
  "reset": true,
  "isFinal": false,
  "isSliceComplete": false,
  "mails": [
    {
      "id": "[redacted Outlook EntryID or stable id]",
      "subject": "Sample mail",
      "sender": {
        "recipientKind": "sender",
        "displayName": "Sender",
        "smtpAddress": "sender@example.invalid",
        "rawAddress": "/O=ORG/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=sender",
        "addressType": "EX",
        "entryUserType": "olExchangeUserAddressEntry",
        "isGroup": false,
        "isResolved": true,
        "members": []
      },
      "toRecipients": [],
      "ccRecipients": [],
      "bccRecipients": [],
      "receivedTime": "2026-05-04T01:30:00Z",
      "body": "",
      "bodyHtml": "",
      "folderPath": "\\\\主要信箱 - User\\Inbox",
      "attachmentCount": 1,
      "attachmentNames": "sample.pdf"
    }
  ],
  "message": ""
}
```

同一個 folder slice 可能找到大量郵件，AddIn 必須用多次 `PushMailSearchSliceResult` 分段回推。每批約 `3` 到 `5` 封 mail metadata，前面批次使用 `isSliceComplete=false`；該 folder 的最後一批才使用 `isSliceComplete=true`。只有整個 search 的最後一個 folder slice 最後一批可以使用 `isFinal=true`，或另外呼叫 `CompleteMailSearchSlice`。

AddIn 不需要回報 search progress；只需回推 slice result 與 complete event。

### MailPropertiesCommandRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "isRead": true,
  "flagInterval": "today",
  "flagRequest": "今天",
  "taskStartDate": "2026-05-03T16:00:00Z",
  "taskDueDate": "2026-05-03T16:00:00Z",
  "taskCompletedDate": null,
  "categories": ["Customer", "Follow-up"],
  "newCategories": [
    {
      "name": "Customer",
      "color": "olCategoryColorGreen",
      "colorValue": 5,
      "shortcutKey": ""
    }
  ]
}
```

`mailId` 不可為空。若工作機 AddIn push 回來的 mail 沒有 `id`，後續 mail mutation command 無法可靠執行。

`flagInterval` 目前預期值：

- `none`
- `today`
- `tomorrow`
- `this_week`
- `next_week`
- `no_date`
- `custom`
- `complete`

### CategoryCommandRequest

```json
{
  "name": "Project",
  "color": "olCategoryColorGreen",
  "colorValue": 5,
  "shortcutKey": ""
}
```

### CreateFolderRequest

```json
{
  "parentFolderPath": "\\\\Mailbox - User\\Projects",
  "name": "Sample Folder"
}
```

### DeleteFolderRequest

```json
{
  "folderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

`delete_folder` 的語意是 soft delete：AddIn 必須將指定 folder 移到同一個 Outlook store / mailbox 的 default Deleted Items folder，不得永久刪除 folder，也不得呼叫會直接永久移除 folder tree 的 API。目的 folder 必須透過 Outlook default folder identity 定位，例如 `Store.GetDefaultFolder(olFolderDeletedItems)` 或等效流程；不可用顯示名稱、本地化名稱、`folderPath` 字串包含 `Deleted Items` / `刪除的郵件` 等方式猜測。

AddIn 應用 `folderPath` 找到 Outlook folder object，再把該 folder move 到 default Deleted Items folder。若目標是 store root、hidden/system folder 或 Outlook object model 拒絕 move，AddIn 只需回報實際 automation 失敗診斷；不得改用永久刪除。完成後用 folder 增量同步回推 source parent、Deleted Items folder 與被移動 folder tree 的最新狀態；若目前 mail list 指向被移動 folder，也要回推 `PushMails` 清掉或更新畫面。

### MoveMailRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample Folder"
}
```

AddIn 應用 `mailId` 找到 Outlook item，將 `destinationFolderPath` 解析成 Outlook `Folder` object，呼叫 Outlook `MailItem.Move(destinationFolder)`，完成後回推最新 `PushMails`，並用 folder 增量同步更新 folder count。

若 `destinationFolderPath` 解析到 Outlook default Deleted Items folder，這仍然只是移動郵件到該 folder，不是永久刪除。AddIn 必須沿用同一個 `MailItem.Move(destinationFolder)` 流程，不可因目的 folder 是 Deleted Items 而改呼叫 `MailItem.Delete()`。

注意：Microsoft 文件說 Outlook `MailItem.EntryID` 在 item save 或 send 後才會存在，跨 store 移動時可能改變。因此 AddIn 若使用 EntryID 當 `MailItemDto.id`，移動後應重新讀取並回推最新 mail snapshot。相關官方依據請看 `features-checklist.md`。

### MoveMailsRequest

```json
{
  "mailIds": [
    "[redacted Outlook EntryID or stable id 1]",
    "[redacted Outlook EntryID or stable id 2]"
  ],
  "sourceFolderPath": "\\\\Mailbox - User\\Inbox",
  "sourceFolderPaths": ["\\\\Mailbox - User\\Inbox"],
  "destinationFolderPath": "\\\\Mailbox - User\\Projects\\Sample Folder",
  "continueOnError": true
}
```

AddIn 應逐封用 `mailIds` 找到 Outlook item，將 `destinationFolderPath` 解析成 Outlook `Folder` object，對每封 mail 呼叫 Outlook `MailItem.Move(destinationFolder)`。`sourceFolderPath` 是單一來源 folder 的簡寫；`sourceFolderPaths` 供搜尋結果跨 folder 批次移動時更新多個來源 folder count。若部分 mail 找不到，`continueOnError=true` 時繼續處理剩餘 mail，最後在 `ReportCommandResult.payload` 放簡短統計，不要塞完整郵件內容。

HTTP API `POST /api/outlook/request-move-mails` 單次最多接受 500 個 `mailIds`。若 caller 需要移動更多郵件，必須分批呼叫，避免單一 HTTP payload 與 Outlook automation command 執行時間過大。

完成後 AddIn 應回推最新 `PushMails`，並用 folder 增量同步更新所有來源 folder 與 destination folder item count。跨 PST / OST 移動後 `EntryID` 可能改變，因此 AddIn 應重新讀取最新 mail snapshot。

### DeleteMailRequest

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox"
}
```

AddIn 應用 `mailId` 與 `folderPath` 找到 Outlook item，再以該 item 所在 store / mailbox 的 default Deleted Items folder 作為 destination，呼叫 Outlook `MailItem.Move(destinationFolder)`。目的 folder 必須透過 Outlook default folder identity 定位，例如 `Store.GetDefaultFolder(olFolderDeletedItems)` 或等效流程；不可用顯示名稱、本地化名稱、`folderPath` 字串或顯示文字猜測。完成後回推最新 `PushMails`，並用 folder 增量同步更新 source folder 與 Deleted Items folder count。這個 command 不得呼叫 Outlook `MailItem.Delete()`。

## Push Payload Sample

### FolderSyncBatchDto

```json
{
  "syncId": "folder-sync-001",
  "sequence": 1,
  "reset": true,
  "isFinal": false,
  "stores": [
    {
      "storeId": "[redacted primary store id]",
      "displayName": "主要信箱 - User",
      "storeKind": "ost",
      "storeFilePath": "C:\\Users\\User\\AppData\\Local\\Microsoft\\Outlook\\user@example.com.ost",
      "rootFolderPath": "\\\\主要信箱 - User"
    }
  ],
  "folders": [
    {
      "name": "主要信箱 - User",
      "folderPath": "\\\\主要信箱 - User",
      "parentFolderPath": "",
      "itemCount": 0,
      "storeId": "[redacted primary store id]",
      "isStoreRoot": true,
      "folderType": "StoreRoot",
      "defaultItemType": -1,
      "isHidden": false,
      "isSystem": false
    },
    {
      "name": "Inbox",
      "folderPath": "\\\\主要信箱 - User\\Inbox",
      "parentFolderPath": "\\\\主要信箱 - User",
      "itemCount": 18,
      "storeId": "[redacted primary store id]",
      "isStoreRoot": false,
      "folderType": "Inbox",
      "defaultItemType": 0,
      "isHidden": false,
      "isSystem": false
    }
  ]
}
```

### MailItemDto

```json
[
  {
    "id": "[redacted Outlook EntryID or stable id]",
    "subject": "[redacted] sample subject",
    "sender": {
      "recipientKind": "sender",
      "displayName": "Sample Sender",
      "smtpAddress": "sender@example.invalid",
      "rawAddress": "/O=ORG/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=sender",
      "addressType": "EX",
      "entryUserType": "olExchangeUserAddressEntry",
      "isGroup": false,
      "isResolved": true,
      "members": []
    },
    "toRecipients": [
      {
        "recipientKind": "to",
        "displayName": "Group A",
        "smtpAddress": "group-a@example.invalid",
        "rawAddress": "group-a@example.invalid",
        "addressType": "SMTP",
        "entryUserType": "olExchangeDistributionListAddressEntry",
        "isGroup": true,
        "isResolved": true,
        "members": [
          {
            "recipientKind": "member",
            "displayName": "Member One",
            "smtpAddress": "member.one@example.invalid",
            "rawAddress": "member.one@example.invalid",
            "addressType": "SMTP",
            "entryUserType": "olExchangeUserAddressEntry",
            "isGroup": false,
            "isResolved": true,
            "members": []
          }
        ]
      }
    ],
    "ccRecipients": [],
    "bccRecipients": [],
    "receivedTime": "2026-05-04T01:30:00Z",
    "body": "",
    "bodyHtml": "",
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "messageClass": "IPM.Note",
    "conversationId": "[redacted ConversationID]",
    "conversationTopic": "[redacted normalized conversation topic]",
    "conversationIndex": "[redacted ConversationIndex]",
    "categories": "Customer, Follow-up",
    "isRead": false,
    "isMarkedAsTask": true,
    "flagRequest": "今天",
    "flagInterval": "today",
    "taskStartDate": "2026-05-03T16:00:00Z",
    "taskDueDate": "2026-05-03T16:00:00Z",
    "taskCompletedDate": null,
    "importance": "high",
    "sensitivity": "normal"
  }
]
```

`fetch_mails` 回推的 `MailItemDto` 應只包含 metadata，`body` / `bodyHtml` 留空。收到 `fetch_mail_body` 後，AddIn 再用 `PushMailBody` 回推內容：

Inbox 內可開啟的會議邀請/更新是 Outlook `MeetingItem`，常見 `messageClass` 為 `IPM.Schedule.Meeting.*`。AddIn 不可只用 `MailItem` cast 讀取列表與內文；至少 `fetch_folder_mails_slice`、`fetch_mail_search_slice`、`fetch_mail_body`、`fetch_mail_attachments` 應能讀取 `MeetingItem` metadata/body/attachments。`move_mail`、`move_mails` 與 `delete_mail` 應用 EntryID 找回可移動的 Outlook item 後執行 move；回覆會議、接受/拒絕、分類/旗標或修改行事曆則屬 Calendar/meeting command，不應混入一般 mail properties mutation。

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "body": "[redacted plain text body]",
  "bodyHtml": "<p>[redacted html body]</p>"
}
```

### MailConversationDto

`fetch_mail_conversation` 的 request object：

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "maxCount": 100,
  "includeBody": true
}
```

AddIn 應用 Outlook `MailItem.GetConversation()` 取得 conversation，優先用 `Conversation.GetRootItems()` / `GetChildren()` 保留討論串順序；若 store 不支援 conversation 或 conversation 為 null，可回傳只包含原 mail 的 `MailConversationDto`。Microsoft 官方文件指出 `GetConversation()` 可能因 item 未保存/未寄出、registry 關閉 conversation，或 store 不支援 conversation view 而回 null；也可用 `Store.IsConversationEnabled` 先判斷。

```json
{
  "mailId": "[redacted Outlook EntryID or stable id]",
  "folderPath": "\\\\Mailbox - User\\Inbox",
  "conversationId": "[redacted ConversationID]",
  "conversationTopic": "[redacted normalized conversation topic]",
  "mails": [
    {
      "id": "[redacted Outlook EntryID or stable id]",
      "subject": "Re: [redacted]",
      "body": "[redacted plain text body when includeBody=true]",
      "bodyHtml": "<p>[redacted html body when includeBody=true]</p>"
    }
  ]
}
```

### OutlookRuleDto

```json
[
  {
    "name": "Move customer mail",
    "enabled": true,
    "executionOrder": 1,
    "ruleType": "receive",
    "conditions": ["sender contains example.com"],
    "actions": ["move to \\\\Mailbox - User\\Customers"],
    "exceptions": []
  }
]
```

### OutlookCategoryDto

```json
[
  {
    "name": "Customer",
    "color": "olCategoryColorGreen",
    "colorValue": 5,
    "shortcutKey": ""
  }
]
```

### CalendarEventDto

```json
[
  {
    "id": "[redacted appointment id if available]",
    "subject": "[redacted meeting subject]",
    "start": "2026-05-04T02:00:00Z",
    "end": "2026-05-04T03:00:00Z",
    "location": "Meeting Room",
    "organizer": {
      "recipientKind": "organizer",
      "displayName": "Sample Organizer",
      "smtpAddress": "organizer@example.invalid",
      "rawAddress": "organizer@example.invalid",
      "addressType": "SMTP",
      "entryUserType": "olExchangeUserAddressEntry",
      "isGroup": false,
      "isResolved": true,
      "members": []
    },
    "requiredAttendees": [
      {
        "recipientKind": "required",
        "displayName": "Sample Attendee",
        "smtpAddress": "attendee@example.invalid",
        "rawAddress": "attendee@example.invalid",
        "addressType": "SMTP",
        "entryUserType": "olExchangeUserAddressEntry",
        "isGroup": false,
        "isResolved": true,
        "members": []
      }
    ],
    "isRecurring": false,
    "busyStatus": "busy"
  }
]
```

### ChatMessageDto

```json
{
  "id": "[optional client-generated id]",
  "source": "outlook",
  "text": "AddIn message",
  "timestamp": "2026-05-04T02:00:00Z"
}
```

AddIn 透過 `SendChatMessage(message)` 送出時，可省略 `timestamp`；`source` 空白時會視為 `outlook`。

## Request Endpoint 背景

AddIn 不需要呼叫下列 endpoint。這些 endpoint 只列作 command 來源對照，方便工作機測試時確認「哪個 HTTP request 會變成哪個 `OutlookCommand`」。

| Method | Path | Command |
| --- | --- | --- |
| `POST` | `/api/outlook/request-folders` | `fetch_folder_roots` |
| `POST` | `/api/outlook/request-folder-children` | `fetch_folder_children` |
| `POST` | `/api/outlook/request-mails` | `fetch_mails` |
| `POST` | `/api/outlook/request-folder-mails` | `fetch_folder_mails_slice` |
| `POST` | `/api/outlook/request-mail-search` | `fetch_mail_search_slice` |
| `POST` | `/api/outlook/request-mail-body` | `fetch_mail_body` |
| `POST` | `/api/outlook/request-mail-attachments` | `fetch_mail_attachments` |
| `POST` | `/api/outlook/request-mail-conversation` | `fetch_mail_conversation` |
| `POST` | `/api/outlook/request-export-mail-attachment` | `export_mail_attachment` |
| `POST` | `/api/outlook/open-exported-attachment` | 不 dispatch 給 AddIn |
| `POST` | `/api/outlook/request-rules` | `fetch_rules` |
| `POST` | `/api/outlook/request-categories` | `fetch_categories` |
| `POST` | `/api/outlook/request-signalr-ping` | `ping` |
| `POST` | `/api/outlook/request-calendar` | `fetch_calendar` |
| `POST` | `/api/outlook/request-address-book` | `fetch_address_book` |
| `POST` | `/api/outlook/request-update-mail-properties` | `update_mail_properties` |
| `POST` | `/api/outlook/request-upsert-category` | `upsert_category` |
| `POST` | `/api/outlook/request-create-folder` | `create_folder` |
| `POST` | `/api/outlook/request-delete-folder` | `delete_folder` |
| `POST` | `/api/outlook/request-move-mail` | `move_mail` |
| `POST` | `/api/outlook/request-move-mails` | `move_mails` |
| `POST` | `/api/outlook/request-delete-mail` | `delete_mail` |

沒有 AddIn SignalR connection 時，HTTP request endpoint 通常仍會先回 `accepted`，正式 caller 應用 paired `fetch-result-*` 取得後續狀態。paired fetch result 會把 AddIn unavailable 映射成：

```json
{
  "requestId": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "request": "request-mails",
  "state": "unavailable",
  "message": "No Outlook AddIn SignalR connection is available.",
  "next": {
    "cursor": "",
    "hasMore": false
  },
  "data": {}
}
```

外部 client 可用下列 endpoint 查詢 command 執行狀態；AddIn 不需要呼叫：

| Method | Path | 說明 |
| --- | --- | --- |
| `GET` | `/api/outlook/command-results/{commandId}` | 查詢指定 command 的 `pending` / `completed` / `failed` / `addin_unavailable` 狀態 |
| `GET` | `/api/outlook/command-results` | 查詢最近 command 執行狀態 |
| `GET` | `/api/outlook/mail-search/progress/{searchId}` | 查詢指定 search id 的進度 |
| `GET` | `/api/outlook/mail-search/progress/by-command/{commandId}` | 用 command id 查詢對應 search 進度 |

## DTO 欄位速查

### MailItemDto

- `id`: string
- `subject`: string
- `sender`: `OutlookRecipientDto`，寄件者；AddIn 應使用 Outlook 通訊錄解析後的直觀 `displayName`。
- `toRecipients`: `OutlookRecipientDto[]`
- `ccRecipients`: `OutlookRecipientDto[]`
- `bccRecipients`: `OutlookRecipientDto[]`
- `receivedTime`: DateTime，transport 必須是 UTC。
- `body`: string，`fetch_mails` 時應留空；只在單封內容載入後填入。
- `bodyHtml`: string，`fetch_mails` 時應留空；只在單封內容載入後填入。
- `folderPath`: string
- `messageClass`: string，Outlook `MessageClass`，例如一般郵件 `IPM.Note`、會議邀請/更新 `IPM.Schedule.Meeting.*`。
- `categories`: string
- `isRead`: boolean
- `isMarkedAsTask`: boolean
- `attachmentCount`: number，附件數 metadata；未知時可為 `0`，完整 metadata 仍以 `fetch_mail_attachments` / `PushMailAttachments` 為準。
- `attachmentNames`: string，附件名稱摘要；多個附件名稱建議以 `、` 或 `, ` 串接；避免放入檔案內容或本機路徑。
- `flagRequest`: string
- `flagInterval`: string
- `taskStartDate`: DateTime 或 `null`，transport 必須是 UTC。
- `taskDueDate`: DateTime 或 `null`，transport 必須是 UTC。
- `taskCompletedDate`: DateTime 或 `null`，transport 必須是 UTC。
- `importance`: string，預設 `normal`
- `sensitivity`: string，預設 `normal`

### OutlookRecipientDto

- `recipientKind`: string，`sender`、`to`、`cc`、`bcc`、`organizer`、`required` 或 `member`。
- `displayName`: string，client 預設顯示的名稱；應使用 Outlook resolved recipient / address book 的直觀名稱。
- `smtpAddress`: string，若 Outlook / Exchange 可解析出 SMTP address 則填入。
- `rawAddress`: string，原始 Outlook address；可能是 `/O=.../OU=.../CN=...` 這類 Exchange legacyDN。
- `addressType`: string，Outlook 常見值為 `SMTP` 或 `EX`。
- `entryUserType`: string，Outlook `AddressEntryUserType` 名稱，例如 `olExchangeUserAddressEntry` 或 `olExchangeDistributionListAddressEntry`。
- `isGroup`: boolean，distribution list / group 時為 `true`。
- `isResolved`: boolean，Outlook recipient/address entry 是否已解析。
- `members`: `OutlookRecipientDto[]`，group 已展開的成員；無法或未展開時保持空陣列。

### MailAttachmentDto

AddIn 處理 `fetch_mail_attachments` 時，請從 Outlook `MailItem.Attachments` 逐筆建立 metadata；依 Microsoft Outlook Interop 文件，附件名稱應優先使用 `Attachment.FileName`，沒有檔名時再用 `Attachment.DisplayName`，大小使用 `Attachment.Size`。附件識別請使用同一封 mail 內穩定可 round-trip 的值；Office COM collection 為 1-based，因此可用 `Attachment.Index.ToString()` 作為 `attachmentId`，export 時再用此值取回 `Attachments.Item(index)`。

- `mailId`: string，必須等於 request 的 `mailId`。
- `attachmentId`: string，必填；建議使用 Outlook `Attachment.Index` 的字串值，或 AddIn 自己能在同一封 mail 內穩定查回的 id。
- `index`: number，可填 Outlook `Attachment.Index`。
- `name`: string，顯示名稱；建議填 `Attachment.FileName`，沒有時填 `Attachment.DisplayName`。
- `fileName`: string，可填 Outlook `Attachment.FileName`。
- `displayName`: string，可填 Outlook `Attachment.DisplayName`。
- `contentType`: string，可空；Outlook Object Model 沒有直接暴露 MIME type 時不要硬猜。
- `size`: number，可填 Outlook `Attachment.Size`；Microsoft 文件說部分情況可能拿不到實際大小而回 `0`。
- `isExported`: boolean，尚未匯出時為 `false`。
- `exportedAttachmentId`: string，尚未匯出時空白。
- `exportedPath`: string，尚未匯出時空白。

### ExportMailAttachmentRequest

收到 `export_mail_attachment` 時會帶下列欄位。對 Outlook COM/VSTO AddIn，請優先使用 `index` 或可解析為整數的 `attachmentId` 取回 `mail.Attachments.Item(index)`；這是配合 Microsoft Outlook Object Model 的 1-based `Attachments` collection。若 `attachmentId` 不是數字，AddIn 可用自己在 metadata 階段建立的 mapping 查回附件。

- `mailId`: string，目標 mail id。
- `folderPath`: string，目標 mail 所在 folder path。
- `attachmentId`: string；若 metadata 有 `index`，request 會帶 `index.ToString()`，方便 AddIn 直接呼叫 `Attachments.Item(index)`。
- `index`: number；Outlook `Attachment.Index`。
- `name`: string；目前顯示的附件名稱。
- `fileName`: string；metadata 中的 Outlook `Attachment.FileName`。
- `displayName`: string；metadata 中的 Outlook `Attachment.DisplayName`。
- `exportRootPath`: string；允許的 attachment export root。AddIn 輸出檔案必須放在此 root 底下。

### ExportedMailAttachmentDto

AddIn 處理 `export_mail_attachment` 時，請用 request 的 `attachmentId` 找回同一個 Outlook `Attachment`，將檔案儲存到 request 指定的 attachment root 底下，呼叫 Outlook `Attachment.SaveAsFile(path)` 後再 `PushExportedMailAttachment`。

- `mailId`: string，必須等於 request 的 `mailId`。
- `folderPath`: string。
- `attachmentId`: string，必須等於 request 的 `attachmentId`。
- `exportedAttachmentId`: string，可由 AddIn 產生；空白時可由接收端補值。
- `name`: string，建議與 metadata 階段相同。
- `fileName`: string，可填 Outlook `Attachment.FileName`。
- `displayName`: string，可填 Outlook `Attachment.DisplayName`。
- `contentType`: string，可空。
- `size`: number，建議填實際輸出檔案長度；拿不到時可用 Outlook `Attachment.Size`。
- `exportedPath`: string，必填；必須是 `SaveAsFile(path)` 實際輸出的完整本機路徑。
- `exportedAt`: DateTime，transport 必須是 UTC。

### FolderDto

- `name`: string
- `entryId`: string，Outlook folder `EntryID`，後續 request 會搭配 `storeId` 指定 parent。
- `folderPath`: string
- `parentEntryId`: string，store root 可為空字串。
- `parentFolderPath`: string，store root 可為空字串。
- `itemCount`: number
- `storeId`: string，Outlook Store ID 或 AddIn 內可追蹤的 store identifier。
- `isStoreRoot`: boolean，folder 是否是該 store 的 root folder。
- `folderType`: `OutlookFolderType` enum 字串，正式值見下方；AddIn 依 Outlook `OlDefaultFolders`、`Folder.DefaultItemType` 與 MAPI flags 判定後回傳 enum，不回傳本地化 folder name policy。
- `defaultItemType`: number，Outlook `OlItemType` 數值；mail folder 必須是 `0` / `olMailItem`，store root 或無法判定時填 `-1`。
- `isHidden`: boolean，AddIn 以 MAPI `PR_ATTR_HIDDEN` / proptag `0x10F4000B` 讀到的值；讀取失敗時填 `false` 並在測試回報記錄。
- `isSystem`: boolean，AddIn 以 MAPI `PR_ATTR_SYSTEM` / proptag `0x10F5000B` 讀到的值；讀取失敗時填 `false` 並在測試回報記錄。
- `hasChildren`: boolean，該 folder 是否可能有直接 children。
- `childrenLoaded`: boolean，該 folder 的直接 children 是否已由指定 command 載入。
- `discoveryState`: string，預期 `partial`、`loaded` 或 `failed`。

`OutlookFolderType` 正式 enum 值：

- `Unknown`
- `StoreRoot`
- `Mail`
- `Inbox`
- `Sent`
- `Drafts`
- `Deleted`
- `Junk`
- `Archive`
- `Outbox`
- `SyncIssues`
- `Conflicts`
- `LocalFailures`
- `ServerFailures`
- `Calendar`
- `Contacts`
- `Tasks`
- `Notes`
- `Journal`
- `RssFeeds`
- `ConversationHistory`
- `ConversationActionSettings`
- `OtherSystem`

`FolderDto` 不再包含 `subFolders`，也不再重複保存 store display name / type / file path。tree 由 `parentFolderPath` 與 `storeId` 組回。AddIn 不得保留全量 folder tree command。`defaultItemType != 0`、`isHidden == true`、`isSystem == true`，或 `folderType` 是 blocked enum 的非 root folder，視為不可操作 folder。AddIn 不得用本地化 folder name 猜測是否為系統資料夾。

### OutlookStoreDto

- `storeId`: string，Outlook Store ID 或 AddIn 內可追蹤的 store identifier。
- `displayName`: string，Outlook store 顯示名稱。
- `storeKind`: string，目前預期 `ost`、`pst`、`exchange` 或 `other`。
- `storeFilePath`: string，`.pst` 或 `.ost` 的完整檔案路徑；沒有檔案路徑時可空字串。
- `rootFolderPath`: string，該 store root folder 的 `folderPath`。

### OutlookRuleDto

- `name`: string
- `enabled`: boolean
- `executionOrder`: number
- `ruleType`: string，預設 `receive`
- `conditions`: string[]
- `actions`: string[]
- `exceptions`: string[]

### OutlookCategoryDto

- `name`: string
- `color`: string，Outlook `OlCategoryColor` enum name，例如 `olCategoryColorGreen`。
- `colorValue`: number，Outlook `OlCategoryColor` enum numeric value，例如 `5`。Add-in 寫入 Outlook 時應優先使用此欄位。
- `shortcutKey`: string

### CalendarEventDto

- `id`: string
- `subject`: string
- `start`: DateTime，transport 必須是 UTC。
- `end`: DateTime，transport 必須是 UTC。
- `location`: string
- `organizer`: `OutlookRecipientDto`
- `requiredAttendees`: `OutlookRecipientDto[]`
- `isRecurring`: boolean
- `busyStatus`: string

### ChatMessageDto

- `id`: string，可空；未填時可由接收端產生預設 id。
- `source`: string，AddIn 建議填 `outlook`。
- `text`: string
- `timestamp`: DateTime，transport 必須是 UTC。，AddIn 送出時可空；transport 必須是 UTC。

### AddinLogEntry

- `level`: string，預期 `info`、`warn` 或 `error`
- `message`: string
- `timestamp`: DateTime，transport 必須是 UTC。
