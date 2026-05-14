# Outlook AddIn 功能實作 Checklist

本文件是工作機 AI 實作 Outlook AddIn 功能的入口。請先用本 checklist 對照缺口；只有需要 payload 細節或 Outlook object model 依據時，再查看後面的參考文件。

AddIn 的角色必須保持單純：listen `OutlookCommand`、呼叫 Outlook object model / Office automation、invoke `Push*` 與 `ReportCommandResult`。不要在 AddIn 裡實作跨 command 排程、跨 folder 負載管理、資料合併或對外 client workflow。

## 先讀這份，必要時再查

- SignalR / DTO / payload 細節：`signalr-contract.md`
- Office 2016 / Outlook 官方文件入口：`outlook-references.md`
- 工作機測試回報格式：`test-report.md`

## 完成定義

工作機 AddIn 視為完成時，必須同時符合：

- [ ] AddIn 可連線 `/hub/outlook-addin`，並 invoke `RegisterOutlookAddin(info)`。
- [ ] AddIn 可 listen `OutlookCommand` 並處理本文件列出的 command。
- [ ] 每個 command 完成後都有 `ReportCommandResult`，失敗時 message 可診斷且不得含敏感資料。
- [ ] 會改變 Outlook snapshot 的 command 會 invoke 對應 `Push*`。
- [ ] 所有 `MailItemDto.id` 都是非空，且 AddIn 可用它找回該 Outlook item。
- [ ] `MailItemDto` 若 Outlook 有提供，應填入 `conversationId`、`conversationTopic`、`conversationIndex`，並支援 `fetch_mail_conversation` / `PushMailConversation`。
- [ ] Folder tree 第一層是 Outlook Store：主要 OST 與每個 PST 都是獨立頂層。
- [ ] 測試資料與錯誤回報需匿名化。
- [ ] 不保留舊版或未使用功能：不要實作 `/api/outlook/poll`、`/api/outlook/push-*`、HTTP chat 或未列於 contract 的 legacy command。
- [ ] 效能最佳化必須有 Microsoft 官方文件或工作機實測依據；沒有依據時，選擇最薄、最直接、最容易診斷的 Outlook API 呼叫流程。

## 功能總覽

| 功能 | Command / Method | AddIn 完成後必須回推 |
| --- | --- | --- |
| AddIn 連線 | `ping` | `ReportCommandResult` |
| Folder discovery | `fetch_folder_roots`、`fetch_folder_children` | `BeginFolderSync`、`PushFolderBatch`、`CompleteFolderSync` |
| 讀取近期郵件 | `fetch_mails` | `PushMails` metadata |
| 直接列出 folder 郵件 | `fetch_folder_mails_slice` | `BeginFolderMails`、`PushFolderMailsSliceResult`、`CompleteFolderMailsSlice` |
| 搜尋 / 篩選郵件 | `fetch_mail_search_slice` | `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice` |
| 讀取郵件內容 | `fetch_mail_body` | `PushMailBody` |
| 讀取附件清單 | `fetch_mail_attachments` | `PushMailAttachments` |
| 匯出附件 | `export_mail_attachment` | `PushExportedMailAttachment` |
| 修改郵件屬性 | `update_mail_properties` | `PushMail`，必要時 `PushCategories` |
| 移動郵件 | `move_mail`、`move_mails` | `PushMails`、folder 增量同步 |
| 刪除郵件 | `delete_mail` | `PushMails`、folder 增量同步；實作為 move to Outlook default Deleted Items folder |
| Master categories | `fetch_categories`、`upsert_category` | `PushCategories` |
| Rules snapshot / mutation | `fetch_rules`、`manage_rule` | `PushRules` |
| 月曆 | `fetch_calendar` | `PushCalendar` |
| 通訊錄 | `fetch_address_book` | `PushAddressBookBatch` |
| 通訊錄 group members | `fetch_address_book_group_members` | `PushAddressBookGroupMembersBatch` |
| Chat | `SendChatMessage` SignalR server method | `ReportCommandResult` 不適用；method invoke 成功即可 |
| Folder 建立 / 刪除 | `create_folder`、`delete_folder` | folder 增量同步，必要時 `PushMails`；`delete_folder` 是 move to Outlook default Deleted Items folder |

## 必做 Checklist

### 1. SignalR 基礎

- [ ] AddIn 啟動後連線 `/hub/outlook-addin`。
- [ ] 連線成功後 invoke `RegisterOutlookAddin(info)`。
- [ ] AddIn 收到 `type: "ping"`。
- [ ] AddIn 對 `ping` invoke `ReportCommandResult`；只有 Outlook object model 可正常呼叫時才回 `success=true`。
- [ ] AddIn code 不再包含 HTTP poll、HTTP push 或 HTTP chat fallback。

驗收：

- [ ] AddIn 註冊後可收到 command。
- [ ] AddIn 可回報 `ping` 的 command result。

### 2. Folder Discovery

第一層必須是 Outlook Store。AddIn 收到 `fetch_folder_children` 時只讀 request 指定的 parent；不得一次遞迴整棵樹。

- [ ] AddIn 收到 `fetch_folder_roots`。
- [ ] `fetch_folder_roots` 只使用 Outlook `Application.Session.Stores` 列出目前 profile 的所有 stores。
- [ ] `fetch_folder_roots` 只對每個 store 使用 `Store.GetRootFolder()` 取得 root folder，不讀 subfolders。
- [ ] AddIn 收到 `fetch_folder_children`。
- [ ] `fetch_folder_children` 使用 `storeId` + `parentEntryId` 優先定位 parent folder，必要時才用 `parentFolderPath`。
- [ ] `fetch_folder_children` 只讀 parent 的直接 children，除非 request 指定較大的 `maxDepth`。
- [ ] invoke `BeginFolderSync` 開始 folder 增量同步。
- [ ] 用 `PushFolderBatch` 分批送回 `OutlookStoreDto[]` 與 flat `FolderDto[]`。
- [ ] invoke `CompleteFolderSync` 結束 folder 增量同步。
- [ ] 主要 OST 作為第一個 store root，`storeKind = "ost"`。
- [ ] 每個 PST 各自作為一個 store root，`storeKind = "pst"`。
- [ ] 每個 `OutlookStoreDto` 都填入：
  - `storeId`
  - `displayName`
  - `storeKind`
  - `storeFilePath`
  - `rootFolderPath`
- [ ] 每個 `FolderDto` 都填入：
  - `name`
  - `entryId`
  - `folderPath`
  - `parentEntryId`
  - `parentFolderPath`
  - `itemCount`
  - `storeId`
  - `isStoreRoot`
  - `folderType`
  - `defaultItemType`
  - `isHidden`
  - `isSystem`
  - `hasChildren`
  - `childrenLoaded`
  - `discoveryState`
- [ ] Store root folder 的 `parentFolderPath = ""` 且 `isStoreRoot = true`，底下 folder 都是 `false`。
- [ ] `folderType` 必須是 `OutlookFolderType` enum 字串；常見值包含 `StoreRoot`、`Mail`、`Inbox`、`Sent`、`Drafts`、`Deleted`、`Junk`、`Archive`、`Outbox`、`SyncIssues`、`Conflicts`、`LocalFailures`、`ServerFailures`、`Calendar`、`Contacts`、`Tasks`、`Notes`、`Journal`、`RssFeeds`、`ConversationHistory`、`ConversationActionSettings`、`OtherSystem`。
- [ ] `defaultItemType` 必須來自 Outlook `Folder.DefaultItemType`；mail folder 為 `0` / `olMailItem`，無法判定時填 `-1`。
- [ ] `isHidden` 與 `isSystem` 必須來自 MAPI `PR_ATTR_HIDDEN` / `PR_ATTR_SYSTEM`，不得用 folder name 或本地化顯示文字猜測。
- [ ] AddIn 可用 Outlook `Store.GetDefaultFolder(OlDefaultFolders)` 與目前 folder 的 identity 比對來判定 `folderType`，但 contract 只回傳 enum，不回傳 EntryID 對照表。
- [ ] 日誌、記事、Calendar、Contacts、Tasks、Sync Issues、Local Failures、Server Failures 等不可操作 folder 必須能用 `folderType`、`defaultItemType`、`isHidden` 或 `isSystem` 判定。
- [ ] `FolderDto` 不再包含 `subFolders`，也不重複傳 store metadata。
- [ ] `.pst` / `.ost` 的真實位置填在 `storeFilePath`；回報文件中必須匿名化路徑。

驗收：

- [ ] `OutlookStoreDto[]` 至少包含目前 profile 可見的 store。
- [ ] 每個 folder 的 `storeId` 可對回同一批 `OutlookStoreDto`。
- [ ] 展開 PST/OST 時，folder path 與 parent path 可組回正確 tree。

### 3. Mail List 與 Mail Identity

- [ ] AddIn 收到 `fetch_mails`。
- [ ] 依 `mailsRequest.folderPath` 讀取該 folder 的 mail。
- [ ] 支援 request 內的 `receivedFrom` / `receivedTo` date-time 邊界。
- [ ] 支援 `maxCount`。
- [ ] 回推 `PushMails(mails)`，只包含 metadata。
- [ ] 每筆 mail 的 `id` 必填，建議使用 Outlook `MailItem.EntryID` 或 AddIn 可穩定找回 item 的識別。
- [ ] `folderPath` 必須對應目前 mail 所在 folder。
- [ ] `body` 與 `bodyHtml` 在 `fetch_mails` 回應中必須留空，避免一次載入大量郵件內容。
- [ ] AddIn 收到 `fetch_mail_body`。
- [ ] 依 `mailBodyRequest.mailId` 與 `folderPath` 找回單封 mail。
- [ ] 回推 `PushMailBody(body)`，包含 `mailId`、`folderPath`、`body` 與 `bodyHtml`。
- [ ] AddIn 收到 `fetch_mail_attachments`。
- [ ] 回推 `PushMailAttachments(attachments)`，只包含附件 metadata。
- [ ] AddIn 收到 `export_mail_attachment`。
- [ ] 將指定附件匯出到 request 指定的 attachment root，回推 `PushExportedMailAttachment(exported)`。
- [ ] AddIn 只負責 export，不負責開啟附件。

驗收：

- [ ] 每封 mail 都有非空 `id`。
- [ ] `fetch_mails` 不載入完整 body。
- [ ] `fetch_mail_body` 可用 `mailId` 找回並回推同一封 mail 的內容。
- [ ] `fetch_mail_attachments` 與 `export_mail_attachment` 可用同一組附件識別 round-trip。
- [ ] 沒有出現「缺少 id」警告。
- [ ] 已讀、flag、category、move/delete 都能用有效 `mailId` 執行。

### 4. 直接列出 Folder 郵件

這一段只處理「使用者指定 folder，請把該 folder 底下的 mail id 列出來」。這不是搜尋，也不是 keyword search。

- [ ] AddIn 收到 `fetch_folder_mails_slice`。
- [ ] AddIn 使用 `folderMailsSliceRequest.storeId` 與 `folderEntryId` 優先定位單一 Outlook folder；`folderPath` 只作為顯示與 fallback。
- [ ] AddIn 只能使用該 folder 的 `Items` 讀取 mail metadata；必要時可用 `Items.Restrict` 套用 `receivedFrom` / `receivedTo`。
- [ ] AddIn 不得為 `fetch_folder_mails_slice` 呼叫 `Application.AdvancedSearch`、DASL content search 或全域搜尋。
- [ ] AddIn 使用 `BeginFolderMails`、`PushFolderMailsSliceResult`、`CompleteFolderMailsSlice` 回傳 folder mails 結果。
- [ ] `PushFolderMailsSliceResult` 必須帶回 `folderMailsId`、`commandId`、`parentCommandId`、`sliceIndex` 與 `sliceCount`。
- [ ] 結果只回 metadata，`body` / `bodyHtml` 留空。
- [ ] 同一個 folder slice 的結果必須分段回推，每批約 `3` 到 `5` 封 mail metadata；前面批次 `isSliceComplete=false`，最後一批才設為 `true`。
- [ ] 大量結果不得一次把整個 folder 的郵件塞進單一 SignalR payload。
- [ ] 發生 Outlook busy、timeout 或使用者取消時，使用 `CompleteFolderMailsSlice(success=false)` 並以匿名化 message 說明。

驗收：

- [ ] 指定單一 folder 時，AddIn 能回傳該 folder 的 mail metadata 與非空 `id`。
- [ ] 指定日期區間時，只回傳區間內 mails。
- [ ] 執行 folder mails 時沒有呼叫 Outlook search。
- [ ] Outlook busy、timeout 或取消時，有失敗結果而不是卡住。

### 5. Mail Search / Filter

這一段才處理「依條件找郵件」。AddIn 收到 `fetch_mail_search_slice` 時，必須只在 request 指定的單一 folder slice 內搜尋或篩選，不得因為 scope 欄位空白就自行展開整個 Outlook。

- [ ] AddIn 收到 `fetch_mail_search_slice`。
- [ ] AddIn 使用 `mailSearchSliceRequest.storeId` 與 `folderEntryId` 優先定位單一 Outlook folder；`folderPath` 只作為顯示、search scope 組合與 fallback。
- [ ] AddIn 若收到空 `storeId`、空 `folderEntryId` 或空 `folderPath`，必須用 `CompleteMailSearchSlice(success=false)` 結束該 slice；不得自行全域掃描。
- [ ] AddIn 依 `mailSearchSliceRequest.executionMode` 選擇流程：`items_filter` 使用 folder `Items` / `Items.Restrict` 或等效逐項 metadata filter；`outlook_search` 才使用 Outlook 內容搜尋。
- [ ] AddIn 在單一 folder 內套用 `keyword`、`textFields`、分類、附件、旗標、已讀狀態與時間；不要實作 typo-tolerant fuzzy search，也不要套用掃描數量限制。
- [ ] search result 只回 metadata，`body` / `bodyHtml` 留空。
- [ ] 使用 `BeginMailSearch`、`PushMailSearchSliceResult`、`CompleteMailSearchSlice` 回傳搜尋結果；不要用 `PushMails` 覆蓋目前 folder list。
- [ ] `PushMailSearchSliceResult` 必須帶回 `commandId`、`parentCommandId`、`sliceIndex` 與 `sliceCount`。
- [ ] 同一個 folder slice 的搜尋結果必須分段回推，每批約 `3` 到 `5` 封 mail metadata；前面批次 `isSliceComplete=false`，最後一批才設為 `true`。
- [ ] 大量結果不得一次把整個 folder 的符合郵件塞進單一 SignalR payload。
- [ ] AddIn 不需要自行做跨 folder 排程；單一 folder 搜尋仍應避免 blocking Outlook UI。
- [ ] 發生 Outlook busy、search timeout 或使用者取消時，使用 `CompleteMailSearchSlice(success=false)` 並以匿名化 message 說明。

驗收：

- [ ] `items_filter` 不呼叫 `Application.AdvancedSearch`。
- [ ] 只有 `executionMode="outlook_search"` 的 body keyword 搜尋會走 Outlook 內容搜尋。
- [ ] AddIn 不會收到空 scope 後自行全域掃描。
- [ ] Outlook busy、timeout 或取消時，有失敗結果而不是卡住。

### 6. 修改郵件屬性

目前只實作 `update_mail_properties` 作為郵件屬性 mutation 入口。不要再新增或維護舊的單一 marker command handler，除非 contract 明確恢復使用。

- [ ] AddIn 收到 `update_mail_properties`。
- [ ] 用 `mailPropertiesRequest.mailId` 找回 Outlook mail item。
- [ ] 套用 `isRead`：`isRead = true` 時 Outlook `UnRead = false`。
- [ ] 套用 flag：
  - `flagInterval = "none"`：清除 task/follow-up flag。
  - `today`、`tomorrow`、`this_week`、`next_week`、`no_date`：標記 task/follow-up，並設定 `FlagRequest` 與日期。
  - `custom`：使用 payload 的 `taskStartDate` / `taskDueDate`。
  - `complete`：設定完成狀態與 `taskCompletedDate`。
- [ ] 套用 mail categories：把 `categories` 寫回 Outlook mail item。
- [ ] 若 `newCategories` 不存在於 master category list，先建立或更新 master category。
- [ ] 儲存 mail item。
- [ ] invoke `ReportCommandResult`。
- [ ] invoke `PushMail` 更新畫面中的同一封 mail；不要重新抓取整個 mail list。
- [ ] 若 master category 有變更，invoke `PushCategories`。

驗收：

- [ ] Outlook mail item 的 read / flag / category 狀態正確保存。
- [ ] 回推的 `PushMail` 包含最新 snapshot。
- [ ] 若新增 category，回推的 `PushCategories` 包含最新 master category list。

### 7. 移動與刪除郵件

刪除郵件有獨立 `delete_mail` command；但唯一允許實作仍是移動到同一個 Outlook store / mailbox 的 default Deleted Items folder。AddIn 必須用 Outlook object model 的 default folder identity 定位目的地，不得用顯示名稱、本地化名稱或 `folderPath` 字串猜測。AddIn 不得直接呼叫 `MailItem.Delete()` 或永久刪除郵件。

- [ ] AddIn 收到 `move_mail`。
- [ ] AddIn 收到 `move_mails` 時，逐封用 `moveMailsRequest.mailIds` 找回 mail item；單次最多處理 500 封，更多郵件由 caller 分批呼叫。
- [ ] AddIn 收到 `delete_mail` 時，用同一套移動流程移到 Outlook default Deleted Items folder。
- [ ] `delete_mail` 的 destination 透過 `Store.GetDefaultFolder(olFolderDeletedItems)` 或等效 Outlook default folder identity 定位，不依賴 `Deleted Items`、`刪除的郵件` 或其他本地化 folder name。
- [ ] 用 `moveMailRequest.mailId` 找回 mail item。
- [ ] 用 `destinationFolderPath` 找到 Outlook destination `Folder`。
- [ ] 呼叫 Outlook `MailItem.Move(destinationFolder)`。
- [ ] 若 command 是 `delete_mail` 或 destination 是 Outlook default Deleted Items folder，仍只呼叫 `Move(destinationFolder)`，不可呼叫 `Delete()`。
- [ ] 即使 source mail 已經位於 default Deleted Items folder 或其子層，AddIn 仍不得永久刪除；只回報實際 move / no-op 結果。
- [ ] 移動後重新讀取目前 source folder 或以正確方式移除已移動 mail。
- [ ] invoke `ReportCommandResult`。
- [ ] invoke `PushMails`，讓目前 mail list 反映移動後結果。
- [ ] 用 folder 增量同步更新 source 與 destination folder item count。

驗收：

- [ ] `move_mail` 可把 mail 移到指定 folder。
- [ ] `move_mails` 可把多封 mail 移到指定 folder；部分失敗時依 `continueOnError` 回報統計。
- [ ] `delete_mail` 只會 move to Outlook default Deleted Items folder，不會永久刪除。
- [ ] Source folder mail snapshot 不再包含已移動 mail。
- [ ] 目的 folder item count 增加，source folder item count 減少。
- [ ] 跨 PST / OST 移動若 EntryID 改變，AddIn 仍會回推最新 mail snapshot。

### 8. Master Categories

- [ ] AddIn 收到 `fetch_categories`。
- [ ] 從 Outlook session master category list 讀取所有 category。
- [ ] 回推 `PushCategories(categories)`。
- [ ] AddIn 收到 `upsert_category`。
- [ ] 若 category 不存在，建立 category。
- [ ] 若 category 已存在，更新 color / shortcut key。
- [ ] 回推 `PushCategories(categories)`。

驗收：

- [ ] `PushCategories` 能反映 Outlook master category list。
- [ ] 新增或更新 category 後，回推最新 category snapshot。

### 9. Calendar 月曆

- [ ] AddIn 收到 `fetch_calendar`。
- [ ] 使用 `calendarRequest.startDate` 與 `calendarRequest.endDate` 讀取整個月份。
- [ ] `startDate` 含當日，`endDate` 不含當日。
- [ ] 回推區間內所有 calendar events。
- [ ] invoke `PushCalendar(events)`。
- [ ] AddIn 收到 `fetch_calendar_rooms`。
- [ ] 從 Outlook address lists 讀取可用 room/resource，回推 `PushCalendarRooms(rooms)`。
- [ ] AddIn 收到 `create_calendar_event`。
- [ ] 使用 Outlook `AppointmentItem` 建立 event，寫入 SmartOffice ownership marker，儲存後回推 calendar snapshot。
- [ ] AddIn 收到 `update_calendar_event` 或 `delete_calendar_event` 時，必須先確認 Outlook item 有 SmartOffice ownership marker，且 `smartOfficeEventId` 與 request 相符。
- [ ] 非 SmartOffice-owned event 或 `smartOfficeEventId` 不相符時，必須回報 `not_smartoffice_owned`，不得更新或刪除。

驗收：

- [ ] 回推的 event 落在 requested date range 內。
- [ ] Event 欄位包含 subject、時間、location、organizer、attendees、busy status。
- [ ] 只有 SmartOffice 建立的 event 可被更新或刪除。

### 10. Address Book 通訊錄

- [ ] AddIn 收到 `fetch_address_book`。
- [ ] 讀取 Outlook default Contacts folder 的 `ContactItem` metadata。
- [ ] 在設定允許時讀取 `Application.Session.AddressLists` 中可用的 Outlook address list / GAL metadata。
- [ ] 尊重 `addressBookRequest.maxContacts`、`maxAddressEntriesPerList`、`maxGroupMembers` 與 `maxGroupDepth`，不得無限制枚舉大型 GAL 或無限制展開 nested group。
- [ ] `maxGroupMembers=0` 時不得展開 group members，只回 group metadata。
- [ ] Distribution list / group 要回填 `isGroup`、`memberCount`、有限的 `memberSmtpAddresses` 與 `memberGroupSmtpAddresses`；無法展開時仍回傳 group metadata。
- [ ] `fetch_address_book` 必須是 read-only；不得呼叫 Outlook contact / address book 的 `Save()`、`Delete()`、`Move()`、`Items.Add()`、`Application.CreateItem()` 或任何會新增、刪除、修改通訊錄 entry 的 API。
- [ ] 回推 `PushAddressBookBatch(batch)`；不得讀取 mail body。
- [ ] AddIn 收到 `fetch_address_book_group_members` 時，只展開指定 group 的 direct members。
- [ ] `fetch_address_book_group_members` 不得自動遞迴展開 nested group；nested group 只標示 `isGroup=true`，由 Hub/Web UI 控制下一次手動展開。
- [ ] 通訊錄 partial progress 只能回推新增的 contact batch，不得送出 50、100、150 這類累積 snapshot。
- [ ] 完整 snapshot 也必須拆成固定大小 batch，第一個 batch 設 `reset=true`，最後一個 batch 設 `isFinal=true`。
- [ ] 完成後回報 `ReportCommandResult`。

驗收：

- [ ] Web UI 通訊錄同步後可查到 Outlook Contacts / GAL 來源的 contact。
- [ ] 若 Exchange / GAL 無法離線讀取，AddIn 仍回報清楚錯誤或至少回傳 Contacts folder 結果。
- [ ] Windows 主機實測確認 Outlook UI 沒有長時間卡住。

### 11. Rules Snapshot / Mutation

- [ ] AddIn 收到 `fetch_rules`。
- [ ] 讀取 Outlook rules snapshot。
- [ ] 回推 `PushRules(rules)`。
- [ ] AddIn 收到 `manage_rule`。
- [ ] `operation = "upsert"` 時，使用 Outlook `Rules.Create` 新增 rule，或用 `originalRuleName` + `originalExecutionOrder` 定位既有 rule 後修改支援的 definition。
- [ ] `operation = "set_enabled"` 時，只修改既有 rule 的 `Enabled`。
- [ ] `operation = "delete"` 時，使用 Outlook `Rules.Remove` 刪除指定 rule。
- [ ] 支援建立的條件包含 subject/body/body-or-subject/header text、sender/recipient address、category、has attachment、importance、to me、to or cc me、only to me、meeting invite/update；has attachment 只支援 `true`，不得承諾建立「無附件」rule。
- [ ] 支援建立的動作包含 move/copy to folder、assign/clear categories、mark as task（含 interval）、delete、desktop alert 與 stop processing more rules。
- [ ] 既有 rule 含 Rules object model 無法建立的特殊條件或動作時，仍可回推 snapshot，但 `canModifyDefinition = false`；AddIn 不得嘗試以自訂掃描或其他 automation 假裝支援。
- [ ] 任何 rule 變更後都呼叫 `Rules.Save(false)` 或等效流程保存，並回推最新 `PushRules(rules)`。
- [ ] `Rules.Save` 失敗時回 `ReportCommandResult(success=false)`，message 可診斷且不含敏感資料。

驗收：

- [ ] 回推的 rules 包含 rule name、enabled、order、conditions、actions、exceptions。
- [ ] 可新增一條 subject/sender/recipient/category/attachment/importance 條件搭配 move/copy/category/task/delete/alert/stop 動作的 receive rule。
- [ ] 可啟用/停用、刪除既有 rule。
- [ ] 對 `canModifyDefinition = false` 的特殊 rule，不應送完整 definition 修改。

### 11. Chat

AddIn 送 chat 必須使用 `/hub/outlook-addin` 的 SignalR method，不要再用 HTTP `/api/outlook/chat`。

- [ ] AddIn 要送 chat message 時 invoke `SendChatMessage(message)`。
- [ ] `message.source` 建議填 `outlook`。
- [ ] `message.text` 填入要顯示的訊息。
- [ ] 不需要自行呼叫其他 SignalR hub 或 HTTP endpoint。

驗收：

- [ ] AddIn invoke `SendChatMessage` 成功。
- [ ] AddIn code 不再呼叫 HTTP `/api/outlook/chat`。

### 12. Folder 建立與刪除

- [ ] AddIn 收到 `create_folder`。
- [ ] 用 `parentFolderPath` 找到 parent folder。
- [ ] 建立 `name` 指定的新 folder。
- [ ] 用 folder 增量同步回推 folder 變更。
- [ ] AddIn 收到 `delete_folder`。
- [ ] 用 `folderPath` 找到 folder object，但不得永久刪除 folder。
- [ ] `delete_folder` 必須將 folder 移到同一個 Outlook store / mailbox 的 default Deleted Items folder。
- [ ] Deleted Items 目的 folder 透過 `Store.GetDefaultFolder(olFolderDeletedItems)` 或等效 Outlook default folder identity 定位，不依賴 `Deleted Items`、`刪除的郵件` 或其他本地化 folder name。
- [ ] 若目標是 store root、hidden/system folder 或 Outlook object model 拒絕 move，不得永久刪除；只回報實際 automation 失敗診斷。
- [ ] 若目標已經位於 default Deleted Items folder 或其子層，仍不得永久刪除；只回報實際 move / no-op 結果。
- [ ] 用 folder 增量同步回推 source parent、Deleted Items folder 與被移動 folder tree 的變更。
- [ ] 若目前 mail list 指向已移動 folder，回推 `PushMails` 清掉或更新畫面。

驗收：

- [ ] 新增子 folder 後，folder snapshot 包含新 folder。
- [ ] `delete_folder` 後，folder snapshot 顯示該 folder 已位於 default Deleted Items folder 底下，而不是從 Outlook 永久消失。

## 常見失敗對照

| 現象 | 優先檢查 |
| --- | --- |
| 已讀/未讀出現 missing mail id | `PushMails` 是否填 `MailItemDto.id` |
| `move_mail` 沒有執行 | mail 是否有 `id`，以及 destination folder path 是否可解析 |
| AddIn 收不到 command | AddIn 是否有 SignalR connection |
| Category 空白 | AddIn 是否處理 `fetch_categories` 並 `PushCategories` |
| Flag 修改沒效果 | AddIn 是否處理 `update_mail_properties` 的 flag 欄位並儲存 item |
| Calendar 空白 | AddIn 是否處理 `fetch_calendar` 的 `startDate/endDate` 並 `PushCalendar` |
| AddIn chat 沒送出 | 是否 invoke `/hub/outlook-addin` 的 `SendChatMessage`，而不是 HTTP `/api/outlook/chat` |
| Folder 沒分 OST/PST | `OutlookStoreDto` 是否正確填入 `storeId`、`storeKind` 與 `rootFolderPath` |

## 需要時查看的官方文件

- Outlook Stores / PST / OST：
  - `NameSpace.Stores`: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.stores
  - `Store.GetRootFolder`: https://learn.microsoft.com/en-us/office/vba/api/outlook.store.getrootfolder
  - `Store.FilePath`: https://learn.microsoft.com/office/vba/api/Outlook.store.filepath
  - `Store.ExchangeStoreType`: https://learn.microsoft.com/en-us/office/vba/api/outlook.store.exchangestoretype
- Mail identity / lookup：
  - `MailItem.EntryID`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.entryid
  - `NameSpace.GetItemFromID`: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.getitemfromid
- Mail read / move：
  - `MailItem.UnRead`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.unread
  - `MailItem.Move`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.move
- Folder：
  - `Folder`: https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
  - `Folders.Add`: https://learn.microsoft.com/en-us/office/vba/api/outlook.folders.add
- Category：
  - `MailItem.Categories`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.categories
  - `Categories.Add`: https://learn.microsoft.com/en-us/office/vba/api/outlook.categories.add
- Flag / task：
  - `MailItem.FlagRequest`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.flagrequest
  - `MailItem.MarkAsTask`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.markastask
  - `MailItem.ClearTaskFlag`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.cleartaskflag
  - `MailItem.TaskDueDate`: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.taskduedate
## Contract 對照

實作時請以 `signalr-contract.md` 作為 JSON / DTO 欄位準則。本文件只描述「要做哪些功能」與「怎樣算做對」；payload 範例、DTO 欄位速查與 SignalR method 名稱以 contract 文件為準。
