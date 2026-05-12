# Office 2016 Add-in 線上文件

本文件只紀錄 Office 2016 AddIn 實作時可查的線上文件入口。SignalR payload 與 DTO 格式請看 `signalr-contract.md`；工作機實測資料、差異與錯誤回報格式請看 `test-report.md`。

最後確認日期：2026-05-12。

## 使用原則

- 優先使用 Microsoft Learn 官方文件。
- 第三方文章只能作為輔助，不應作為 AddIn contract 或 Outlook 行為依據。
- Office 2016 desktop 是主要目標環境；不要只看最新 API 文件就假設 Office 2016 可用。
- 如果工作機實測結果與文件描述不一致，或 Outlook API 行為會影響 AddIn mapping、檔案寫入或 DTO 欄位，請用 `test-report.md` 的格式回報。
- 除非 Microsoft 官方文件明確指出某個 API 呼叫方式可改善效能，否則 AddIn 應選擇最簡單的 Outlook object model 流程；不要為了預想的效能優化加入額外排程、快取或 legacy fallback。
- 遇到 `COMException`、`RPC_E_SERVERCALL_RETRYLATER`、Outlook busy、UI thread 卡頓、大量 mail/folder 枚舉或 COM object lifetime 問題時，修改前必須先查本文件的官方入口；若官方資料無法解釋實際錯誤，再查 Microsoft Q&A、Stack Overflow 或相關 issue 作為輔助，且回報時要明確標示社群資料只是經驗佐證。
- 任何效能修正都要先確認責任邊界：Hub 負責跨 command / 跨 folder 排程與負載管理；AddIn 負責在單一 command 內用可中止、可診斷、低 COM 壓力的方式處理 Outlook object model。

## VSTO / COM Add-in

Office 2016 desktop 深度整合通常會碰到 VSTO、COM automation 或 Outlook object model。這些文件最適合查詢 `Application`、`NameSpace`、`Folder`、`MailItem`、`Items` 等行為。

- [Office solutions development overview (VSTO)](https://learn.microsoft.com/en-us/visualstudio/vsto/office-solutions-development-overview-vsto?view=visualstudio)：VSTO Office solution 的總覽。
- [Threading support in Office](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022)：Office object model 不具 thread-safe；Office solution code 跑在 main UI thread，背景 thread 呼叫 Office object model 會經 STA marshaling，Office busy 時可能轉成 `COMException`。處理大量 Outlook automation 前必讀。
- [Limitation of Asynchronous Programming to Object Model](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/asynchronous-programming-to-object-model)：Office object model 的 asynchronous / reentrancy 限制；雖以 Excel 為例，但原則適用於避免從非序列化 callback 直接打 Office OM。
- [Marshal.ReleaseComObject(Object)](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject)：明確釋放 COM RCW 的官方行為與風險；只應釋放自己取得、生命週期清楚的短生命週期 COM object，避免釋放仍可能被其他程式碼使用的 singleton/shared RCW。
- [Outlook object model overview](https://learn.microsoft.com/en-us/visualstudio/vsto/outlook-object-model-overview?view=vs-2022)：Outlook VSTO 專案如何使用 Outlook object model。
- [Outlook VBA object model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)：Outlook object model 的 VBA 參考；VSTO C# 常需要把 VBA sample 翻成 C# interop。
- [NameSpace object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace)：MAPI root、default folders、store access。
- [Folder object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder)：Outlook folder 與 nested folders。
- [Folders object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folders)：同一層 folder collection。
- [Folder.Folders property (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.folders)：讀取子資料夾。
- [Store.GetRootFolder method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.store.getrootfolder)：從單一 Store root 列舉 folder tree；Microsoft 文件也指出這不同於 `NameSpace.Folders` 直接拿目前 profile 所有 stores 的 folders。
- [Store.GetDefaultFolder method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.store.getdefaultfolder)：依 `OlDefaultFolders` 回傳該 store 的 well-known folder；AddIn 應用 Outlook identity 判定 well-known folders，不用本地化顯示名稱猜測。
- [OlDefaultFolders enumeration (Outlook)](https://learn.microsoft.com/office/vba/api/Outlook.OlDefaultFolders)：包含 `olFolderSyncIssues`、`olFolderConflicts`、`olFolderLocalFailures`、`olFolderServerFailures` 等 Exchange special folders。
- [Folder.DefaultItemType property (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.defaultitemtype)：回報 folder 預設 Outlook item type；mail command 只應處理 `olMailItem` / `0` folder。
- [OlItemType enumeration (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.olitemtype)：`olMailItem = 0`、`olJournalItem = 4`、`olNoteItem = 5` 等。
- [Folder.PropertyAccessor property (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder.propertyaccessor)：讀取 Outlook object model 未直接 exposed 的 folder MAPI properties，例如 hidden/system flags。
- [Default folder is missing in Outlook and Outlook on the web](https://learn.microsoft.com/en-us/troubleshoot/outlook/user-interface/default-folder-is-missing)：Microsoft support 文件說明 `PR_ATTR_HIDDEN` 與 `PR_ATTR_SYSTEM` 這兩個 folder 屬性。
- [MailItem object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem)：郵件 item、subject、sender、body、received time 等欄位。
- [Application.AdvancedSearch method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearch)：非同步搜尋；scope 可含同一個 store 內的多個 folder，不能跨 store。Microsoft 文件提醒大量 simultaneous search 會造成顯著搜尋活動並影響 Outlook performance；AddIn 收到 `fetch_mail_search_slice` 時只處理 request 指定的單一 folder。
- [Application.AdvancedSearchComplete event (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.advancedsearchcomplete)：`AdvancedSearch` 完成事件，避免以 blocking loop 等待。
- [Search the Inbox for Items with Subject Containing Office](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/search-the-inbox-for-items-with-subject-containing-office)：Microsoft 的 Subject contains 範例，示範以 DASL `ci_phrasematch` 查詢 Subject 內含關鍵字；正式搜尋應參考這類 Outlook 內建搜尋流程。
- [Enumerating, Searching, and Filtering Items in a Folder](https://learn.microsoft.com/office/vba/outlook/How-to/Search-and-Filter/enumerating-searching-and-filtering-items-in-a-folder)：比較 `Items`、`Table`、`Selection` 等枚舉方式；Microsoft 說明 `Table` 是 lightweight rowset，適合大量枚舉與 filtering。若列表效能不足，應優先評估 `Folder.GetTable` / `Search.GetTable`，但需在工作機確認欄位 mapping。
- [Table object (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.table)：`Table` 是 read-only dynamic rowset；每列只含指定欄位，適合快速列舉 metadata；若需要修改 item，再用 EntryID 取回完整 item。
- [Folder.GetTable method (Outlook)](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mapifolder.gettable?view=outlook-pia)：從 folder 取得 filtered `Table`；可用 `Columns.RemoveAll` / `Columns.Add` 控制欄位，降低大量 mail 列表的 COM property 存取成本。
- [Items.Restrict method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.restrict)：在單一 folder items 內做條件篩選，適合搭配日期、分類、已讀狀態等條件縮小結果。Microsoft 文件指出 `Restrict` 適合大型 collection 先縮小結果，但也明確說明不能做 Subject contains；文字 contains 請優先使用 Outlook 內建搜尋 / DASL content index。
- [Managing Rules in the Outlook Object Model](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/rules/managing-rules-in-the-outlook-object-model)：官方說明 Rules object model 支援 programmatic adding、editing、deleting rules；也說明 `Store.GetRules`、`Rules.Create`、`Rules.Remove`、`Rules.Save`、`Rule.Enabled` 與 `Rule.ExecutionOrder` 的行為。
- [Specifying Rule Conditions](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/rules/specifying-rule-conditions)：列出哪些 rule conditions 可由 object model 建立；特殊條件只能列舉或啟用/停用，不應承諾可完整建立。
- [Specifying Rule Actions](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/rules/specifying-rule-actions)：列出哪些 rule actions 可由 object model 建立；例如 move/copy、assign category、mark as task 與 stop processing more rules 可建立，run script、server reply、print 等不可建立。
- [Rules.Save method (Outlook)](https://learn.microsoft.com/en-us/office/vba/api/outlook.rules.save)：保存 rules collection；官方文件提醒 slow Exchange connection 可能昂貴，且不相容或定義不完整的 rule 會造成 save error。

## Office JavaScript Add-in / Office.js

如果工作機 Add-in 是 Office.js 或混合架構，必須先查 Office 2016 支援的 requirement set。

- [Requirements for running Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/add-in-requirements)：Office Add-in 的 client / server / Outlook account 需求。
- [Office versions and requirement sets](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-versions-and-requirement-sets)：不同 Office 版本可用 API 的判斷方式。
- [Office Common API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/office-add-in-requirement-sets)：Common API requirement set 清單。
- [Outlook JavaScript API requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)：Outlook `Mailbox` requirement set 與 manifest 宣告方式。
- [Outlook add-ins overview](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/read-scenario)：Outlook add-in activation、read / compose mode 與支援帳號。
- [Specify Office applications and API requirements with the add-in only manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)：使用 manifest 限制 host 與 API requirement。
