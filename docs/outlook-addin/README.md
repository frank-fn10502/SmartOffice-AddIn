# Outlook AddIn 實作者文件

本資料夾是給工作機 Outlook AddIn 實作者看的文件。這裡只描述 AddIn 需要遵守的 Outlook automation 與 SignalR contract；除非直接影響 AddIn 實作，否則不在此解釋外部 client、服務端流程、測試替身或開發機架構。

AddIn 的定位很薄：收到 command、呼叫 Outlook API、把結果轉成 DTO 回推。AddIn 不自行實作跨 command 排程、跨 folder 負載管理、資料合併或對外 API。

## 建議閱讀順序

1. `features-checklist.md`：AddIn 必須實作的 command、完成定義與驗收項目。
2. `signalr-contract.md`：SignalR method、command payload、DTO 欄位與 JSON 範例。
3. `outlook-references.md`：需要確認 Outlook / Office 2016 行為時再查看的官方文件入口。
4. `test-report.md`：工作機測到差異、錯誤或真實資料形狀時的回報格式。

## 使用原則

- checklist 是任務入口；contract 是欄位規格。
- 實作新 AddIn command 前，請先確認對應 Hub contract 已在 `../SmartOffice.Hub` 完成：Microsoft 官方文件可行性確認、HTTP request/fetch-result endpoint、SignalR command/result DTO、mock backend、Web UI 手動檢查路徑、基本 UI sanity check、Hub 端負載控管與文件都已對齊，且 Mock 環境可運作。AddIn 真實 Office automation 是第二階段，不應先於 Hub/mock contract 落地。
- 不相容舊版 AddIn contract；不要保留 `/api/outlook/poll`、`/api/outlook/push-*`、舊 chat HTTP endpoint 或沒有實際使用的 command handler。
- 不維護未使用功能。若目前 contract 沒列出、工作機也沒有實際需求，請刪除或不要新增。
- AddIn 不負責負載管理。除非 Microsoft 官方文件明確指出某個 Outlook API 呼叫方式可改善效能，否則優先採用最單純、可診斷、可中止的實作。
- 涉及 Outlook/VSTO/COM automation、Office UI thread、`COMException`、大量 mail/folder 枚舉、效能頓挫或 Outlook object model 行為判斷時，實作前必須先查 Microsoft 官方文件；官方資料不足時，才補查 Microsoft Q&A、Stack Overflow 或相關 issue 作為輔助。回報時要區分官方依據、社群經驗與仍需工作機實測的假設。
- 開始 VSTO 真實實作前，請重新查一次與該功能直接相關的官方文件與實際討論，因為 Hub/mock 階段只能證明 contract 與 UI workflow 可行，不能證明 Outlook COM 行為穩定。若採用社群討論中的 workaround，必須在回報中標示其不是官方 contract。
- AddIn 處理 Hub 傳來的大量資訊時，要盡量降低 Outlook COM/STA 壓力：列表與搜尋結果只回 metadata，不讀完整 body、全部 recipients 或附件檔名；不要在持有短生命週期 COM object 時等待 SignalR/HTTP I/O；自己取得的短生命週期 COM object 應在可控範圍釋放。
- AddIn 不應承擔 Hub 可以處理的負載管理。若發現 Hub 需要一次丟出大量 command 才能完成使用者意圖，請先回到 Hub contract 重新設計 batch/slice/paging/progress，而不是在 AddIn 端硬接大量 command。
- AddIn 不應用測試資料反推 Outlook object model 行為。
- 真實 mail body、folder name、PST path、category name 與 chat message 都可能含敏感 business data；回報時必須匿名化。
- 郵件列表採兩段式載入：`fetch_mails` 只回 metadata，不應載入或回推完整 `body` / `bodyHtml`；收到 `fetch_mail_body` 時才讀取該封內容。
- 附件採兩段式處理：`fetch_mail_attachments` 只回附件 metadata；收到 `export_mail_attachment` 時才匯出到約定路徑。AddIn 不負責開檔。
