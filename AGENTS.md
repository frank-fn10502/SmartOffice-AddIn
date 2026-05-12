# AGENTS.md

這份文件是 `SmartOffice/` 的 AI coding agent 入口。此專案是 Windows / Outlook / VSTO Add-in solution，和旁邊的 `SmartOffice.Hub/` 共同組成 SmartOffice PoC。

## 必讀限制

- 請使用繁體中文與使用者溝通；技術名詞、API name、file path、command、class name 與 JSON field 可保留英文。
- `SmartOffice/` 需要 Windows 主機、Visual Studio、Outlook、Office/VSTO runtime 才能完整編譯與測試。
- 目前這台主機預設只能測 `../SmartOffice.Hub`，不能可靠執行 Outlook Add-in。
- 修改本專案時，最終回覆請明確區分：
  - 已完成的程式碼或文件修改。
  - 已能在本機做的靜態檢查或純邏輯驗證。
  - 仍需 Windows 主機編譯與 Outlook 實測的項目。

## 專案內容

- `SmartOffice.sln`：Windows 主機上開啟的 solution。
- `OutlookAddIn/`：.NET Framework 4.8 VSTO Outlook Add-in。
- `OutlookAddIn/ThisAddIn*.cs`：Outlook lifecycle、folders、mails、rules、calendar、categories、readers、search、polling 等實作切分。
- `OutlookAddIn/HubClient.cs`、`SignalRClient.cs`：與 Hub 溝通的 client layer。
- `OutlookAddIn/ChatPane*.cs`：Outlook task pane UI。
- `OutlookAddIn/SmartOfficeRibbon.*`：Ribbon button 與 UI entry point。
- `OutlookAddIn.Tests/`：不依賴 COM/VSTO 的純邏輯測試，目前用 xUnit / .NET 8 測 `FolderFilter` 這類可抽離邏輯。
- `docs/outlook-addin/`：Outlook Add-in 實作者文件、SignalR contract、功能 checklist、官方文件入口與工作機測試回報格式。
- `docs/MailAutomation-Legacy/`：舊版 mail automation 文件，採用前請先比對目前 Hub contract。

## 與 SmartOffice.Hub 的關係

`../SmartOffice.Hub/` 是本地 Hub/API/Web UI/contract/mock 專案。Outlook Add-in 應把 Hub 視為命令與回報結果的中介，不應把 Web UI 或 mock 邏輯搬進 Add-in。

責任分工：

- Add-in 負責 Outlook COM/VSTO automation、讀取與修改本機 Outlook 資料、處理 Office UI 與 lifecycle。
- Hub 負責 HTTP API、SignalR endpoint、command routing、temporary state 與 Web UI 讀取資料。
- Web UI 負責檢視、手動 request、chat 與 diagnostics。

需要 contract 時，請讀：

- `docs/outlook-addin/README.md`
- `docs/outlook-addin/signalr-contract.md`
- `docs/outlook-addin/features-checklist.md`
- `../SmartOffice.Hub/docs/ai/protocols.md`

## 修改原則

- Office 2016 與受限企業環境是設計約束。避免引入需要新平台、新 runtime 或 cloud-only dependency 的做法。
- Outlook COM object 使用要小心釋放與例外處理；避免長時間阻塞 Outlook UI thread。
- 涉及 Outlook/VSTO/COM automation、Office UI thread、`COMException`、大量 mail/folder 枚舉、效能頓挫或 Outlook object model 行為判斷時，修改前必須先查 Microsoft 官方文件；若官方文件不足以解釋真實錯誤，再補查 Microsoft Q&A、Stack Overflow 或套件 issue 等討論區回報作為輔助。最終回覆要簡短列出採用的依據，並區分「官方依據」與「社群經驗」。
- 對效能修正要站在 Add-in 角度降低 COM/STA 負擔：避免列表階段讀完整 body、全部 recipients、附件檔名或跨 folder 自行排程；避免在持有短生命週期 Outlook COM object 時等待網路 I/O；每個自己取得的短生命週期 COM object 要在可控範圍釋放。
- 只有純邏輯才適合抽到 `OutlookAddIn.Tests/` 在非 Outlook 環境測試。
- 若要改 DTO、command type、SignalR method 或 route，請同步檢查 `../SmartOffice.Hub` 的 model、controller、hub 與文件。
- 不要將 Hub mock 行為當成 Outlook 真實行為；mock 只用來讓 Hub/Web UI 在本機可驗證。
- 請假設 mail body、subject、recipient、folder path、attachment name、calendar content 與 chat message 都可能是敏感 business data。

## 驗證期待

在目前主機：

- 可做程式碼閱讀、文件更新、純邏輯重構與測試規劃。
- 若本機 .NET SDK 可用，`OutlookAddIn.Tests` 這種純邏輯測試可能可以執行。
- 不應宣稱 VSTO Add-in 已完成編譯或 Outlook 實測。

在 Windows 主機：

- 用 Visual Studio 開啟 `SmartOffice.sln`。
- 編譯 `OutlookAddIn`。
- 啟動 Outlook 並確認 Add-in 載入。
- 搭配 `SmartOffice.Hub` 驗證 folders、mails、search、rules、categories、calendar、chat 與 command result。
