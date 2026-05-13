# Office 2016 工作機測試回報

本文件規範工作機實測資料、差異與錯誤要如何回傳。回報目標很單純：讓 AddIn contract、Outlook API mapping、檔案寫入或 DTO 欄位能依據真實 Office 2016 行為修正。

請避免把此文件寫成外部 client、測試替身或服務端設計回報。若回報內容需要開發機調整其他層，請只描述「哪個 AddIn command / DTO 欄位受到影響」與匿名化的實測證據。

回報要兼顧兩件事：不得洩漏真實 business data，但要讓開發機能重現 contract 或 mapping 問題。因此請用可分享的形狀、計數、enum、錯誤碼、匿名化 path segment 與已遮蔽 stack 摘要，不要貼完整郵件、完整 thread 或未遮蔽路徑。

## 何時需要回報

當工作機測試發現目前 AddIn contract 或 Outlook mapping 與 Office 2016 實際行為不一致時，請回傳一份 markdown。即使沒有錯誤，也可以回報代表性資料形狀、欄位限制、排序、空值、附件、folder path 或檔案系統限制。

常見回報情境：

- 格式與目前 contract 不符合，或 Office 2016 只能提供不同欄位。
- Outlook API 的實際行為會影響 AddIn 如何讀取、排序、過濾、寫入或匯出。
- 工作機環境限制會影響可用 API、資料量、編碼、檔名、路徑或權限。
- Microsoft 官方文件與工作機實測不一致。
- 社群討論或實際環境 workaround 影響 implementation decision；請附連結並標示它是社群經驗，不是官方依據。
- 某個效能最佳化有官方文件或可重現實測依據，且能讓 AddIn 呼叫 Outlook API 更簡單或更不阻塞。

不要只描述「壞掉」或「資料長這樣」；請提供足夠資訊讓開發機知道要調整哪個 command、DTO 欄位或 Outlook mapping，以及哪些資料已被匿名化。

建議檔名：

```text
workstation-report-YYYYMMDD-HHMM-command-type.md
```

## 必填內容

回報包必須包含：

- 測試日期、工作機代號、Office application、Office 版本與 bitness。
- AddIn 類型：VSTO / COM / Office.js / mixed。
- SmartOffice service 版本、API URL、測試 route。
- 收到的 `OutlookCommand` JSON。
- AddIn 呼叫的 Office API、物件類型與官方文件連結。
- 若引用 Microsoft Q&A、Stack Overflow 或 issue，請提供連結、摘要與採用原因，並標示為社群經驗。
- 轉換前的 Office 實測資料結構摘要。
- 實際 invoke 的 SignalR method 與 payload。
- SignalR invoke 結果或 exception 摘要。
- 預期格式與實際格式的差異；若沒有差異，請寫明這份資料用於校準哪個 AddIn contract 需求。
- 建議修正：改 AddIn mapping、改 DTO 欄位說明、調整檔案寫入策略，或移除未使用 / 舊版功能。
- 已匿名化的最小 sample。
- 若錯誤可能由 COM/STA/Outlook busy 造成，請回報 HRESULT、exception type、發生階段、是否可重試、當時 command 數量或 item 數量；不要貼含敏感資料的完整 stack。

不得包含：

- 真實 mail body。
- 客戶名稱、內部專案名稱、帳號、token。
- 完整 email thread。
- 未遮蔽的 folder name、mail address、PST/OST path 或 business data。

建議保留但要匿名化的資訊：

- Store kind、folder type、item count、batch size、command type、request id。
- Outlook API 名稱、HRESULT、exception type、是否發生在 UI thread 或 background worker。
- 欄位是否為空、是否 throw、是否需要 fallback；例如 `conversation=null`、`HTMLBody unavailable`、`Store.IsConversationEnabled=false`。
- 去識別化的 path shape，例如 `\\Mailbox - [redacted]\Inbox\[redacted-child]`。

## 回報範本

~~~markdown
# 工作機 Office 2016 測試回報

## Summary

- Date: 2026-04-29 09:35 +08:00
- Workstation: WS-REDACTED-01
- Office app: Outlook 2016
- Office version / build: 16.0.xxxxx.xxxxx
- Office bitness: 32-bit
- AddIn type: VSTO
- SmartOffice service version: abc1234
- API URL: http://dev-machine:2805
- Scenario: fetch_mails

## OutlookCommand

```json
{
  "id": "7f5d9b7d-1f86-49b5-a40e-5f2a3d1e9f88",
  "type": "fetch_mails",
  "mailsRequest": {
    "folderPath": "\\\\Mailbox - User\\Inbox",
    "receivedFrom": "2026-05-01T09:30:00+08:00",
    "receivedTo": "2026-05-08T09:30:00+08:00",
    "maxCount": 30
  }
}
```

## Office API Used

- API: `Application.Session.GetDefaultFolder`, `Folder.Items`, `MailItem.HTMLBody`
- NameSpace doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace
- Folder doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.folder
- MailItem doc: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem
- Community reference, if used: none

## Observed Office Data

Describe the object shape here. Redact sensitive values.

```json
{
  "folderPathObserved": "\\\\Mailbox - User\\Inbox",
  "itemType": "Outlook.MailItem",
  "subject": "[redacted]",
  "receivedTimeKind": "Local",
  "htmlBodyAvailable": false,
  "itemCount": 30,
  "batchSize": 5
}
```

## Error / Performance Notes

- Exception type: none
- HRESULT: none
- Retryable: unknown
- UI responsiveness: no visible freeze / brief freeze / noticeable freeze
- Command count sent by Hub for this user action: 1
- Outlook item count touched by AddIn: 30

## SignalR Payload Invoked

Method: `PushMails`

```json
[
  {
    "id": "[redacted Outlook EntryID]",
    "subject": "[redacted]",
    "sender": {
      "recipientKind": "sender",
      "displayName": "Sample Sender",
      "smtpAddress": "sender@example.invalid",
      "rawAddress": "sender@example.invalid",
      "addressType": "SMTP",
      "entryUserType": "olExchangeUserAddressEntry",
      "isGroup": false,
      "isResolved": true,
      "members": []
    },
    "toRecipients": [],
    "ccRecipients": [],
    "bccRecipients": [],
    "receivedTime": "2026-04-29T09:30:00+08:00",
    "body": "",
    "bodyHtml": "",
    "folderPath": "\\\\Mailbox - User\\Inbox"
  }
]
```

## Invoke Result

- Result: success
- Exception: none

## Difference From Current Contract

`bodyHtml` can be unavailable in this mailbox mode. `fetch_mails` should keep both `body` and `bodyHtml` empty, and body content should only be loaded by `fetch_mail_body`.

## Development Use

Use this sample to clarify `MailItemDto.bodyHtml` behavior in the AddIn contract. No legacy fallback or unused command should be added.

## Suggested Fix

- Keep `fetch_mails` metadata-only.
- Document that `bodyHtml` can be empty after `fetch_mail_body` if Outlook cannot provide HTML.
- Do not add a compatibility field unless the active contract explicitly needs it.

## Attachments / Extra Notes

- AddIn exception stack trace, if any.
- Screenshot filename, if needed and scrubbed.
- Any relevant Office Trust Center, Exchange, or account-mode constraint.
~~~

## 開發機收到回報後

工作機回報確認後，開發機應依影響範圍更新：

- `signalr-contract.md`：工作機傳送或接收格式改變時更新。
- `features-checklist.md`：AddIn 完成定義或驗收項目改變時更新。
- `outlook-references.md`：新增官方依據或 Office 2016 限制時更新。
- `Models/Dtos.cs`：欄位語意需要調整時更新；不要為舊版 AddIn 新增相容欄位。
- `Hubs/OutlookAddinHub.cs` 或 `Services/Stores.cs`：只有 SignalR method、command result 或 DTO 接收行為需要調整時更新。

不建議做法：

- 不要因工作機單次測試就 rename JSON field；若 contract 必須改，請同步移除舊欄位與舊 handler，不做雙軌相容。
- 不要把真實 mail body、folder name、PST/OST path 或客戶資訊 commit 到 repo。
- 不要用測試替身的假資料反向定義 Office 2016 真實格式。
- 不要因回報不是錯誤就忽略；代表性實測資料也可以是調整 AddIn mapping 與檔案輸出行為的依據。
- 不要在沒有官方文件或實測回報的情況下假設 Office.js 最新文件適用於 Office 2016。
- 不要為了預想的舊版相容保留未使用 command、HTTP endpoint 或 fallback。
