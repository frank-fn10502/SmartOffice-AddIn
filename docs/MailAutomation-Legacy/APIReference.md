# OutlookAgent API 參考（來自 MailAutomation）

> 來源：`MailAutomation/SmartMail/SmartMail.Agent/Outlook2016/OutlookAgent*.cs`

---

## OutlookAgent 類別

`sealed partial class OutlookAgent : IDisposable`

透過 COM Interop 操作 Outlook 2016 的核心代理類別。

---

## 公開 API

### ListInboxMeta(int maxCount) → List\<MailMeta\>

列出收件匣前 N 封信的 Meta。

- 按 ReceivedTime 降冪排列
- 只讀 Meta（不讀 Body），效能較佳
- 單封例外不影響整批

---

### ReadMailDocumentByEntryId(string entryId) → MailDocument

依 EntryId 讀取完整郵件內容。

- 含 Meta + Body + Attachments
- 若非 MailItem 則 throw

---

### ListConversationMetaByEntryId(string entryId, int maxCount = 500) → List\<MailMeta\>

依任一封信展開整個 Conversation。

- 優先用 `Conversation.GetTable()` API
- Fallback：依 ConversationID 掃描 Inbox
- 若只有 1 封信回傳空集合

---

### ListAllFolders() → List\<FolderEntry\>

列出所有 Store 的完整資料夾樹。

- 含 Archive、PST 等所有 Store
- 遞迴列出所有子資料夾
- 回傳 StoreName / FolderPath / EntryId

---

## 對應 SmartOffice API 映射建議

| Legacy COM API | SmartOffice Hub Endpoint (建議) |
|---|---|
| `ListInboxMeta(n)` | `GET /api/outlook/inbox?count=n` |
| `ReadMailDocumentByEntryId(id)` | `GET /api/outlook/mail/{entryId}` |
| `ListConversationMetaByEntryId(id)` | `GET /api/outlook/conversation/{entryId}` |
| `ListAllFolders()` | `GET /api/outlook/folders` |

在新架構中：
- **Add-in** 負責呼叫 Outlook COM 取得資料
- **Hub** 負責接收、處理、存儲、分析
- Add-in → Hub 走 HTTP/IPC
