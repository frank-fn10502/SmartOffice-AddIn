# MailAutomation (Legacy COM 方案) 功能整理

> 此文件整理自 `MailAutomation/SmartMail` 專案，該專案使用 COM Interop 直接操作 Outlook 2016。
> 現已決定改用 VSTO Add-in 架構（SmartOffice），此處僅保留功能需求與設計參考。
> 遷移任何功能時，請先查 Microsoft 官方文件確認 Outlook/Office API 概念可行，再在 `SmartOffice.Hub` 建立目前 contract、mock backend、API/Web UI 驗證與文件；確認 Mock 與文件一致、基本 UI 可操作、Hub 不會把工作拆成大量 AddIn command 後，才修改 `SmartOffice/OutlookAddIn` 的 VSTO 真實實作。不要直接依本 Legacy 文件改 Add-in。

---

## 專案概述

| 項目 | 說明 |
|------|------|
| 專案名稱 | SmartMail.Agent |
| 類型 | WinForms (.NET 8) |
| Outlook 存取方式 | COM Interop (`Microsoft.Office.Interop.Outlook`) |
| 主要用途 | 讀取/Dump Outlook 收件匣、交談、資料夾結構 |

---

## 已實現功能

### 1. 收件匣讀取 (Inbox)

- 列出收件匣前 N 封信的 Meta 資訊（快速掃描，不讀 Body）
- 依 `EntryId` 讀取完整 `MailDocument`（含 Body + 附件資訊）
- 排序方式：ReceivedTime 由新到舊

### 2. 交談模式 (Conversation)

- 依任一封信的 EntryId 展開整個 Conversation
- 優先使用 `Conversation.GetTable()` API
- Fallback：依 `ConversationID` 在 Inbox 掃描
- 若 Conversation 只有 1 封信則略過

### 3. 資料夾列舉 (Folders)

- 列出所有 Store（含 Archive、PST）的完整資料夾樹
- 遞迴走訪所有子資料夾
- 輸出每個資料夾的 StoreName / FolderPath / EntryId

### 4. Dump 輸出

- **單封 Dump**：輸出為 txt，含完整 Meta + Body
- **Conversation Dump**：每個交談建一個子資料夾，含 index 概覽 + 各封信 txt
- **Folder Dump**：輸出所有資料夾清單（全部 + 按 Store 分檔）

---

## 資料模型

### MailId
| 欄位 | 用途 |
|------|------|
| EntryId | Outlook 本機操作用（GetItemFromID） |
| InternetMessageId | 全域去重、外部追蹤 |
| ConversationId | 交談分組 key |
| ConversationTopic | 人類可讀交談主題 |
| ConversationIndex | 交談排序 |

### MailMeta
| 欄位 | 用途 |
|------|------|
| Subject | 主旨 |
| ReceivedTime | 接收時間 |
| SentOn | 寄出時間 |
| Unread | 未讀狀態 |
| Importance | 重要性 (0/1/2) |
| Sensitivity | 機密性 |
| Categories | 分類標籤（`;` 分隔） |
| Flag | Follow Up 資訊 |
| From | 寄件者 |
| To | 收件者清單 |
| Cc | 副本清單 |

### MailFlagInfo
| 欄位 | 用途 |
|------|------|
| FlagStatus | None / Complete / Flagged |
| FlagRequest | 描述文字 |
| FlagDueBy | 到期日 |
| IsMarkedAsTask | 是否為 Task |
| FlagCompleteTime | 完成時間 |

### MailDocument
- Meta (MailMeta)
- Body (信件內文)
- Attachments (附件清單)

### MailAddress
| 欄位 | 用途 |
|------|------|
| Name | 顯示名稱 |
| Email | SMTP / EX 地址 |
| Raw | 原始字串 |

---

## 原始 TODO（需遷移到 SmartOffice）

1. 在 Outlook 加上分類 / Flag
2. 移動資料夾能力（移到 Archive / Server 資料夾）
3. Mail 與 Snapshot 同步問題
4. 同一科相同 Mail 的同步問題
5. YAML 設計
6. 新增 Category / Flag 能力
7. 在 Outlook 建立資料夾
8. 查看 Outlook 行事曆
9. 統計信件能力（需 DB）
10. 對外開放 RESTful Server 與 MCP Protocol

---

## 技術筆記

### COM 操作注意事項

- 所有 COM 物件使用後須 `Marshal.FinalReleaseComObject` 釋放
- `InternetMessageID` 某些環境下會 throw，需用 `PropertyAccessor` fallback
- 取 Flag 日期時 Outlook 可能回傳 `1/1/4501`（代表無效），需正規化為 null
- Conversation API 不一定可用，需準備 fallback 路徑
- 單封信讀取失敗不應影響整批操作

### 架構選擇說明

| 方案 | 優點 | 缺點 |
|------|------|------|
| COM Interop (本專案) | 功能最完整、可存取所有 MAPI | 依賴本機 Outlook、穩定性差、需手動 COM release |
| VSTO Add-in (SmartOffice) | 在 Office 進程內運行、官方支援、可用 Ribbon/TaskPane | 需 .NET Framework 4.8、部署較複雜 |

---

## 對 SmartOffice 的建議

1. **資料模型可直接復用**：`MailId`, `MailMeta`, `MailDocument`, `MailAddress`, `MailFlagInfo` 設計合理，可搬到 `SmartOffice.Contracts`
2. **Conversation 邏輯**：在 Add-in 內可直接存取 `MailItem.GetConversation()`，比 COM 外部呼叫更穩定
3. **Folder 操作**：Add-in 內可直接操作 `Application.Session.Stores`，不需繞路
4. **Dump 功能**：改由 SmartOffice.Hub 處理存檔/索引，Add-in 只負責傳送資料
