# 資料模型定義（來自 MailAutomation）

> 來源：`MailAutomation/SmartMail/SmartMail.Agent/Models/OutlookModels.cs`

---

## MailId

```csharp
public sealed class MailId
{
    public string EntryId { get; set; } = "";
    public string InternetMessageId { get; set; } = "";
    public string ConversationId { get; set; } = "";
    public string ConversationTopic { get; set; } = "";
    public string ConversationIndex { get; set; } = "";
}
```

---

## MailAddress

```csharp
public sealed class MailAddress
{
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public string Raw { get; set; } = "";
}
```

---

## MailFlagStatus

```csharp
public enum MailFlagStatus
{
    None = 0,
    Complete = 1,
    Flagged = 2,
}
```

---

## MailFlagInfo

```csharp
public sealed class MailFlagInfo
{
    public MailFlagStatus FlagStatus { get; set; }
    public string FlagRequest { get; set; } = "";
    public DateTime? FlagDueBy { get; set; }
    public bool IsMarkedAsTask { get; set; }
    public DateTime? FlagCompleteTime { get; set; }
}
```

---

## MailAttachmentInfo

```csharp
public sealed class MailAttachmentInfo
{
    public int Index { get; set; }        // 1-based
    public string FileName { get; set; } = "";
    public int Size { get; set; }         // bytes
}
```

---

## MailMeta

```csharp
public sealed class MailMeta
{
    public MailId Id { get; set; } = new MailId();
    public string Subject { get; set; } = "";
    public DateTime ReceivedTime { get; set; }
    public DateTime? SentOn { get; set; }
    public bool Unread { get; set; }
    public int Importance { get; set; }
    public int Sensitivity { get; set; }
    public string Categories { get; set; } = "";
    public MailFlagInfo Flag { get; set; } = new MailFlagInfo();
    public MailAddress From { get; set; } = new MailAddress();
    public List<MailAddress> To { get; set; } = new List<MailAddress>();
    public List<MailAddress> Cc { get; set; } = new List<MailAddress>();
}
```

---

## MailDocument

```csharp
public sealed class MailDocument
{
    public MailMeta Meta { get; set; }
    public MailBody Body { get; set; }
    public List<MailAttachmentInfo> Attachments { get; set; }
}
```

---

## FolderEntry

```csharp
public sealed class FolderEntry
{
    public string StoreName { get; set; }
    public string FolderName { get; set; }
    public string FolderPath { get; set; }
    public string EntryId { get; set; }
}
```
