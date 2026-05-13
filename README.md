# SmartOffice - POC

## Architecture

```
Web Browser (Dashboard)
    |  SignalR + REST
    v
SmartOffice.Hub (.NET 8 Web API)
    |  Long-polling REST
    v
Outlook Add-in (.NET Framework 4.8 VSTO)
    |
    v
Outlook COM API
```

**Flow: Web requests data from Outlook (not the other way around)**

1. Web UI clicks "Fetch Mails" or "Load Folders"
2. Hub queues a command
3. Outlook Add-in long-polls Hub, picks up command
4. Outlook reads data via COM API, pushes results back to Hub
5. Hub broadcasts to Web UI via SignalR

## Features

### Web UI (`http://localhost:2805`)
- **Folder Browser** - View all Outlook folders/subfolders with item counts. Click a folder to select it.
- **Mail Viewer** - Fetch mails with filters:
  - Time range: Today / Last 7 days / Last 30 days
  - Count: 10 / 20 / 30 / 100 mails
  - Expandable/collapsible mail body
- **Bidirectional Chat** - Messages sent from Web appear in Outlook, and vice versa

### Outlook Add-in
- **Chat Pane** - Toggle via Ribbon button. Shows all chat messages from both Web and Outlook.
- **Background Service** - Long-polls Hub for commands (fetch folders, fetch mails)

## How to Run

### Option 1: Multiple Startup Projects (Recommended)

1. Open `SmartOffice.sln` in Visual Studio
2. Right-click Solution > **Configure Startup Projects...**
3. Set both `SmartOffice.Hub` and `OutlookAddIn` to **Start**
4. Press **F5**
5. Browser opens `http://localhost:2805`
6. Outlook opens with Add-in loaded

### Option 2: Run Separately

**Hub:**
```bash
cd SmartOffice.Hub
dotnet run
```

**Outlook Add-in:**
- Set OutlookAddIn as startup project in VS, press F5

## Usage

1. Open `http://localhost:2805` in browser
2. Click **"Load Folders"** to see your Outlook folder structure
3. Click a folder to select it
4. Choose time range and mail count
5. Click **"Fetch Mails"** - mails appear after Outlook processes the request
6. Click any mail header to expand/collapse the body
7. Use the Chat section to send messages between Web and Outlook

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/outlook/request-mails` | Web requests mails (queues command) |
| POST | `/api/outlook/request-folders` | Web requests folders (queues command) |
| GET | `/api/outlook/mails` | Get cached mails |
| GET | `/api/outlook/folders` | Get cached folders |
| POST | `/api/outlook/chat` | Send chat message |
| GET | `/api/outlook/chat` | Get chat history |
| GET | `/api/outlook/poll` | Outlook long-polls for commands |
| POST | `/api/outlook/push-mails` | Outlook pushes mail data |
| POST | `/api/outlook/push-folders` | Outlook pushes folder data |
| SignalR | `/hub/notifications` | Real-time notifications |

## Port Configuration

Hub runs on `http://localhost:2805` (configured in `SmartOffice.Hub/Properties/launchSettings.json`).

If you change the port, also update `OutlookAddIn/Clients/HubClient.cs` > `BaseUrl`.
