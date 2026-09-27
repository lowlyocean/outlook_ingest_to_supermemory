# Outlook Email Ingestor

A C# console application that connects to Outlook Classic via COM Interop, reads inbox and sent emails, strips quoted reply history, and batches ingests them into Supermemory.

## Prerequisites

- Windows with Outlook 2016+ installed
- [.NET 10.0 SDK](https://dotnet.microsoft.com/download)
- A running Supermemory instance

## Configuration

Set the following environment variables:

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `SUPERMEMORY_API_KEY` | Yes | — | Your Supermemory API key |
| `SUPERMEMORY_SERVER_URL` | No | `http://localhost:6767` | Supermemory server base URL |

You can also set `SUPERMEMORY_SERVER_URL` to point at a self-hosted instance.

## Usage

```bash
export SUPERMEMORY_API_KEY="your-api-key"
cd outlook-ingest
dotnet run
```

## How It Works

1. **Connects** to the local Outlook instance via COM Interop
2. **Selects** the local mailbox store (shortest `StoreID`)
3. **Reads** both Inbox and Sent Messages folders
4. **Strips** quoted reply history from inbox emails (splits on blank line + `From:` at line start)
5. **Keeps** sent emails intact (your own contributions)
6. **Assigns** `ThreadId` per reply chain (hash of subject + sender)
7. **Sorts** all emails oldest-first (required by Supermemory)
8. **Batches** ingests 100 documents per request with exponential backoff retry

## Project Structure

```
outlook-ingest/
├── OutlookIngest.csproj          # .NET 10.0 project file
├── Program.cs                    # Entry point, orchestration
├── OutlookClient.cs              # COM Interop layer
├── EmailConverter.cs             # MailItem → Supermemory content
├── ThreadCollator.cs             # Thread ID assignment, oldest-first sort
├── SupermemoryClient.cs          # HTTP batch ingestion client
├── Models/
│   ├── EmailRecord.cs            # Email POCO
│   └── IngestionResult.cs        # Batch response wrapper
├── AppSettings.cs                # Configuration model
├── appsettings.json              # Config template
├── AGENTS.md                     # Agent instructions
├── README.md                     # This file
└── .gitignore
```

## Output Format

Each email ingests as a separate document with metadata:

```json
{
  "subject": "...",
  "from": "...",
  "to": "...",
  "hasAttachments": false,
  "folder": "Inbox" | "Sent",
  "threadId": "thread_abc12345"
}
```

All documents get the container tag `outlook_emails`.
