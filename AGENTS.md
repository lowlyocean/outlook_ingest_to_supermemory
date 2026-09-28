# Agent Instructions — Outlook Email Ingestor

## Configuration

Settings are loaded from `appsettings.json` with environment variable overrides:

| Setting | Default | Env Var |
|---------|---------|---------|
| `ApiKey` | — | `SUPERMEMORY_API_KEY` |
| `ServerUrl` | `http://localhost:6767` | `SUPERMEMORY_SERVER_URL` |
| `ContainerTag` | `outlook_emails` | — |
| `BatchSize` | 100 | — |

## Running

```bash
dotnet run
```

## Architecture

```
Outlook (COM) → Email Collector → ThreadCollator → Supermemory API
```

## Files

| File | Purpose |
|------|---------|
| `Program.cs` | Entry point — load config, connect Outlook, orchestrate pipeline |
| `OutlookClient.cs` | COM Interop — connects to Outlook, reads Inbox + Sent Messages |
| `EmailConverter.cs` | Transforms MailItem → document content, strips quoted reply history |
| `ThreadCollator.cs` | Assigns ThreadId per reply chain, sorts oldest-first |
| `SupermemoryClient.cs` | HTTP client — batches 100 docs, retry with exponential backoff |
| `Models/EmailRecord.cs` | POCO for email data |
| `Models/IngestionResult.cs` | Batch response wrapper |
| `AppSettings.cs` | Configuration (API key, server URL, container tag, batch size) |
| `appsettings.json` | Config template |
| `OutlookIngest.csproj` | .NET 10.0 project file |
| `refs/` | Local COM interop DLL copies (Outlook PIA + Office core `office.dll`), referenced via HintPath, not tracked in git |

## Important Details

- **Store selection**: Sort `StoreID` length ascending, pick shortest (local mailbox)
- **Quoted history stripping**: Split inbox bodies on blank line + `From:` at line start, keep only the top part
- **Sent emails**: Keep intact (user's own contributions)
- **ThreadId**: `thread_{sha256(Subject|From).Substring(0,8)}` — same thread gets same ID
- **No body merging**: Each email stays as its own document
- **Sort order**: Oldest-first by `ReceivedTime` (required by Supermemory)
- **Batch size**: 100 documents per request
- **Container tag**: `outlook_emails`
- **Server URL**: Defaults to `http://localhost:6767`, configurable via `SUPERMEMORY_SERVER_URL` env var
- **Chunking**: The Supermemory server handles chunking automatically; no client-side chunking needed

## Adding New Features

- Configuration changes go in `AppSettings.cs` and `appsettings.json`
- Keep `Program.cs` as the thin orchestrator — delegate logic to dedicated components
