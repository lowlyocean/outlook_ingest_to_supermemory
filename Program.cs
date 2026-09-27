using OutlookIngest;
using OutlookIngest.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Office.Interop.Outlook;

// 1. Load configuration from appsettings.json + env var overrides
var config = LoadConfig(args);

// 2. Connect to Outlook
using var outlookClient = new OutlookClient();
var store = outlookClient.GetFirstStore();
Console.WriteLine($"Using store: {store.StoreID}");

// 3. Get inbox + sent items
var inboxItems = outlookClient.GetInboxItems(store);
Console.WriteLine($"Found {inboxItems.Count()} inbox items");

var sentItems = outlookClient.GetSentItems(store);
Console.WriteLine($"Found {sentItems.Count()} sent items");

// 4. Convert to EmailRecords
var inboxEmails = inboxItems
    .Select(item => new EmailRecord
    {
        Subject = item.Subject ?? "(no subject)",
        From = item.SenderName ?? item.SenderEmailAddress ?? item.To,
        To = item.To,
        Body = item.Body ?? "",
        ReceivedTime = item.ReceivedTime,
        IsRead = false,
        HasAttachments = item.Attachments.Count > 0,
        AttachmentNames = item.Attachments.Cast<Attachment>()
            .Select(a => a.FileName).ToArray(),
        Folder = "Inbox"
    })
    .ToList();

var sentEmails = sentItems
    .Select(item => new EmailRecord
    {
        Subject = item.Subject ?? "(no subject)",
        From = item.SenderName ?? item.SenderEmailAddress ?? item.To,
        To = item.To,
        Body = item.Body ?? "",
        ReceivedTime = item.ReceivedTime,
        IsRead = true,
        HasAttachments = item.Attachments.Count > 0,
        AttachmentNames = item.Attachments.Cast<Attachment>()
            .Select(a => a.FileName).ToArray(),
        Folder = "Sent"
    })
    .ToList();

Console.WriteLine($"Converted {inboxEmails.Count} inbox + {sentEmails.Count} sent emails");

// 5. Collate by thread — assign ThreadId, sort by ReceivedTime
var collator = new ThreadCollator();
var collatedEmails = collator.Collate(inboxEmails, sentEmails);
Console.WriteLine($"Collated into {collatedEmails.Count} documents");

// 6. Sort oldest-first and ingest into Supermemory
var supermemoryClient = new SupermemoryClient(
    config.ApiKey, config.ServerUrl, config.ContainerTag, config.BatchSize);
await supermemoryClient.BatchIngest(collatedEmails);

Console.WriteLine("Ingestion complete.");

static AppSettings LoadConfig(string[] args)
{
    var builder = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
        .AddEnvironmentVariables();

    var config = builder.Build();

    var apiKey = config.GetValue<string>("ApiKey")
        ?? throw new InvalidOperationException("API key not configured. Set SUPERMEMORY_API_KEY or add ApiKey to appsettings.json");
    var serverUrl = config.GetValue<string>("ServerUrl") ?? "http://localhost:6767";
    var containerTag = config.GetValue<string>("ContainerTag") ?? "outlook_emails";
    var batchSize = config.GetValue<int?>("BatchSize") ?? 100;

    return new AppSettings
    {
        ApiKey = apiKey,
        ServerUrl = serverUrl,
        ContainerTag = containerTag,
        BatchSize = batchSize
    };
}
