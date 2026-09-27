namespace OutlookIngest.Models;

public class EmailRecord
{
    public string Subject { get; set; } = "";
    public string From { get; set; } = "";
    public string To { get; set; } = "";
    public string Body { get; set; } = "";
    public DateTime ReceivedTime { get; set; }
    public bool IsRead { get; set; }
    public bool HasAttachments { get; set; }
    public string[] AttachmentNames { get; set; } = [];
    public string Folder { get; set; } = "";       // "Inbox" or "Sent"
    public string ThreadId { get; set; } = "";      // Shared across reply chain
}
