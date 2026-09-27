using System.Security.Cryptography;
using System.Text;
using OutlookIngest.Models;

namespace OutlookIngest;

public class ThreadCollator
{
    /// <summary>
    /// Assigns a thread ID to each email based on subject + sender hash.
    /// Emails in the same reply chain get the same ThreadId.
    /// </summary>
    public string ComputeThreadId(EmailRecord email)
    {
        // Simple heuristic: hash of subject + first sender
        var hash = Convert.ToHexString(
            SHA256.HashData(Encoding.UTF8.GetBytes($"{email.Subject}|{email.From}"))
        );
        return $"thread_{hash.Substring(0, 8)}";
    }

    /// <summary>
    /// Assigns ThreadId to all emails and returns them sorted by ReceivedTime.
    /// Each email remains a separate document — no body merging.
    /// </summary>
    public List<EmailRecord> Collate(IEnumerable<EmailRecord> inboxEmails, IEnumerable<EmailRecord> sentEmails)
    {
        var allEmails = new List<EmailRecord>();

        foreach (var email in inboxEmails)
        {
            email.ThreadId = ComputeThreadId(email);
            email.Folder = "Inbox";
            allEmails.Add(email);
        }

        foreach (var email in sentEmails)
        {
            email.ThreadId = ComputeThreadId(email);
            email.Folder = "Sent";
            allEmails.Add(email);
        }

        // Sort oldest-first (required by Supermemory)
        return allEmails.OrderBy(e => e.ReceivedTime).ToList();
    }
}
