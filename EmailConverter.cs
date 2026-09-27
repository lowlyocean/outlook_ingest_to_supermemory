using System.Text.RegularExpressions;
using System.Text;
using OutlookIngest.Models;

namespace OutlookIngest;

public class EmailConverter
{
    /// <summary>
    /// Strips quoted reply history from an email body.
    /// Splits on "From:" appearing at the beginning of a line with a blank line preceding it.
    /// </summary>
    public string StripQuotedHistory(string body)
    {
        if (string.IsNullOrEmpty(body)) return body;

        // Split on blank line followed by "From:" at start of line
        var pattern = @"\r?\n\r?\n\s*From:";
        var parts = Regex.Split(body, pattern);

        // If we found a split, return only the top part (new contribution)
        if (parts.Length > 1)
        {
            return parts[0].Trim();
        }

        return body;
    }

    public string Convert(EmailRecord email)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Subject: {email.Subject}");
        sb.AppendLine($"From: {email.From}");
        sb.AppendLine($"To: {email.To}");
        sb.AppendLine($"Date: {email.ReceivedTime:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();
        sb.AppendLine("Body:");

        // Strip quoted history for inbox items (preserves user's own contributions from Sent folder)
        var body = email.Folder == "Inbox" ? StripQuotedHistory(email.Body) : email.Body;
        sb.AppendLine(body);
        return sb.ToString();
    }
}
