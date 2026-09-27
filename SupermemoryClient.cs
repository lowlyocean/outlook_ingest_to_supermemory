using System.Net.Http.Json;
using System.Net.Http.Headers;
using OutlookIngest.Models;

namespace OutlookIngest;

public class SupermemoryClient
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private readonly string _serverUrl;
    private readonly string _containerTag;
    private readonly int _batchSize;
    private readonly EmailConverter _converter;

    public SupermemoryClient(string? apiKey, string serverUrl, string containerTag, int batchSize = 100)
    {
        _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
        _serverUrl = serverUrl.TrimEnd('/');
        _containerTag = containerTag;
        _batchSize = batchSize;
        _httpClient = new HttpClient();
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        _converter = new EmailConverter();
    }

    /// <summary>
    /// Batches documents for ingestion. Each email is a single document — the Supermemory
    /// server handles chunking automatically; no client-side chunking needed.
    /// </summary>
    public async Task<IngestionResult> BatchIngest(IEnumerable<EmailRecord> emails)
    {
        var sortedEmails = emails.OrderBy(e => e.ReceivedTime).ToList();
        var total = sortedEmails.Count;
        Console.WriteLine($"Ingesting {total} documents in batches of {_batchSize}");

        for (int i = 0; i < sortedEmails.Count; i += _batchSize)
        {
            var batchNum = i / _batchSize + 1;
            var batch = sortedEmails.Skip(i).Take(_batchSize).ToList();
            Console.WriteLine($"  Batch {batchNum}: {batch.Count} documents");

            var documents = batch.Select(email => new
            {
                content = _converter.Convert(email),
                customId = $"email_{email.ReceivedTime:yyyy-MM-ddTHH_mm_ss}_{Guid.NewGuid()}",
                documentDate = email.ReceivedTime.ToString("yyyy-MM-ddTHH:mm:ssZ"),
                metadata = new Dictionary<string, object>
                {
                    { "subject", email.Subject },
                    { "from", email.From },
                    { "to", email.To },
                    { "hasAttachments", email.HasAttachments },
                    { "folder", email.Folder },
                    { "threadId", email.ThreadId }
                }
            });

            var payload = new
            {
                containerTag = _containerTag,
                documents = documents.ToArray()
            };

            await RetryWithBackoff(() => SendBatch(payload, batchNum));
        }

        Console.WriteLine($"Ingestion complete: {total} documents processed");
        return new IngestionResult { TotalProcessed = total, Failed = 0, Success = total };
    }

    private async Task SendBatch(object payload, int batchNum)
    {
        var response = await _httpClient.PostAsJsonAsync(
            $"{_serverUrl}/v3/documents/batch", payload);

        if (!response.IsSuccessStatusCode)
        {
            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                throw new InvalidOperationException("Supermemory API key is invalid (401 Unauthorized)");
            }

            var body = await response.Content.ReadAsStringAsync();
            throw new InvalidOperationException(
                $"Batch {batchNum} failed: {response.StatusCode} - {body}");
        }

        var result = await response.Content.ReadFromJsonAsync<IngestionResult>();
        if (result?.Failed > 0)
        {
            throw new InvalidOperationException(
                $"{result.Failed} documents failed in batch {batchNum}");
        }
    }

    /// <summary>
    /// Retries with exponential backoff on 429 / 5xx responses.
    /// </summary>
    private async Task RetryWithBackoff(Func<Task> action)
    {
        int retries = 3;
        for (int i = 0; i < retries; i++)
        {
            try
            {
                await action();
                return;
            }
            catch (Exception ex) when (i < retries - 1 && IsRetryable(ex))
            {
                var delay = TimeSpan.FromMilliseconds(Math.Pow(2, i) * 1000);
                Console.WriteLine($"  Retry {i + 1}/{retries} after {delay.TotalMilliseconds}ms: {ex.Message}");
                await Task.Delay((int)delay.TotalMilliseconds);
            }
        }
    }

    private bool IsRetryable(Exception ex)
    {
        if (ex is InvalidOperationException ioEx)
        {
            if (ioEx.Message.Contains("429") || ioEx.Message.Contains("5"))
                return true;
        }
        if (ex is HttpRequestException httpEx && httpEx.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
            return true;
        return false;
    }
}
