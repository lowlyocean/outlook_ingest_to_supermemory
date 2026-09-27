namespace OutlookIngest;

public class AppSettings
{
    public string? ApiKey { get; set; }
    public string ServerUrl { get; set; } = "http://localhost:6767";
    public string ContainerTag { get; set; } = "outlook_emails";
    public int BatchSize { get; set; } = 100;
}
