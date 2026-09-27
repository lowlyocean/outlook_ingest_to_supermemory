namespace OutlookIngest.Models;

public class IngestionResult
{
    public int TotalProcessed { get; set; }
    public int Failed { get; set; }
    public int Success { get; set; }
}
