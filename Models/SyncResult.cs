namespace HotsyncApp.Models;

/// <summary>
/// Tracks the outcome of a single sync run for reporting purposes.
/// </summary>
public sealed class SyncResult
{
    public int TotalFetched { get; set; }
    public int Inserted { get; set; }
    public int Updated { get; set; }
    public int Skipped { get; set; }
    public int Errors { get; set; }
    public DateTime StartedAt { get; set; }
    public DateTime? CompletedAt { get; set; }

    public TimeSpan? Duration => CompletedAt.HasValue
        ? CompletedAt.Value - StartedAt
        : null;

    public bool HasErrors => Errors > 0;

    public override string ToString()
    {
        var dur = Duration.HasValue ? $"{Duration.Value.TotalSeconds:F1}s" : "in progress";
        return $"[SyncResult] Fetched={TotalFetched} | Inserted={Inserted} | Updated={Updated} | Skipped={Skipped} | Errors={Errors} | Duration={dur}";
    }
}
