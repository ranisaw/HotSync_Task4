namespace HotsyncApp.Models;

/// <summary>
/// Represents a single customer record retrieved from the SharePoint list.
/// Field names match the internal SharePoint column names exactly.
/// Nullable fields reflect that SharePoint items may not have every field populated.
/// </summary>
public sealed class CustomerRecord
{
    /// <summary>SharePoint list item ID — used as the unique sync key.</summary>
    public int SharePointId { get; set; }

    /// <summary>Customer / company name.</summary>
    public string? Name { get; set; }

    /// <summary>Customer's physical or registered address.</summary>
    public string? Address { get; set; }

    /// <summary>Number of oil rigs operated by the customer.</summary>
    public int? NumberOfRigs { get; set; }

    /// <summary>Number of jack-up rigs operated by the customer.</summary>
    public int? NumberOfJackUps { get; set; }

    /// <summary>Number of Mobile Offshore Drilling Units (MODUs).</summary>
    public int? NumberOfModus { get; set; }

    /// <summary>Comma-separated list of customer site URLs.</summary>
    public string? SiteUrls { get; set; }

    /// <summary>Timestamp of the last SharePoint modification (used for change detection).</summary>
    public DateTime? LastModified { get; set; }
}
