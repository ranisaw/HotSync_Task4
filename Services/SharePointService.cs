using HotsyncApp.Configuration;
using HotsyncApp.Helpers;
using HotsyncApp.Models;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;

namespace HotsyncApp.Services;

/// <summary>
/// Authenticates against Microsoft Entra ID (client credentials / app-only)
/// and retrieves all customer records from a SharePoint list via the
/// MICROSOFT GRAPH API — not the legacy SharePoint REST API.
///
/// Why Graph instead of SharePoint REST?
///   SharePoint REST API rejects app-only tokens with "Unsupported app only token"
///   when the site collection has app-only access disabled (tenant default).
///   Microsoft Graph API does not have this restriction and is the modern,
///   Microsoft-recommended approach for app-only SharePoint access.
///
/// Required App Registration permission (Application type, admin-consented):
///   Sites.Read.All  (Microsoft Graph)
///
/// Pagination: Graph returns up to 200 items per page and provides an
///   @odata.nextLink URL when more pages exist. This service follows all pages.
/// </summary>
public sealed class SharePointService
{
    private readonly SharePointConfig _spConfig;
    private readonly AzureAdConfig _adConfig;
    private readonly HttpClient _httpClient;

    // Graph API base URL — always this fixed value.
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    // Graph scope for app-only access — always this fixed value.
    private const string GraphScope = "https://graph.microsoft.com/.default";

    public SharePointService(AppConfig config, HttpClient httpClient)
    {
        _spConfig = config.SharePoint;
        _adConfig = config.AzureAd;
        _httpClient = httpClient;
    }

    // -------------------------------------------------------------------------
    // Authentication
    // -------------------------------------------------------------------------

    /// <summary>
    /// Acquires a Microsoft Graph bearer token from Entra ID using
    /// the OAuth 2.0 Client Credentials flow (app-only, no user sign-in).
    /// MSAL handles token caching and renewal internally.
    /// </summary>
    private async Task<string> AcquireGraphTokenAsync(CancellationToken ct)
    {
        Logger.Info("Acquiring Microsoft Graph access token from Entra ID...");

        var app = ConfidentialClientApplicationBuilder
            .Create(_adConfig.ClientId)
            .WithTenantId(_adConfig.TenantId)
            .WithClientSecret(_adConfig.ClientSecret)
            .Build();

        try
        {
            var result = await app
                .AcquireTokenForClient(new[] { GraphScope })
                .ExecuteAsync(ct);

            Logger.Success($"Graph token acquired. Expires: {result.ExpiresOn:yyyy-MM-dd HH:mm:ss} UTC");
            return result.AccessToken;
        }
        catch (MsalServiceException ex)
        {
            throw new InvalidOperationException(
                $"Entra ID authentication failed.\n" +
                $"Verify TenantId, ClientId and ClientSecret are correct, " +
                $"and that the App Registration has 'Sites.Read.All' (Graph) application permission with admin consent.\n" +
                $"MSAL error code: {ex.ErrorCode}\n" +
                $"MSAL message: {ex.Message}", ex);
        }
    }

    // -------------------------------------------------------------------------
    // Site resolution
    // -------------------------------------------------------------------------

    /// <summary>
    /// Resolves the SharePoint site URL to a Graph site ID.
    /// Graph requires the internal site ID (not the friendly URL) for list queries.
    ///
    /// Example site ID format: "m3p7fgr.sharepoint.com,abc123...,def456..."
    /// </summary>
    private async Task<string> ResolveSiteIdAsync(string token, CancellationToken ct)
    {
        // Extract host and site path from the configured SiteUrl.
        // e.g. "https://m3p7fgr.sharepoint.com/sites/Fennex-Onboarding-Rani"
        //   → host = "m3p7fgr.sharepoint.com"
        //   → path = "/sites/Fennex-Onboarding-Rani"
        var uri = new Uri(_spConfig.SiteUrl.TrimEnd('/'));
        var host = uri.Host;
        var path = uri.AbsolutePath;

        var url = $"{GraphBaseUrl}/sites/{host}:{path}";
        Logger.Info($"Resolving SharePoint site ID via Graph: {url}");

        var json = await GraphGetAsync(url, token, ct);
        var parsed = JObject.Parse(json);

        var siteId = parsed["id"]?.ToString()
            ?? throw new InvalidOperationException(
                $"Could not resolve site ID from Graph response. " +
                $"Response: {json[..Math.Min(json.Length, 500)]}");

        Logger.Success($"Site ID resolved: {siteId}");
        return siteId;
    }

    // -------------------------------------------------------------------------
    // List discovery
    // -------------------------------------------------------------------------

    /// <summary>
    /// Finds the Graph list ID for a list by its display title.
    /// Graph list queries require the internal list ID, not the display name.
    /// </summary>
    private async Task<string> ResolveListIdAsync(string siteId, string token, CancellationToken ct)
    {
        var url = $"{GraphBaseUrl}/sites/{siteId}/lists?$select=id,displayName";
        Logger.Info($"Searching for list '{_spConfig.ListTitle}' in site...");

        var json = await GraphGetAsync(url, token, ct);
        var parsed = JObject.Parse(json);
        var lists = parsed["value"] as JArray ?? [];

        var matched = lists.FirstOrDefault(l =>
            string.Equals(
                l["displayName"]?.ToString(),
                _spConfig.ListTitle,
                StringComparison.OrdinalIgnoreCase));

        if (matched == null)
        {
            // Log available lists to help the user correct the ListTitle config.
            var available = string.Join(", ", lists.Select(l => $"\"{l["displayName"]}\""));
            throw new InvalidOperationException(
                $"List '{_spConfig.ListTitle}' not found in the SharePoint site.\n" +
                $"Available lists: {available}\n" +
                $"Check the 'ListTitle' value in appsettings.json.");
        }

        var listId = matched["id"]!.ToString();
        Logger.Success($"List found. ID: {listId}");
        return listId;
    }

    // -------------------------------------------------------------------------
    // Main public method
    // -------------------------------------------------------------------------

    /// <summary>
    /// Retrieves all customer records from the configured SharePoint list via Graph.
    /// Handles pagination transparently — follows @odata.nextLink until all pages consumed.
    /// </summary>
    public async Task<List<CustomerRecord>> GetAllCustomersAsync(CancellationToken ct)
    {
        var token = await AcquireGraphTokenAsync(ct);

        // Step 1: Resolve site and list IDs.
        var siteId = await ResolveSiteIdAsync(token, ct);
        var listId = await ResolveListIdAsync(siteId, token, ct);

        // Step 2: Fetch all list items with field expansion.
        // Graph returns list item metadata at the top level and custom fields
        // under the "fields" property. We select only what we need.
        var customers = new List<CustomerRecord>();

        // Request up to 200 items per page (Graph maximum for list items).
        // $expand=fields brings in custom column values.
        // $select on fields narrows the payload to only the columns we map.
        var fieldSelect = "id,Title,Address,NumberOfRigs,NumberOfJackUps,NumberOfMODUs,SiteURLs,Modified";
        var currentUrl = $"{GraphBaseUrl}/sites/{siteId}/lists/{listId}/items" +
                         $"?expand=fields($select={fieldSelect})&$top=200";

        var pageNumber = 0;
        Logger.Info($"Fetching items from list '{_spConfig.ListTitle}' via Microsoft Graph...");

        while (!string.IsNullOrEmpty(currentUrl))
        {
            ct.ThrowIfCancellationRequested();
            pageNumber++;
            Logger.Info($"  Page {pageNumber}: {currentUrl[..Math.Min(currentUrl.Length, 100)]}...");

            var json = await GraphGetAsync(currentUrl, token, ct);
            var parsed = JObject.Parse(json);

            var items = parsed["value"] as JArray ?? [];
            Logger.Info($"  Page {pageNumber}: received {items.Count} item(s).");

            foreach (var item in items)
            {
                try
                {
                    customers.Add(MapToCustomerRecord(item));
                }
                catch (Exception ex)
                {
                    Logger.Warn($"  Could not map item id={item["id"]} — skipping. Reason: {ex.Message}");
                }
            }

            // Follow next page link if present.
            currentUrl = parsed["@odata.nextLink"]?.ToString() ?? string.Empty;
        }

        Logger.Success($"Graph fetch complete. Total records retrieved: {customers.Count}");
        return customers;
    }

    // -------------------------------------------------------------------------
    // HTTP helper
    // -------------------------------------------------------------------------

    /// <summary>
    /// Executes a GET request against the Graph API with the bearer token.
    /// Throws a clear exception on any non-success HTTP status.
    /// </summary>
    private async Task<string> GraphGetAsync(string url, string token, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization =
            new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.ParseAdd("application/json");

        HttpResponseMessage response;
        try
        {
            response = await _httpClient.SendAsync(request, ct);
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException(
                $"Network error calling Graph API: {url}\nDetail: {ex.Message}", ex);
        }

        var body = await response.Content.ReadAsStringAsync(ct);

        if (!response.IsSuccessStatusCode)
        {
            // Try to extract the Graph error message for a clearer error.
            string detail;
            try
            {
                var err = JObject.Parse(body);
                detail = err["error"]?["message"]?.ToString() ?? body;
            }
            catch
            {
                detail = body;
            }

            throw new InvalidOperationException(
                $"Microsoft Graph API returned HTTP {(int)response.StatusCode} ({response.ReasonPhrase}).\n" +
                $"URL: {url}\n" +
                $"Error: {detail[..Math.Min(detail.Length, 800)]}");
        }

        return body;
    }

    // -------------------------------------------------------------------------
    // Field mapping
    // -------------------------------------------------------------------------

    /// <summary>
    /// Maps a Graph list item JSON object to a <see cref="CustomerRecord"/>.
    ///
    /// Graph structure for a list item:
    /// {
    ///   "id": "1",                      ← string item ID (top level)
    ///   "fields": {
    ///     "Title": "...",               ← custom field values
    ///     "Address": "...",
    ///     ...
    ///   }
    /// }
    ///
    /// All field reads are null-safe. If a field is missing or null in SharePoint,
    /// the corresponding C# property is set to null rather than throwing.
    /// </summary>
    private static CustomerRecord MapToCustomerRecord(JToken item)
    {
        // Graph returns item ID as a string at the top level.
        var idStr = item["id"]?.ToString()
            ?? throw new InvalidOperationException("Graph item has no 'id' field.");

        if (!int.TryParse(idStr, out var sharePointId))
            throw new InvalidOperationException($"Graph item 'id' is not an integer: '{idStr}'.");

        // Custom field values are nested under "fields".
        var fields = item["fields"]
            ?? throw new InvalidOperationException($"Graph item id={idStr} has no 'fields' object.");

        return new CustomerRecord
        {
            SharePointId    = sharePointId,
            Name            = fields["Title"]?.Value<string>(),
            Address         = fields["Address"]?.Value<string>(),
            NumberOfRigs    = TryParseNullableInt(fields["NumberOfRigs"]),
            NumberOfJackUps = TryParseNullableInt(fields["NumberOfJackUps"]),
            NumberOfModus   = TryParseNullableInt(fields["NumberOfMODUs"]),
            SiteUrls        = fields["SiteURLs"]?.Value<string>(),
            LastModified    = fields["Modified"]?.Value<DateTime?>()
        };
    }

    /// <summary>
    /// Safely parses an integer from a JToken that may be null, empty, or a string.
    /// </summary>
    private static int? TryParseNullableInt(JToken? token)
    {
        if (token == null || token.Type == JTokenType.Null)
            return null;

        var raw = token.ToString();
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        return int.TryParse(raw, out var val) ? val : null;
    }
}
