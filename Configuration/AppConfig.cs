namespace HotsyncApp.Configuration;

/// <summary>
/// Strongly-typed binding for Azure Active Directory (Entra ID) settings.
/// Values are loaded from appsettings.json or environment variables — never hardcoded.
/// </summary>
public sealed class AzureAdConfig
{
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
}

/// <summary>
/// SharePoint site and list configuration.
/// </summary>
public sealed class SharePointConfig
{
    public string SiteUrl { get; set; } = string.Empty;
    public string ListTitle { get; set; } = string.Empty;
}

/// <summary>
/// Azure SQL database configuration.
/// </summary>
public sealed class DatabaseConfig
{
    public string ConnectionString { get; set; } = string.Empty;
}

/// <summary>
/// Runtime sync options.
/// </summary>
public sealed class SyncConfig
{
    public int BatchSize { get; set; } = 100;
    public string LogLevel { get; set; } = "Information";
}

/// <summary>
/// Root application configuration model — all sections composed here.
/// </summary>
public sealed class AppConfig
{
    public AzureAdConfig AzureAd { get; set; } = new();
    public SharePointConfig SharePoint { get; set; } = new();
    public DatabaseConfig Database { get; set; } = new();
    public SyncConfig Sync { get; set; } = new();

    /// <summary>
    /// Validates that all required configuration values have been set.
    /// Throws an InvalidOperationException listing every missing field.
    /// </summary>
    public void Validate()
    {
        var missing = new List<string>();

        if (string.IsNullOrWhiteSpace(AzureAd.TenantId) || AzureAd.TenantId == "YOUR_TENANT_ID")
            missing.Add("AzureAd:TenantId");

        if (string.IsNullOrWhiteSpace(AzureAd.ClientId) || AzureAd.ClientId == "YOUR_CLIENT_APP_ID")
            missing.Add("AzureAd:ClientId");

        if (string.IsNullOrWhiteSpace(AzureAd.ClientSecret) || AzureAd.ClientSecret == "YOUR_CLIENT_SECRET")
            missing.Add("AzureAd:ClientSecret");

        if (string.IsNullOrWhiteSpace(SharePoint.SiteUrl) || SharePoint.SiteUrl.Contains("YOUR_TENANT"))
            missing.Add("SharePoint:SiteUrl");

        if (string.IsNullOrWhiteSpace(SharePoint.ListTitle))
            missing.Add("SharePoint:ListTitle");

        if (string.IsNullOrWhiteSpace(Database.ConnectionString) || Database.ConnectionString.Contains("YOUR_SERVER"))
            missing.Add("Database:ConnectionString");

        if (missing.Count > 0)
        {
            throw new InvalidOperationException(
                $"Missing or placeholder configuration values detected. Please update appsettings.json or set environment variables for:\n  - {string.Join("\n  - ", missing)}\n\nSee README.md for setup instructions.");
        }
    }
}
