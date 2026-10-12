using HotsyncApp.Configuration;
using HotsyncApp.Helpers;
using HotsyncApp.Services;
using Microsoft.Extensions.Configuration;

// ============================================================
// Fennex Hotsync Application — Task 4
// Copies customer records from SharePoint → Azure SQL Database
// ============================================================

Logger.Section("FENNEX HOTSYNC — Task 4");
Logger.Info($"Application started at {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC");
Logger.Info($"Runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");

// ----------------------------------------------------------------
// 1. Build configuration from appsettings.json + environment vars
//    Environment variables override appsettings.json values.
//    Format for env var override: AzureAd__ClientSecret, Database__ConnectionString, etc.
// ----------------------------------------------------------------
IConfiguration configuration;
try
{
    configuration = new ConfigurationBuilder()
        .SetBasePath(AppContext.BaseDirectory)
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
        .AddEnvironmentVariables()   // allows secrets via env vars in CI/CD pipelines
        .Build();
}
catch (Exception ex)
{
    Logger.Error("Failed to load appsettings.json.", ex);
    Logger.Error("Ensure appsettings.json is present in the output directory and is valid JSON.");
    return 2;
}

// ----------------------------------------------------------------
// 2. Bind configuration to strongly-typed AppConfig model
// ----------------------------------------------------------------
var appConfig = new AppConfig();
configuration.GetSection("AzureAd").Bind(appConfig.AzureAd);
configuration.GetSection("SharePoint").Bind(appConfig.SharePoint);
configuration.GetSection("Database").Bind(appConfig.Database);
configuration.GetSection("Sync").Bind(appConfig.Sync);

// ----------------------------------------------------------------
// 3. Validate — fail fast with a clear message if any placeholder remains
// ----------------------------------------------------------------
try
{
    appConfig.Validate();
    Logger.Success("Configuration validated successfully.");
}
catch (InvalidOperationException ex)
{
    Logger.Error(ex.Message);
    return 1;
}

// ----------------------------------------------------------------
// 4. Wire up services
// ----------------------------------------------------------------
using var cts = new CancellationTokenSource();

// Allow Ctrl+C to cancel the sync gracefully.
Console.CancelKeyPress += (_, e) =>
{
    e.Cancel = true;
    Logger.Warn("Cancellation requested — finishing current operation...");
    cts.Cancel();
};

var httpClient = new HttpClient
{
    Timeout = TimeSpan.FromSeconds(120)
};

var sharePointService = new SharePointService(appConfig, httpClient);
var databaseService = new DatabaseService(appConfig);
var syncService = new SyncService(sharePointService, databaseService);

// ----------------------------------------------------------------
// 5. Run the sync
// ----------------------------------------------------------------
try
{
    var result = await syncService.RunAsync(cts.Token);

    // Exit code convention:
    //   0 = success (all records synced, no errors)
    //   1 = partial success (some records had errors, but others were saved)
    //   2 = configuration or connectivity failure
    var exitCode = result.HasErrors ? 1 : 0;
    Logger.Info($"Exiting with code {exitCode}.");
    return exitCode;
}
catch (OperationCanceledException)
{
    Logger.Warn("Sync was cancelled by the user.");
    return 3;
}
catch (InvalidOperationException ex)
{
    // Pre-flight or fatal connectivity error — already logged inside the service.
    Logger.Error("Fatal error — sync could not complete.", ex);
    Logger.Error(ex.Message);
    return 2;
}
catch (Exception ex)
{
    Logger.Error("Unexpected unhandled exception.", ex);
    Logger.Error($"  {ex}");
    return 2;
}
finally
{
    httpClient.Dispose();
    Logger.Info("Application shutdown complete.");
}
