using HotsyncApp.Configuration;
using HotsyncApp.Helpers;
using HotsyncApp.Models;

namespace HotsyncApp.Services;

/// <summary>
/// Orchestrates a single end-to-end sync run.
/// Coordinates between SharePointService and DatabaseService and produces a SyncResult.
///
/// Execution order:
///   1. Pre-flight: test SQL connection and ensure table exists.
///   2. Fetch: retrieve all customer records from SharePoint (paginated).
///   3. Load: read existing SharePoint IDs from SQL for upsert decisions.
///   4. Upsert: insert new records, update changed records.
///   5. Report: log final counts and duration.
/// </summary>
public sealed class SyncService
{
    private readonly SharePointService _sharePoint;
    private readonly DatabaseService _database;

    public SyncService(SharePointService sharePointService, DatabaseService databaseService)
    {
        _sharePoint = sharePointService;
        _database = databaseService;
    }

    /// <summary>
    /// Runs the full synchronisation pipeline.
    /// Returns a populated <see cref="SyncResult"/> regardless of whether errors occurred.
    /// Only throws on pre-flight failures (cannot continue at all).
    /// </summary>
    public async Task<SyncResult> RunAsync(CancellationToken ct = default)
    {
        var result = new SyncResult { StartedAt = DateTime.UtcNow };

        // ----------------------------------------------------------------
        // Step 1 — Pre-flight checks
        // ----------------------------------------------------------------
        Logger.Section("PRE-FLIGHT CHECKS");

        var sqlVersion = await _database.TestConnectionAsync(ct);
        Logger.Success($"SQL connected. Server: {sqlVersion}");

        await _database.EnsureTableExistsAsync(ct);

        // ----------------------------------------------------------------
        // Step 2 — Fetch from SharePoint
        // ----------------------------------------------------------------
        Logger.Section("SHAREPOINT FETCH");

        var customers = await _sharePoint.GetAllCustomersAsync(ct);
        result.TotalFetched = customers.Count;

        if (customers.Count == 0)
        {
            Logger.Warn("No records returned from SharePoint. Sync complete with nothing to do.");
            result.CompletedAt = DateTime.UtcNow;
            return result;
        }

        // ----------------------------------------------------------------
        // Step 3 — Load existing IDs from SQL
        // ----------------------------------------------------------------
        Logger.Section("DATABASE SYNC");

        var existingIds = await _database.GetExistingSharePointIdsAsync(ct);

        // ----------------------------------------------------------------
        // Step 4 — Upsert
        // ----------------------------------------------------------------
        Logger.Info($"Beginning upsert for {customers.Count} record(s)...");

        var (inserted, updated, errors) = await _database.UpsertAllAsync(customers, existingIds, ct);

        result.Inserted = inserted;
        result.Updated = updated;
        result.Errors = errors;
        result.Skipped = customers.Count - inserted - updated - errors;
        result.CompletedAt = DateTime.UtcNow;

        // ----------------------------------------------------------------
        // Step 5 — Final report
        // ----------------------------------------------------------------
        Logger.Section("SYNC COMPLETE");
        Logger.Info(result.ToString());

        if (result.HasErrors)
            Logger.Warn($"{errors} record(s) failed — review the error messages above. Other records were saved successfully.");
        else
            Logger.Success("All records synchronised without errors.");

        return result;
    }
}
