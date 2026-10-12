using HotsyncApp.Configuration;
using HotsyncApp.Helpers;
using HotsyncApp.Models;
using Microsoft.Data.SqlClient;

namespace HotsyncApp.Services;

/// <summary>
/// Manages all Azure SQL Database interactions for the Hotsync application.
///
/// Responsibilities:
///   1. Ensure the target table exists (idempotent DDL).
///   2. Load the set of SharePoint IDs already in the database (for efficient upsert decisions).
///   3. INSERT new records / UPDATE existing records — never creates duplicates.
///
/// Design decisions:
///   - No Entity Framework. All SQL is written as explicit parameterised statements.
///   - SharePointId is the natural sync key — it is stable, unique, and present on every item.
///   - Each upsert is wrapped in its own try/catch so one bad row doesn't abort the batch.
///   - The connection string is injected from config; this class never constructs credentials itself.
/// </summary>
public sealed class DatabaseService
{
    private readonly string _connectionString;

    // Table and column names are constants so they appear exactly once in the codebase.
    private const string TableName = "Customers";

    public DatabaseService(AppConfig config)
    {
        _connectionString = config.Database.ConnectionString;
    }

    // -------------------------------------------------------------------------
    // Connection helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Opens and returns a new SQL connection.
    /// Caller is responsible for disposing it (use 'await using').
    /// </summary>
    private async Task<SqlConnection> OpenConnectionAsync(CancellationToken ct)
    {
        var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync(ct);
        return connection;
    }

    // -------------------------------------------------------------------------
    // Schema initialisation
    // -------------------------------------------------------------------------

    /// <summary>
    /// Creates the Customers table if it does not already exist.
    /// Safe to call on every run — completely idempotent.
    /// </summary>
    public async Task EnsureTableExistsAsync(CancellationToken ct)
    {
        Logger.Info($"Verifying SQL table '{TableName}' exists...");

        const string ddl = $"""
            IF NOT EXISTS (
                SELECT 1
                FROM   sys.tables t
                JOIN   sys.schemas s ON t.schema_id = s.schema_id
                WHERE  s.name = 'dbo'
                AND    t.name = '{TableName}'
            )
            BEGIN
                CREATE TABLE dbo.{TableName} (
                    -- Primary key — auto-generated SQL identity, distinct from SharePoint ID.
                    Id              INT             NOT NULL IDENTITY(1,1) PRIMARY KEY,

                    -- The SharePoint list item ID — used as the sync key (must be unique).
                    SharePointId    INT             NOT NULL UNIQUE,

                    -- Customer fields sourced from the SharePoint list.
                    Name            NVARCHAR(255)   NULL,
                    Address         NVARCHAR(1000)  NULL,
                    NumberOfRigs    INT             NULL,
                    NumberOfJackUps INT             NULL,
                    NumberOfModus   INT             NULL,
                    SiteUrls        NVARCHAR(MAX)   NULL,

                    -- Audit columns managed by this application.
                    SpLastModified  DATETIME2       NULL,   -- SharePoint's Modified timestamp
                    SyncedAt        DATETIME2       NOT NULL DEFAULT GETUTCDATE(),
                    CreatedAt       DATETIME2       NOT NULL DEFAULT GETUTCDATE()
                );

                -- Index on SharePointId to make upsert lookups fast.
                CREATE UNIQUE NONCLUSTERED INDEX IX_{TableName}_SharePointId
                    ON dbo.{TableName} (SharePointId);

                PRINT 'Table {TableName} created.';
            END
            ELSE
            BEGIN
                PRINT 'Table {TableName} already exists — skipping creation.';
            END
            """;

        await using var conn = await OpenConnectionAsync(ct);
        await using var cmd = new SqlCommand(ddl, conn);
        cmd.CommandTimeout = 60;
        await cmd.ExecuteNonQueryAsync(ct);

        Logger.Success($"Table '{TableName}' is ready.");
    }

    // -------------------------------------------------------------------------
    // Upsert operations
    // -------------------------------------------------------------------------

    /// <summary>
    /// Retrieves the set of SharePoint IDs already present in the database.
    /// Used to decide INSERT vs UPDATE without issuing a separate SELECT per row.
    /// </summary>
    public async Task<HashSet<int>> GetExistingSharePointIdsAsync(CancellationToken ct)
    {
        Logger.Info("Loading existing SharePoint IDs from database...");

        var ids = new HashSet<int>();
        const string sql = $"SELECT SharePointId FROM dbo.{TableName}";

        await using var conn = await OpenConnectionAsync(ct);
        await using var cmd = new SqlCommand(sql, conn);
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        while (await reader.ReadAsync(ct))
            ids.Add(reader.GetInt32(0));

        Logger.Info($"Found {ids.Count} existing record(s) in SQL.");
        return ids;
    }

    /// <summary>
    /// Inserts a new customer row into the database.
    /// </summary>
    /// <returns>True if the insert succeeded.</returns>
    public async Task<bool> InsertCustomerAsync(SqlConnection conn, CustomerRecord record, CancellationToken ct)
    {
        const string sql = $"""
            INSERT INTO dbo.{TableName}
                (SharePointId, Name, Address, NumberOfRigs, NumberOfJackUps, NumberOfModus, SiteUrls, SpLastModified, SyncedAt, CreatedAt)
            VALUES
                (@SharePointId, @Name, @Address, @NumberOfRigs, @NumberOfJackUps, @NumberOfModus, @SiteUrls, @SpLastModified, GETUTCDATE(), GETUTCDATE())
            """;

        await using var cmd = BuildUpsertCommand(sql, conn, record);
        await cmd.ExecuteNonQueryAsync(ct);
        return true;
    }

    /// <summary>
    /// Updates an existing customer row identified by SharePointId.
    /// Updates all data fields and refreshes the SyncedAt timestamp.
    /// </summary>
    /// <returns>True if a row was actually updated (rowcount > 0).</returns>
    public async Task<bool> UpdateCustomerAsync(SqlConnection conn, CustomerRecord record, CancellationToken ct)
    {
        const string sql = $"""
            UPDATE dbo.{TableName}
            SET
                Name            = @Name,
                Address         = @Address,
                NumberOfRigs    = @NumberOfRigs,
                NumberOfJackUps = @NumberOfJackUps,
                NumberOfModus   = @NumberOfModus,
                SiteUrls        = @SiteUrls,
                SpLastModified  = @SpLastModified,
                SyncedAt        = GETUTCDATE()
            WHERE
                SharePointId    = @SharePointId
            """;

        await using var cmd = BuildUpsertCommand(sql, conn, record);
        var rowsAffected = await cmd.ExecuteNonQueryAsync(ct);
        return rowsAffected > 0;
    }

    /// <summary>
    /// Processes a full list of customer records — inserting new ones, updating changed ones.
    /// Opens a single SQL connection for the entire batch for efficiency.
    /// Each record is handled independently so one failure doesn't block the rest.
    /// </summary>
    public async Task<(int inserted, int updated, int errors)> UpsertAllAsync(
        IEnumerable<CustomerRecord> records,
        HashSet<int> existingIds,
        CancellationToken ct)
    {
        int inserted = 0, updated = 0, errors = 0;

        await using var conn = await OpenConnectionAsync(ct);

        foreach (var record in records)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                if (existingIds.Contains(record.SharePointId))
                {
                    var ok = await UpdateCustomerAsync(conn, record, ct);
                    if (ok) updated++;
                    else Logger.Warn($"  UPDATE matched 0 rows for SharePointId={record.SharePointId} — possible race condition.");
                }
                else
                {
                    await InsertCustomerAsync(conn, record, ct);
                    existingIds.Add(record.SharePointId); // prevent duplicate inserts in the same batch
                    inserted++;
                }
            }
            catch (SqlException ex)
            {
                errors++;
                Logger.Error($"  SQL error for SharePointId={record.SharePointId}: {ex.Message}");
            }
            catch (Exception ex)
            {
                errors++;
                Logger.Error($"  Unexpected error for SharePointId={record.SharePointId}: {ex.Message}");
            }
        }

        return (inserted, updated, errors);
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Builds a parameterised SqlCommand from a customer record.
    /// All parameters use DBNull for null C# values — never sends raw nulls.
    /// </summary>
    private static SqlCommand BuildUpsertCommand(string sql, SqlConnection conn, CustomerRecord r)
    {
        var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 30;

        cmd.Parameters.AddWithValue("@SharePointId",   r.SharePointId);
        cmd.Parameters.AddWithValue("@Name",            (object?)r.Name          ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@Address",         (object?)r.Address        ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfRigs",    (object?)r.NumberOfRigs   ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfJackUps", (object?)r.NumberOfJackUps ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@NumberOfModus",   (object?)r.NumberOfModus  ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SiteUrls",        (object?)r.SiteUrls       ?? DBNull.Value);
        cmd.Parameters.AddWithValue("@SpLastModified",  (object?)r.LastModified   ?? DBNull.Value);

        return cmd;
    }

    // -------------------------------------------------------------------------
    // Diagnostics
    // -------------------------------------------------------------------------

    /// <summary>
    /// Tests the SQL connection and returns a brief status string.
    /// Called before the sync so we fail-fast with a clear message.
    /// </summary>
    public async Task<string> TestConnectionAsync(CancellationToken ct)
    {
        Logger.Info("Testing SQL database connection...");
        try
        {
            await using var conn = await OpenConnectionAsync(ct);
            await using var cmd = new SqlCommand("SELECT @@VERSION", conn);
            var version = await cmd.ExecuteScalarAsync(ct) as string ?? "unknown";
            return version.Split('\n')[0].Trim(); // return just the first line of @@VERSION
        }
        catch (SqlException ex)
        {
            throw new InvalidOperationException(
                $"Cannot connect to SQL database.\n" +
                $"Verify ConnectionString in appsettings.json and that the firewall allows access from this IP.\n" +
                $"SqlException: {ex.Message}", ex);
        }
    }
}
