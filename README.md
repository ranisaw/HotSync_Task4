# HotSync_Task4 — Fennex Hotsync Application

A **C# .NET 10 Console Application** that synchronises customer data from a SharePoint Online list into an Azure SQL Database using Microsoft Entra ID authentication.

---

## Project Structure

```
HotsyncApp/
├── Configuration/
│   └── AppConfig.cs          # Strongly-typed config models
├── Helpers/
│   └── Logger.cs             # Timestamped console logger
├── Models/
│   ├── CustomerRecord.cs     # SharePoint list item DTO
│   └── SyncResult.cs         # Per-run statistics model
├── Services/
│   ├── SharePointService.cs  # Entra ID auth + SharePoint REST API + pagination
│   ├── DatabaseService.cs    # Azure SQL upsert logic (no Entity Framework)
│   └── SyncService.cs        # Orchestrates the full sync pipeline
├── Program.cs                # Entry point — config, wiring, exit codes
├── appsettings.json          # Configuration (fill in your values)
└── HotsyncApp.csproj
```

---

## Prerequisites

| Requirement | Detail |
|-------------|--------|
| .NET SDK | 10.0 or later — [download](https://dotnet.microsoft.com/download) |
| Azure App Registration | Created in Microsoft Entra ID (Azure AD) |
| SharePoint Permission | `Sites.Read.All` (application permission, admin-consented) |
| Azure SQL Database | Provisioned and accessible from your network |
| SQL Login / Permission | The app registration (or a SQL user) needs `db_datareader` + `db_datawriter` on the target database |

---

## Step 1 — Create an Azure App Registration

1. Go to **Azure Portal → Microsoft Entra ID → App registrations → New registration**
2. Name it something like `HotsyncApp`
3. Note the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets → New client secret** — note the value immediately (shown only once)
5. Go to **API permissions → Add a permission → SharePoint → Application permissions**
   - Add: `Sites.Read.All`
6. Click **Grant admin consent** (requires a Global Admin or SharePoint Admin)

---

## Step 2 — Configure the Application

Copy `appsettings.example.json` to `appsettings.json` and replace the placeholder values:

```bash
cp appsettings.example.json appsettings.json
```

```json
{
  "AzureAd": {
    "TenantId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientId":     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "ClientSecret": "your-client-secret-value"
  },
  "SharePoint": {
    "SiteUrl":    "https://yourcompany.sharepoint.com/sites/YourSite",
    "ListTitle":  "Customers"
  },
  "Database": {
    "ConnectionString": "Server=yourserver.database.windows.net;Database=YourDatabase;User Id=youruser;Password=yourpassword;Encrypt=True;TrustServerCertificate=False;"
  },
  "Sync": {
    "BatchSize": 100
  }
}
```

> **Never commit appsettings.json with real secrets to source control.**  
> Use environment variables or Azure Key Vault in production.

### Using Environment Variables (CI/CD / Production)

You can override any setting with an environment variable using double-underscore (`__`) as a section separator:

```powershell
$env:AzureAd__ClientSecret   = "your-secret"
$env:Database__ConnectionString = "Server=..."
```

---

## Step 3 — SharePoint List Column Names

The application maps the following **internal SharePoint field names** to SQL columns.  
If your list uses different internal column names, update `SharePointService.cs` → `MapToCustomerRecord()`:

| SharePoint Internal Name | SQL Column     | Type        |
|--------------------------|----------------|-------------|
| `Id`                     | `SharePointId` | INT (key)   |
| `Title`                  | `Name`         | NVARCHAR    |
| `Address`                | `Address`      | NVARCHAR    |
| `NumberOfRigs`           | `NumberOfRigs` | INT         |
| `NumberOfJackUps`        | `NumberOfJackUps` | INT      |
| `NumberOfMODUs`          | `NumberOfModus`   | INT      |
| `SiteURLs`               | `SiteUrls`     | NVARCHAR    |
| `Modified`               | `SpLastModified` | DATETIME2 |

> **To find the actual internal field name** for a column, go to SharePoint list Settings → the column → look at the URL parameter `Field=...`

---

## Step 4 — SQL Database Setup

The application **auto-creates the `Customers` table** on first run (idempotent — safe to run multiple times).

If you want to create it manually first, use this DDL:

```sql
CREATE TABLE dbo.Customers (
    Id              INT             NOT NULL IDENTITY(1,1) PRIMARY KEY,
    SharePointId    INT             NOT NULL UNIQUE,
    Name            NVARCHAR(255)   NULL,
    Address         NVARCHAR(1000)  NULL,
    NumberOfRigs    INT             NULL,
    NumberOfJackUps INT             NULL,
    NumberOfModus   INT             NULL,
    SiteUrls        NVARCHAR(MAX)   NULL,
    SpLastModified  DATETIME2       NULL,
    SyncedAt        DATETIME2       NOT NULL DEFAULT GETUTCDATE(),
    CreatedAt       DATETIME2       NOT NULL DEFAULT GETUTCDATE()
);

CREATE UNIQUE NONCLUSTERED INDEX IX_Customers_SharePointId
    ON dbo.Customers (SharePointId);
```

---

## Step 5 — Running the Application

```powershell
# Navigate to the project directory
cd "D:\React projects\hotsync App\HotsyncApp"

# Run (builds and executes)
dotnet run --configuration Release

# Or build first, then run the executable directly
dotnet build --configuration Release
.\bin\Release\net10.0\HotsyncApp.exe
```

### Exit Codes

| Code | Meaning |
|------|---------|
| `0`  | All records synced successfully |
| `1`  | Partial success — some records had errors |
| `2`  | Fatal error — config invalid or connectivity failure |
| `3`  | Cancelled by user (Ctrl+C) |

---

## How the Sync Works

```
Program.cs
  └── Validates config (fails fast if placeholders remain)
      └── SyncService.RunAsync()
            ├── DatabaseService.TestConnectionAsync()     ← pre-flight SQL check
            ├── DatabaseService.EnsureTableExistsAsync()  ← idempotent DDL
            ├── SharePointService.GetAllCustomersAsync()  ← Entra ID auth + paginated REST
            ├── DatabaseService.GetExistingSharePointIdsAsync()
            └── DatabaseService.UpsertAllAsync()
                  ├── If SharePointId not in DB → INSERT
                  └── If SharePointId exists   → UPDATE
```

**Upsert key**: `SharePointId` (the SharePoint list item integer ID).  
This ensures the same item is never inserted twice, regardless of how many times the sync runs.

---

## Troubleshooting

| Error | Likely Cause | Fix |
|-------|-------------|-----|
| `AADSTS70011` | Wrong scope or permission not granted | Check `Sites.Read.All` is admin-consented |
| `HTTP 403 Forbidden` from SharePoint | Token acquired but no access to the site | Ensure the App Registration has access to the specific site collection |
| `SqlException: Login failed` | Wrong SQL credentials | Check ConnectionString, user permissions |
| `Firewall rule` error | Azure SQL firewall blocks your IP | Add your IP in Azure Portal → SQL Server → Networking |
| Column not found in SP response | Internal field name mismatch | Check field names in SharePoint List Settings |

---

## Security Notes

- Client secrets are read from `appsettings.json` or environment variables — never hardcoded.
- For production, store secrets in **Azure Key Vault** and use **Managed Identity**.
- Add `appsettings.json` to `.gitignore` if it contains real secrets.
- The SQL connection uses `Encrypt=True` by default — data in transit is always encrypted.
