using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelSQLiteWeb.Models;

namespace ExcelSQLiteWeb.Services;

public interface IDataImporter
{
    bool CanHandle(string filePath);

    string Kind { get; }

    IEnumerable<DatasetInfo> ListDatasets(string filePath);

    DatasetSchema GetSchema(string filePath, string datasetId);

    IAsyncEnumerable<RowData> ReadRowsAsync(string filePath, string datasetId, ImportOptions options);

    Task<ImportResult> ImportWorksheetAsync(
        string filePath,
        string worksheetName,
        string? tableName = null,
        string importMode = "text",
        IProgress<ImportProgress>? progress = null,
        CancellationToken cancellationToken = default,
        bool append = false);
}
