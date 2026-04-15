using System;
using System.Collections.Generic;

namespace ExcelSQLiteWeb.Models;

public record DatasetInfo(
    string Id,
    string DisplayName,
    long? EstimatedRows
);

public class ImportOptions
{
    public string Encoding { get; set; } = "UTF-8";
    public string Delimiter { get; set; } = ",";
    public bool HasHeader { get; set; } = true;
    public string TypeInference { get; set; } = "text";
    public int SampleRowsForInference { get; set; } = 200;
    public long? MaxRows { get; set; }
}

public class DatasetSchema
{
    public string DatasetId { get; set; } = "";
    public List<ColumnDef> Columns { get; set; } = new();

    public class ColumnDef
    {
        public string Name { get; set; } = "";
        public string SqliteType { get; set; } = "TEXT";
    }
}

public class RowData
{
    public object?[] Values { get; set; } = Array.Empty<object?>();
}
