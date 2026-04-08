using NAPS2.ImportExport;

namespace NAPS2.Scan.Batch;

public class BatchSettings
{
    public string? ProfileDisplayName { get; set; }

    public BatchScanType ScanType { get; set; }

    public int ScanCount { get; set; }

    public double ScanIntervalSeconds { get; set; }

    public BatchOutputType OutputType { get; set;  }

    public SaveSeparator SaveSeparator { get; set; }

    // Used when splitting by file (one file per page, or our "separate by N pages" option).
    // For FilePerPage mode, this controls the number of pages per output file.
    public int SaveSeparatorPageCount { get; set; } = 1;

    public string? SavePath { get; set; }
}