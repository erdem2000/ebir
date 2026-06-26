using System.Text;
using ClosedXML.Excel;

namespace ExcelBirlestirici;

public sealed class MergeOptions
{
    public bool SkipInvalidRows { get; init; }

    /// <summary>Her Excel dosyasında birleştirilecek ilk sayfa sayısı.</summary>
    public int SheetsToMerge { get; init; } = 1;

    /// <summary>
    /// Açıkken sayfalar indeks bazında ayrı birleştirilir; CSV dosyaları atlanır.
    /// </summary>
    public bool WorkbookProtectedMode { get; init; } = true;

    /// <summary>
    /// Şifre korumalı Excel dosyası için şifre ister.
    /// Parametreler: dosya adı, önceki şifre yanlış mı. null dönerse dosya atlanır.
    /// </summary>
    public Func<string, bool, string?>? RequestPassword { get; init; }
}

public sealed class SheetHeaderValidationResult
{
    public bool IsValid { get; init; }
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();
}

public enum RowIssueAction
{
    Fixed,
    Skipped
}

public sealed class RowIssue
{
    public string FileName { get; init; } = string.Empty;
    public int RowNumber { get; init; }
    public string? ColumnName { get; init; }
    public string Reason { get; init; } = string.Empty;
    public RowIssueAction Action { get; init; }

    public string Summary =>
        ColumnName is null
            ? $"{FileName} satır {RowNumber}: {Reason}"
            : $"{FileName} satır {RowNumber}, kolon {ColumnName}: {Reason}";
}

public sealed class MergeRowException : Exception
{
    public RowIssue Issue { get; }

    public MergeRowException(RowIssue issue)
        : base(issue.Summary)
    {
        Issue = issue;
    }
}

public sealed class MergeResult
{
    public string? OutputPath { get; init; }
    public int FilesProcessed { get; init; }
    public int RowsWritten { get; init; }
    public int DuplicatesSkipped { get; init; }
    public IReadOnlyList<string> DiscoveredFiles { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> MergedFiles { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> FileErrors { get; init; } = Array.Empty<string>();
    public IReadOnlyList<RowIssue> FixedRows { get; init; } = Array.Empty<RowIssue>();
    public IReadOnlyList<RowIssue> SkippedRows { get; init; } = Array.Empty<RowIssue>();
    public bool AbortedDueToInvalidRow { get; init; }
    public bool AbortedDueToHeaderMismatch { get; init; }
}

internal sealed class SheetMergeSlot
{
    public List<string> ColumnOrder { get; } = [];
    public List<IReadOnlyDictionary<string, XLCellValue>> Rows { get; } = [];
    public HashSet<string> SeenPhoneNumbers { get; } = new(StringComparer.OrdinalIgnoreCase);
}

public static class ExcelMergeService
{
    private static readonly string[] InputExtensions = [".xlsx", ".xlsm", ".csv"];
    private const string OutputPrefix = "Birleştirilmiş_";
    private const string PhoneNrColumn = "PhoneNr";
    private const int MaxExcelRows = 1_048_576;
    private const int MaxDataRows = MaxExcelRows - 1;

    public static SheetHeaderValidationResult ValidateSheetHeaders(string folderPath, MergeOptions? options = null)
    {
        options ??= new MergeOptions();

        if (string.IsNullOrWhiteSpace(folderPath))
            throw new ArgumentException("Klasör yolu boş olamaz.", nameof(folderPath));

        var dir = new DirectoryInfo(folderPath);
        if (!dir.Exists)
            throw new DirectoryNotFoundException($"Klasör bulunamadı: {folderPath}");

        var files = EnumerateInputFiles(dir.FullName)
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (files.Count == 0)
        {
            return new SheetHeaderValidationResult
            {
                IsValid = false,
                Errors = ["Uygun dosya bulunamadı (.xlsx / .xlsm / .csv)."]
            };
        }

        var passwordSession = new ExcelPasswordSession(options.RequestPassword);
        return ValidateSheetHeaders(files, options, passwordSession);
    }

    public static MergeResult Merge(string folderPath, MergeOptions? options = null)
    {
        options ??= new MergeOptions();

        if (string.IsNullOrWhiteSpace(folderPath))
            throw new ArgumentException("Klasör yolu boş olamaz.", nameof(folderPath));

        var dir = new DirectoryInfo(folderPath);
        if (!dir.Exists)
            throw new DirectoryNotFoundException($"Klasör bulunamadı: {folderPath}");

        var files = EnumerateInputFiles(dir.FullName)
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var discoveredFileNames = files
            .Select(path => Path.GetFileName(path) ?? path)
            .ToList();

        if (files.Count == 0)
            return new MergeResult
            {
                DiscoveredFiles = discoveredFileNames,
                FileErrors = ["Uygun dosya bulunamadı (.xlsx / .xlsm / .csv)."]
            };

        var passwordSession = new ExcelPasswordSession(options.RequestPassword);
        var headerValidation = ValidateSheetHeaders(files, options, passwordSession);
        if (!headerValidation.IsValid)
        {
            return new MergeResult
            {
                DiscoveredFiles = discoveredFileNames,
                FileErrors = headerValidation.Errors.ToList(),
                AbortedDueToHeaderMismatch = true
            };
        }

        if (options.WorkbookProtectedMode)
            return MergeInWorkbookProtectedMode(dir, files, discoveredFileNames, options, passwordSession);

        var columnOrder = new List<string>();
        var fileErrors = new List<string>();
        var mergedFiles = new List<string>();
        var fixedRows = new List<RowIssue>();
        var skippedRows = new List<RowIssue>();
        var seenPhoneNumbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicatesSkipped = 0;
        var allRows = new List<IReadOnlyDictionary<string, XLCellValue>>();

        var reachedRowLimit = false;
        var abortedDueToInvalidRow = false;

        try
        {
            foreach (var path in files)
            {
                var fileName = Path.GetFileName(path) ?? path;
                var ext = Path.GetExtension(path);

                if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    ProcessCsvFile(
                        path,
                        fileName,
                        options,
                        columnOrder,
                        allRows,
                        mergedFiles,
                        fileErrors,
                        fixedRows,
                        skippedRows,
                        seenPhoneNumbers,
                        ref duplicatesSkipped);

                    if (allRows.Count >= MaxDataRows)
                        reachedRowLimit = true;
                }
                else
                {
                    XLWorkbook wb;
                    try
                    {
                        wb = passwordSession.Open(path, fileName);
                    }
                    catch (ExcelPasswordCancelledException ex)
                    {
                        fileErrors.Add(ex.Message);
                        continue;
                    }

                    using (wb)
                    {
                        var worksheets = wb.Worksheets.Take(Math.Max(1, options.SheetsToMerge)).ToList();
                        if (worksheets.Count == 0)
                        {
                            fileErrors.Add($"{fileName}: çalışma sayfası yok.");
                            continue;
                        }

                        foreach (var ws in worksheets)
                        {
                            var sourceName = $"{fileName} / {ws.Name}";
                            var processed = ProcessExcelWorksheet(
                                ws,
                                sourceName,
                                options,
                                columnOrder,
                                allRows,
                                fixedRows,
                                skippedRows,
                                seenPhoneNumbers,
                                ref duplicatesSkipped,
                                ref reachedRowLimit,
                                fileErrors);

                            if (processed)
                                mergedFiles.Add(sourceName);

                            if (reachedRowLimit)
                                break;
                        }
                    }
                }

                if (reachedRowLimit)
                    break;
            }
        }
        catch (MergeRowException ex)
        {
            abortedDueToInvalidRow = true;
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }
        catch (Exception ex) when (!options.SkipInvalidRows)
        {
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }

        if (columnOrder.Count == 0)
            return new MergeResult
            {
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors.Count > 0 ? fileErrors : ["Hiçbir dosyadan kolon başlığı okunamadı."]
            };

        var outputName = $"{OutputPrefix}{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        var outputPath = Path.Combine(dir.FullName, outputName);

        try
        {
            using var outWb = new XLWorkbook();
            var outWs = outWb.Worksheets.Add("Birleştirilmiş");
            var colCount = columnOrder.Count;

            for (var c = 0; c < colCount; c++)
                outWs.Cell(1, c + 1).Value = columnOrder[c];

            var outRow = 2;
            foreach (var src in allRows)
            {
                if (outRow > MaxExcelRows)
                {
                    fileErrors.Add($"Excel maksimum satır sayısı aşıldı ({MaxExcelRows:N0}).");
                    break;
                }

                var writeResult = TryWriteOutputRow(outWs, outRow, columnOrder, src, options, outRow - 1);
                switch (writeResult.Outcome)
                {
                    case RowReadOutcome.Success:
                        outRow++;
                        break;

                    case RowReadOutcome.Fixed:
                        if (writeResult.Issue is not null)
                            fixedRows.Add(writeResult.Issue);
                        outRow++;
                        break;

                    case RowReadOutcome.Skipped:
                        if (writeResult.Issue is not null)
                            skippedRows.Add(writeResult.Issue);
                        break;

                    case RowReadOutcome.Failed:
                        throw new MergeRowException(writeResult.Issue!);
                }
            }

            outWs.SheetView.FreezeRows(1);
            outWs.Row(1).Style.Font.Bold = true;
            SafeAdjustColumnWidths(outWs, colCount, outRow - 1);

            outWb.SaveAs(outputPath);
        }
        catch (MergeRowException ex)
        {
            abortedDueToInvalidRow = true;
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                RowsWritten = allRows.Count,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }

        return new MergeResult
        {
            OutputPath = outputPath,
            FilesProcessed = files.Count,
            RowsWritten = allRows.Count,
            DuplicatesSkipped = duplicatesSkipped,
            DiscoveredFiles = discoveredFileNames,
            MergedFiles = mergedFiles,
            FixedRows = fixedRows,
            SkippedRows = skippedRows,
            FileErrors = fileErrors,
            AbortedDueToInvalidRow = abortedDueToInvalidRow
        };
    }

    private static MergeResult MergeInWorkbookProtectedMode(
        DirectoryInfo dir,
        IReadOnlyList<string> files,
        IReadOnlyList<string> discoveredFileNames,
        MergeOptions options,
        ExcelPasswordSession passwordSession)
    {
        var sheetsToMerge = Math.Max(1, options.SheetsToMerge);
        var slots = Enumerable.Range(0, sheetsToMerge).Select(_ => new SheetMergeSlot()).ToList();
        var fileErrors = new List<string>();
        var mergedFiles = new List<string>();
        var fixedRows = new List<RowIssue>();
        var skippedRows = new List<RowIssue>();
        var duplicatesSkipped = 0;
        var reachedRowLimit = false;
        var abortedDueToInvalidRow = false;
        var totalRowsWritten = 0;

        try
        {
            foreach (var path in files)
            {
                var fileName = Path.GetFileName(path) ?? path;
                var ext = Path.GetExtension(path);

                if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    fileErrors.Add($"{fileName}: korumalı modda CSV atlandı.");
                    continue;
                }

                XLWorkbook wb;
                try
                {
                    wb = passwordSession.Open(path, fileName);
                }
                catch (ExcelPasswordCancelledException ex)
                {
                    fileErrors.Add(ex.Message);
                    continue;
                }

                using (wb)
                {
                    var worksheets = wb.Worksheets.Take(sheetsToMerge).ToList();
                    if (worksheets.Count == 0)
                    {
                        fileErrors.Add($"{fileName}: çalışma sayfası yok.");
                        continue;
                    }

                    for (var sheetIndex = 0; sheetIndex < worksheets.Count; sheetIndex++)
                    {
                        var ws = worksheets[sheetIndex];
                        var sourceName = $"{fileName} / {ws.Name}";
                        var slot = slots[sheetIndex];
                        var processed = ProcessExcelWorksheet(
                            ws,
                            sourceName,
                            options,
                            slot.ColumnOrder,
                            slot.Rows,
                            fixedRows,
                            skippedRows,
                            slot.SeenPhoneNumbers,
                            ref duplicatesSkipped,
                            ref reachedRowLimit,
                            fileErrors);

                        if (processed)
                            mergedFiles.Add(sourceName);

                        if (reachedRowLimit)
                            break;
                    }
                }

                if (reachedRowLimit)
                    break;
            }
        }
        catch (MergeRowException ex)
        {
            abortedDueToInvalidRow = true;
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }
        catch (Exception ex) when (!options.SkipInvalidRows)
        {
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }

        var hasData = slots.Any(s => s.ColumnOrder.Count > 0);
        if (!hasData)
        {
            return new MergeResult
            {
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors.Count > 0
                    ? fileErrors
                    : ["Hiçbir dosyadan kolon başlığı okunamadı."]
            };
        }

        var outputName = $"{OutputPrefix}{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        var outputPath = Path.Combine(dir.FullName, outputName);

        try
        {
            using var outWb = new XLWorkbook();

            for (var sheetIndex = 0; sheetIndex < slots.Count; sheetIndex++)
            {
                var slot = slots[sheetIndex];
                if (slot.ColumnOrder.Count == 0)
                    continue;

                var outWs = outWb.Worksheets.Add($"Birleştirilmiş {sheetIndex + 1}");
                var colCount = slot.ColumnOrder.Count;

                for (var c = 0; c < colCount; c++)
                    outWs.Cell(1, c + 1).Value = slot.ColumnOrder[c];

                var outRow = 2;
                foreach (var src in slot.Rows)
                {
                    if (outRow > MaxExcelRows)
                    {
                        fileErrors.Add($"Excel maksimum satır sayısı aşıldı ({MaxExcelRows:N0}).");
                        break;
                    }

                    var writeResult = TryWriteOutputRow(outWs, outRow, slot.ColumnOrder, src, options, outRow - 1);
                    switch (writeResult.Outcome)
                    {
                        case RowReadOutcome.Success:
                            outRow++;
                            totalRowsWritten++;
                            break;

                        case RowReadOutcome.Fixed:
                            if (writeResult.Issue is not null)
                                fixedRows.Add(writeResult.Issue);
                            outRow++;
                            totalRowsWritten++;
                            break;

                        case RowReadOutcome.Skipped:
                            if (writeResult.Issue is not null)
                                skippedRows.Add(writeResult.Issue);
                            break;

                        case RowReadOutcome.Failed:
                            throw new MergeRowException(writeResult.Issue!);
                    }
                }

                outWs.SheetView.FreezeRows(1);
                outWs.Row(1).Style.Font.Bold = true;
                SafeAdjustColumnWidths(outWs, colCount, outRow - 1);
            }

            if (outWb.Worksheets.Count == 0)
            {
                return new MergeResult
                {
                    DiscoveredFiles = discoveredFileNames,
                    MergedFiles = mergedFiles,
                    FixedRows = fixedRows,
                    SkippedRows = skippedRows,
                    FileErrors = ["Hiçbir sayfaya veri yazılamadı."]
                };
            }

            outWb.SaveAs(outputPath);
        }
        catch (MergeRowException ex)
        {
            abortedDueToInvalidRow = true;
            fileErrors.Add(ex.Message);
            return new MergeResult
            {
                FilesProcessed = files.Count,
                RowsWritten = totalRowsWritten,
                DuplicatesSkipped = duplicatesSkipped,
                DiscoveredFiles = discoveredFileNames,
                MergedFiles = mergedFiles,
                FixedRows = fixedRows,
                SkippedRows = skippedRows,
                FileErrors = fileErrors,
                AbortedDueToInvalidRow = true
            };
        }

        return new MergeResult
        {
            OutputPath = outputPath,
            FilesProcessed = files.Count,
            RowsWritten = totalRowsWritten,
            DuplicatesSkipped = duplicatesSkipped,
            DiscoveredFiles = discoveredFileNames,
            MergedFiles = mergedFiles,
            FixedRows = fixedRows,
            SkippedRows = skippedRows,
            FileErrors = fileErrors,
            AbortedDueToInvalidRow = abortedDueToInvalidRow
        };
    }

    private enum RowReadOutcome
    {
        Success,
        Fixed,
        Skipped,
        Failed
    }

    private sealed class RowProcessResult
    {
        public RowReadOutcome Outcome { get; init; }
        public Dictionary<string, XLCellValue>? Row { get; init; }
        public RowIssue? Issue { get; init; }
    }

    private static RowProcessResult TryReadExcelRow(
        string fileName,
        int rowNumber,
        Dictionary<int, string> headers,
        IXLRow row,
        MergeOptions options)
    {
        var dict = new Dictionary<string, XLCellValue>(StringComparer.OrdinalIgnoreCase);
        var hadRepair = false;
        string? firstError = null;
        string? firstColumn = null;

        foreach (var col in headers.Keys.OrderBy(c => c))
        {
            var columnName = headers[col];
            var cell = row.Cell(col);

            if (TryReadCellValue(cell, out var value, out var readError))
            {
                var validation = ValidateCellValue(value);
                if (validation.IsValid)
                {
                    dict[columnName] = value;
                    continue;
                }

                firstError ??= validation.Error ?? readError ?? "Hücre değeri geçersiz";
                firstColumn ??= columnName;

                if (options.SkipInvalidRows)
                {
                    return new RowProcessResult
                    {
                        Outcome = RowReadOutcome.Skipped,
                        Issue = new RowIssue
                        {
                            FileName = fileName,
                            RowNumber = rowNumber,
                            ColumnName = columnName,
                            Reason = firstError,
                            Action = RowIssueAction.Skipped
                        }
                    };
                }

                dict[columnName] = SanitizeForOutput(value);
                hadRepair = true;
                firstError = validation.Error ?? firstError;
                continue;
            }

            firstError ??= readError ?? "Hücre değeri okunamadı";
            firstColumn ??= columnName;

            if (options.SkipInvalidRows)
            {
                return new RowProcessResult
                {
                    Outcome = RowReadOutcome.Skipped,
                    Issue = new RowIssue
                    {
                        FileName = fileName,
                        RowNumber = rowNumber,
                        ColumnName = columnName,
                        Reason = firstError,
                        Action = RowIssueAction.Skipped
                    }
                };
            }

            if (TryRepairCellValue(cell, out var repaired, out var repairNote))
            {
                dict[columnName] = repaired;
                hadRepair = true;
                firstError = repairNote ?? firstError;
                continue;
            }

            return new RowProcessResult
            {
                Outcome = RowReadOutcome.Failed,
                Issue = new RowIssue
                {
                    FileName = fileName,
                    RowNumber = rowNumber,
                    ColumnName = columnName,
                    Reason = firstError,
                    Action = RowIssueAction.Skipped
                }
            };
        }

        if (hadRepair)
        {
            return new RowProcessResult
            {
                Outcome = RowReadOutcome.Fixed,
                Row = dict,
                Issue = new RowIssue
                {
                    FileName = fileName,
                    RowNumber = rowNumber,
                    ColumnName = firstColumn,
                    Reason = firstError ?? "Hücre değeri düzeltildi",
                    Action = RowIssueAction.Fixed
                }
            };
        }

        return new RowProcessResult
        {
            Outcome = RowReadOutcome.Success,
            Row = dict
        };
    }

    private static RowProcessResult TryWriteOutputRow(
        IXLWorksheet worksheet,
        int outRow,
        IReadOnlyList<string> columnOrder,
        IReadOnlyDictionary<string, XLCellValue> src,
        MergeOptions options,
        int logicalRowNumber)
    {
        var hadRepair = false;
        string? firstError = null;
        string? firstColumn = null;

        for (var c = 0; c < columnOrder.Count; c++)
        {
            var key = columnOrder[c];
            if (!src.TryGetValue(key, out var value))
                continue;

            var cell = worksheet.Cell(outRow, c + 1);
            var safeValue = SanitizeForOutput(value);
            if (!safeValue.Equals(value))
            {
                hadRepair = true;
                firstError ??= "Geçersiz tarih değeri metne dönüştürüldü";
                firstColumn ??= key;
            }

            if (TryWriteCellValue(cell, safeValue, out var writeError))
                continue;

            firstError ??= writeError ?? "Hücre değeri yazılamadı";
            firstColumn ??= key;

            if (options.SkipInvalidRows)
            {
                return new RowProcessResult
                {
                    Outcome = RowReadOutcome.Skipped,
                    Issue = new RowIssue
                    {
                        FileName = "Çıktı",
                        RowNumber = logicalRowNumber,
                        ColumnName = key,
                        Reason = firstError,
                        Action = RowIssueAction.Skipped
                    }
                };
            }

            if (TryRepairAndWriteCellValue(cell, value, out var repairNote))
            {
                hadRepair = true;
                firstError = repairNote ?? firstError;
                continue;
            }

            return new RowProcessResult
            {
                Outcome = RowReadOutcome.Failed,
                Issue = new RowIssue
                {
                    FileName = "Çıktı",
                    RowNumber = logicalRowNumber,
                    ColumnName = key,
                    Reason = firstError,
                    Action = RowIssueAction.Skipped
                }
            };
        }

        if (hadRepair)
        {
            return new RowProcessResult
            {
                Outcome = RowReadOutcome.Fixed,
                Issue = new RowIssue
                {
                    FileName = "Çıktı",
                    RowNumber = logicalRowNumber,
                    ColumnName = firstColumn,
                    Reason = firstError ?? "Hücre değeri düzeltildi",
                    Action = RowIssueAction.Fixed
                }
            };
        }

        return new RowProcessResult { Outcome = RowReadOutcome.Success };
    }

    private sealed class CellValidationResult
    {
        public bool IsValid { get; init; }
        public string? Error { get; init; }
    }

    private static CellValidationResult ValidateCellValue(XLCellValue value)
    {
        try
        {
            if (value.Type == XLDataType.DateTime)
                _ = value.GetDateTime();

            return new CellValidationResult { IsValid = true };
        }
        catch (Exception ex)
        {
            return new CellValidationResult { IsValid = false, Error = ex.Message };
        }
    }

    private static XLCellValue SanitizeForOutput(XLCellValue value)
    {
        var validation = ValidateCellValue(value);
        if (validation.IsValid)
            return value;

        try
        {
            return value.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool TryReadCellValue(IXLCell cell, out XLCellValue value, out string? error)
    {
        try
        {
            value = cell.Value;
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            value = default;
            error = ex.Message;
            return false;
        }
    }

    private static bool TryRepairCellValue(IXLCell cell, out XLCellValue value, out string? repairNote)
    {
        repairNote = null;
        try
        {
            var formatted = cell.GetFormattedString();
            if (!string.IsNullOrEmpty(formatted))
            {
                value = formatted;
                repairNote = "Biçimlendirilmiş metin olarak düzeltildi";
                return true;
            }
        }
        catch
        {
            // continue to next strategy
        }

        try
        {
            var text = cell.GetString();
            value = text;
            repairNote = "Metin olarak düzeltildi";
            return true;
        }
        catch
        {
            // continue to next strategy
        }

        try
        {
            if (!cell.IsEmpty())
            {
                value = cell.GetDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                repairNote = "Sayısal değer metin olarak düzeltildi";
                return true;
            }
        }
        catch
        {
            // continue to next strategy
        }

        value = default;
        return false;
    }

    private static void SafeAdjustColumnWidths(IXLWorksheet worksheet, int colCount, int lastDataRow)
    {
        try
        {
            var sampleEnd = Math.Min(Math.Max(1, lastDataRow), 500);
            worksheet.Columns(1, colCount).AdjustToContents(1, sampleEnd);
        }
        catch
        {
            // Invalid date/formatted values can break auto-fit; output is still valid.
        }
    }

    private static bool TryWriteCellValue(IXLCell cell, XLCellValue value, out string? error)
    {
        try
        {
            cell.Value = value;
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool TryRepairAndWriteCellValue(IXLCell cell, XLCellValue value, out string? repairNote)
    {
        repairNote = null;
        try
        {
            cell.Value = value.ToString();
            repairNote = "Metin olarak yazıldı";
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static IEnumerable<string> EnumerateInputFiles(string folder)
    {
        foreach (var ext in InputExtensions)
        {
            foreach (var path in Directory.EnumerateFiles(folder, $"*{ext}", SearchOption.TopDirectoryOnly))
            {
                var name = Path.GetFileName(path);
                if (name.StartsWith(OutputPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;

                yield return path;
            }
        }
    }

    private static SheetHeaderValidationResult ValidateSheetHeaders(
        IReadOnlyList<string> files,
        MergeOptions options,
        ExcelPasswordSession passwordSession)
    {
        var errors = new List<string>();
        var sources = CollectSheetHeaderSources(files, options, passwordSession, errors);

        if (options.WorkbookProtectedMode && sources.Count == 0)
        {
            if (errors.Count == 0)
                errors.Add("Korumalı modda işlenecek Excel dosyası bulunamadı.");
            else if (!errors.Any(e => e.Contains("Korumalı modda işlenecek", StringComparison.Ordinal)))
                errors.Insert(0, "Korumalı modda işlenecek Excel dosyası bulunamadı.");

            return new SheetHeaderValidationResult { IsValid = false, Errors = errors };
        }

        if (sources.Count == 0)
        {
            if (errors.Count == 0)
                errors.Add("Hiçbir sayfadan sütun başlığı okunamadı.");

            return new SheetHeaderValidationResult { IsValid = false, Errors = errors };
        }

        if (options.WorkbookProtectedMode)
            ValidateHeadersBySheetIndex(sources, errors);
        else
            ValidateHeadersGlobally(sources, errors);

        return new SheetHeaderValidationResult
        {
            IsValid = errors.Count == 0,
            Errors = errors
        };
    }

    private static void ValidateHeadersGlobally(
        IReadOnlyList<(string Source, IReadOnlyList<string> Headers, int SheetIndex)> sources,
        List<string> errors)
    {
        var referenceSource = sources[0].Source;
        var referenceHeaders = sources[0].Headers;

        for (var i = 1; i < sources.Count; i++)
        {
            var (source, headers, _) = sources[i];
            if (HeadersEqual(referenceHeaders, headers))
                continue;

            errors.Add(
                $"{source}: Sütunlar '{referenceSource}' ile uyuşmuyor.{Environment.NewLine}" +
                $"  Beklenen: {FormatHeaders(referenceHeaders)}{Environment.NewLine}" +
                $"  Bulunan: {FormatHeaders(headers)}");
        }
    }

    private static void ValidateHeadersBySheetIndex(
        IReadOnlyList<(string Source, IReadOnlyList<string> Headers, int SheetIndex)> sources,
        List<string> errors)
    {
        foreach (var group in sources.GroupBy(s => s.SheetIndex).OrderBy(g => g.Key))
        {
            var groupSources = group.ToList();
            if (groupSources.Count <= 1)
                continue;

            var referenceSource = groupSources[0].Source;
            var referenceHeaders = groupSources[0].Headers;
            var sheetNumber = group.Key + 1;

            for (var i = 1; i < groupSources.Count; i++)
            {
                var (source, headers, _) = groupSources[i];
                if (HeadersEqual(referenceHeaders, headers))
                    continue;

                errors.Add(
                    $"Sayfa {sheetNumber} — {source}: Sütunlar '{referenceSource}' ile uyuşmuyor.{Environment.NewLine}" +
                    $"  Beklenen: {FormatHeaders(referenceHeaders)}{Environment.NewLine}" +
                    $"  Bulunan: {FormatHeaders(headers)}");
            }
        }
    }

    private static List<(string Source, IReadOnlyList<string> Headers, int SheetIndex)> CollectSheetHeaderSources(
        IReadOnlyList<string> files,
        MergeOptions options,
        ExcelPasswordSession passwordSession,
        List<string> errors)
    {
        var sources = new List<(string Source, IReadOnlyList<string> Headers, int SheetIndex)>();
        var sheetsToMerge = Math.Max(1, options.SheetsToMerge);

        foreach (var path in files)
        {
            var fileName = Path.GetFileName(path) ?? path;
            var ext = Path.GetExtension(path);

            if (ext.Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                if (options.WorkbookProtectedMode)
                    continue;

                if (TryGetCsvHeaderNames(path, out var csvHeaders, out var csvError))
                    sources.Add((fileName, csvHeaders!, SheetIndex: 0));
                else
                    errors.Add($"{fileName}: {csvError}");
                continue;
            }

            XLWorkbook wb;
            try
            {
                wb = passwordSession.Open(path, fileName);
            }
            catch (ExcelPasswordCancelledException ex)
            {
                errors.Add(ex.Message);
                continue;
            }

            using (wb)
            {
                var worksheets = wb.Worksheets.Take(sheetsToMerge).ToList();
                if (worksheets.Count == 0)
                {
                    errors.Add($"{fileName}: çalışma sayfası yok.");
                    continue;
                }

                for (var sheetIndex = 0; sheetIndex < worksheets.Count; sheetIndex++)
                {
                    var ws = worksheets[sheetIndex];
                    var source = $"{fileName} / {ws.Name}";
                    if (TryGetWorksheetHeaderNames(ws, out var headers, out var sheetError))
                        sources.Add((source, headers!, sheetIndex));
                    else
                        errors.Add($"{source}: {sheetError}");
                }
            }
        }

        return sources;
    }

    private static bool ProcessExcelWorksheet(
        IXLWorksheet ws,
        string sourceName,
        MergeOptions options,
        List<string> columnOrder,
        List<IReadOnlyDictionary<string, XLCellValue>> allRows,
        List<RowIssue> fixedRows,
        List<RowIssue> skippedRows,
        HashSet<string> seenPhoneNumbers,
        ref int duplicatesSkipped,
        ref bool reachedRowLimit,
        List<string> fileErrors)
    {
        var range = ws.RangeUsed();
        if (range is null)
        {
            fileErrors.Add($"{sourceName}: kullanılan hücre yok, atlandı.");
            return false;
        }

        var headerRangeRow = range.FirstRow();
        var headers = BuildUniqueHeaders(headerRangeRow, range.LastColumnUsed().ColumnNumber());

        foreach (var col in headers.Keys.OrderBy(c => c))
        {
            var h = headers[col];
            if (!columnOrder.Contains(h, StringComparer.OrdinalIgnoreCase))
                columnOrder.Add(h);
        }

        var firstDataRow = headerRangeRow.RowNumber() + 1;
        var lastRow = range.LastRowUsed().RowNumber();

        for (var r = firstDataRow; r <= lastRow; r++)
        {
            if (allRows.Count >= MaxDataRows)
            {
                reachedRowLimit = true;
                fileErrors.Add($"Satır limiti aşıldı ({MaxDataRows:N0}). Kalan satırlar atlandı.");
                break;
            }

            var row = ws.Row(r);
            if (IsRowEmpty(row, headers.Keys))
                continue;

            var readResult = TryReadExcelRow(sourceName, r, headers, row, options);
            switch (readResult.Outcome)
            {
                case RowReadOutcome.Success:
                    if (ShouldSkipDuplicateByPhoneNr(readResult.Row!, seenPhoneNumbers))
                    {
                        duplicatesSkipped++;
                        continue;
                    }

                    allRows.Add(readResult.Row!);
                    break;

                case RowReadOutcome.Fixed:
                    if (readResult.Issue is not null)
                        fixedRows.Add(readResult.Issue);

                    if (ShouldSkipDuplicateByPhoneNr(readResult.Row!, seenPhoneNumbers))
                    {
                        duplicatesSkipped++;
                        continue;
                    }

                    allRows.Add(readResult.Row!);
                    break;

                case RowReadOutcome.Skipped:
                    if (readResult.Issue is not null)
                        skippedRows.Add(readResult.Issue);
                    break;

                case RowReadOutcome.Failed:
                    throw new MergeRowException(readResult.Issue!);
            }
        }

        return true;
    }

    private static bool TryGetWorksheetHeaderNames(IXLWorksheet ws, out IReadOnlyList<string>? headers, out string? error)
    {
        var range = ws.RangeUsed();
        if (range is null)
        {
            headers = null;
            error = "kullanılan hücre yok.";
            return false;
        }

        var headerRow = range.FirstRow();
        var built = BuildUniqueHeaders(headerRow, range.LastColumnUsed().ColumnNumber());
        headers = GetOrderedHeaderNames(built);
        error = null;
        return true;
    }

    private static bool TryGetCsvHeaderNames(string path, out IReadOnlyList<string>? headers, out string? error)
    {
        headers = null;
        error = null;

        string[] lines;
        try
        {
            lines = File.ReadAllLines(path);
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }

        if (lines.Length == 0)
        {
            error = "CSV dosyası boş.";
            return false;
        }

        var delimiter = DetectCsvDelimiter(lines[0]);
        var headerCells = ParseCsvLine(lines[0], delimiter);
        if (headerCells.Count == 0)
        {
            error = "başlık satırı okunamadı.";
            return false;
        }

        headers = GetOrderedHeaderNames(BuildUniqueHeaders(headerCells));
        return true;
    }

    private static IReadOnlyList<string> GetOrderedHeaderNames(IReadOnlyDictionary<int, string> headers) =>
        headers.Keys.OrderBy(c => c).Select(c => headers[c]).ToList();

    private static bool HeadersEqual(IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        if (left.Count != right.Count)
            return false;

        var rightNames = new HashSet<string>(right, StringComparer.OrdinalIgnoreCase);
        foreach (var name in left)
        {
            if (!rightNames.Contains(name))
                return false;
        }

        return true;
    }

    private static string FormatHeaders(IReadOnlyList<string> headers) =>
        string.Join(", ", headers);

    private static Dictionary<int, string> BuildUniqueHeaders(IXLRangeRow headerRow, int lastCol)
    {
        var headers = new Dictionary<int, string>();
        var usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (var col = 1; col <= lastCol; col++)
        {
            var raw = headerRow.Cell(col).GetFormattedString().Trim();
            if (string.IsNullOrEmpty(raw))
                raw = $"Sütun{col}";

            var baseName = raw;
            if (!usedNames.TryGetValue(baseName, out var count))
            {
                usedNames[baseName] = 1;
                headers[col] = baseName;
            }
            else
            {
                count++;
                usedNames[baseName] = count;
                headers[col] = $"{baseName}({count})";
            }
        }

        return headers;
    }

    private static bool IsRowEmpty(IXLRow row, IEnumerable<int> columns)
    {
        foreach (var col in columns)
        {
            if (!row.Cell(col).IsEmpty(XLCellsUsedOptions.AllContents))
                return false;
        }

        return true;
    }

    private static void ProcessCsvFile(
        string path,
        string fileName,
        MergeOptions options,
        List<string> columnOrder,
        List<IReadOnlyDictionary<string, XLCellValue>> allRows,
        List<string> mergedFiles,
        List<string> fileErrors,
        List<RowIssue> fixedRows,
        List<RowIssue> skippedRows,
        HashSet<string> seenPhoneNumbers,
        ref int duplicatesSkipped)
    {
        string[] lines;
        try
        {
            lines = File.ReadAllLines(path);
        }
        catch (Exception ex)
        {
            if (options.SkipInvalidRows)
            {
                fileErrors.Add($"{fileName}: {ex.Message}");
                return;
            }

            throw;
        }

        if (lines.Length == 0)
        {
            fileErrors.Add($"{fileName}: CSV dosyası boş.");
            return;
        }

        var delimiter = DetectCsvDelimiter(lines[0]);
        var headerCells = ParseCsvLine(lines[0], delimiter);
        if (headerCells.Count == 0)
        {
            fileErrors.Add($"{fileName}: başlık satırı okunamadı.");
            return;
        }

        var headers = BuildUniqueHeaders(headerCells);
        foreach (var col in headers.Keys.OrderBy(c => c))
        {
            var h = headers[col];
            if (!columnOrder.Contains(h, StringComparer.OrdinalIgnoreCase))
                columnOrder.Add(h);
        }

        for (var i = 1; i < lines.Length; i++)
        {
            if (allRows.Count >= MaxDataRows)
            {
                fileErrors.Add($"Satır limiti aşıldı ({MaxDataRows:N0}). Kalan satırlar atlandı.");
                break;
            }

            var rowNumber = i + 1;
            try
            {
                var cells = ParseCsvLine(lines[i], delimiter);
                if (IsCsvRowEmpty(cells, headers.Count))
                    continue;

                var dict = new Dictionary<string, XLCellValue>(StringComparer.OrdinalIgnoreCase);
                foreach (var col in headers.Keys.OrderBy(c => c))
                {
                    var name = headers[col];
                    var value = col - 1 < cells.Count ? cells[col - 1] : string.Empty;
                    dict[name] = value;
                }

                if (ShouldSkipDuplicateByPhoneNr(dict, seenPhoneNumbers))
                {
                    duplicatesSkipped++;
                    continue;
                }

                allRows.Add(dict);
            }
            catch (Exception ex)
            {
                var issue = new RowIssue
                {
                    FileName = fileName,
                    RowNumber = rowNumber,
                    Reason = ex.Message,
                    Action = RowIssueAction.Skipped
                };

                if (options.SkipInvalidRows)
                {
                    skippedRows.Add(issue);
                    continue;
                }

                throw new MergeRowException(issue);
            }
        }

        mergedFiles.Add(fileName);
    }

    private static Dictionary<int, string> BuildUniqueHeaders(IReadOnlyList<string> headerCells)
    {
        var headers = new Dictionary<int, string>();
        var usedNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (var col = 1; col <= headerCells.Count; col++)
        {
            var raw = headerCells[col - 1].Trim();
            if (string.IsNullOrEmpty(raw))
                raw = $"Sütun{col}";

            var baseName = raw;
            if (!usedNames.TryGetValue(baseName, out var count))
            {
                usedNames[baseName] = 1;
                headers[col] = baseName;
            }
            else
            {
                count++;
                usedNames[baseName] = count;
                headers[col] = $"{baseName}({count})";
            }
        }

        return headers;
    }

    private static char DetectCsvDelimiter(string line)
    {
        var commaCount = CountDelimiterOutsideQuotes(line, ',');
        var semicolonCount = CountDelimiterOutsideQuotes(line, ';');
        var tabCount = CountDelimiterOutsideQuotes(line, '\t');

        if (semicolonCount > commaCount && semicolonCount >= tabCount)
            return ';';
        if (tabCount > commaCount && tabCount > semicolonCount)
            return '\t';
        return ',';
    }

    private static int CountDelimiterOutsideQuotes(string line, char delimiter)
    {
        var inQuotes = false;
        var count = 0;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    i++;
                else
                    inQuotes = !inQuotes;
            }
            else if (ch == delimiter && !inQuotes)
            {
                count++;
            }
        }

        return count;
    }

    private static List<string> ParseCsvLine(string line, char delimiter)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (ch == delimiter && !inQuotes)
            {
                result.Add(current.ToString().Trim());
                current.Clear();
            }
            else
            {
                current.Append(ch);
            }
        }

        result.Add(current.ToString().Trim());
        return result;
    }

    private static bool IsCsvRowEmpty(IReadOnlyList<string> cells, int expectedColumns)
    {
        var max = Math.Min(cells.Count, expectedColumns);
        for (var i = 0; i < max; i++)
        {
            if (!string.IsNullOrWhiteSpace(cells[i]))
                return false;
        }

        return true;
    }

    private static bool ShouldSkipDuplicateByPhoneNr(
        IReadOnlyDictionary<string, XLCellValue> rowValues,
        HashSet<string> seenPhoneNumbers)
    {
        if (!rowValues.TryGetValue(PhoneNrColumn, out var phoneValue))
            return false;

        var normalized = phoneValue.ToString().Trim();
        if (string.IsNullOrWhiteSpace(normalized))
            return false;

        return !seenPhoneNumbers.Add(normalized);
    }
}
