using ClosedXML.Excel;

namespace ExcelBirlestirici.Tests;

public class MergeIntegrationTests
{
    private static MergeOptions ClassicOptions(int sheetsToMerge = 1, bool skipInvalidRows = false) =>
        new()
        {
            WorkbookProtectedMode = false,
            SheetsToMerge = sheetsToMerge,
            SkipInvalidRows = skipInvalidRows
        };

    [Fact]
    public void Merge_AlignsColumnsByName_AndWritesDatedOutput()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş", "Şehir"], [("Ali", 30, ""), ("Ayşe", 25, "")]);
            WriteWorkbook(Path.Combine(root, "b.xlsx"), ["Yaş", "Şehir", "Ad"], [(40, "Ankara", "Mehmet"), (22, "İzmir", "Can")]);

            var result = ExcelMergeService.Merge(root, ClassicOptions());

            Assert.NotNull(result.OutputPath);
            Assert.True(File.Exists(result.OutputPath));
            Assert.StartsWith("Birleştirilmiş_", Path.GetFileName(result.OutputPath), StringComparison.OrdinalIgnoreCase);
            Assert.Contains(DateTime.Now.ToString("yyyyMMdd"), Path.GetFileName(result.OutputPath));

            using var wb = new XLWorkbook(result.OutputPath!);
            var ws = wb.Worksheet("Birleştirilmiş");
            Assert.Equal("Ad", ws.Cell(1, 1).GetString());
            Assert.Equal("Yaş", ws.Cell(1, 2).GetString());
            Assert.Equal("Şehir", ws.Cell(1, 3).GetString());

            Assert.Equal("Ali", ws.Cell(2, 1).GetString());
            Assert.Equal(30, ws.Cell(2, 2).GetDouble());
            Assert.True(ws.Cell(2, 3).IsEmpty());

            Assert.Equal(4, result.RowsWritten);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_IncludesCsvFiles_ByHeaderNames()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciCsvTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş", "Şehir"], [("Ali", 30, "")]);
            WriteCsv(Path.Combine(root, "c.csv"), "Şehir,Yaş,Ad\nAnkara,40,Zeynep\nİzmir,22,Deniz\n");

            var result = ExcelMergeService.Merge(root, ClassicOptions());

            Assert.NotNull(result.OutputPath);
            Assert.Contains("c.csv", result.DiscoveredFiles);
            Assert.Contains("c.csv", result.MergedFiles);
            Assert.Equal(3, result.RowsWritten);

            using var wb = new XLWorkbook(result.OutputPath!);
            var ws = wb.Worksheet("Birleştirilmiş");
            Assert.Equal("Ad", ws.Cell(1, 1).GetString());
            Assert.Equal("Yaş", ws.Cell(1, 2).GetString());
            Assert.Equal("Şehir", ws.Cell(1, 3).GetString());
            Assert.Equal("Ali", ws.Cell(2, 1).GetString());
            Assert.Equal("Ankara", ws.Cell(3, 3).GetString());
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_SkipsRowsWithDuplicatePhoneNr()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciPhoneTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(
                Path.Combine(root, "a.xlsx"),
                ["PhoneNr", "Ad"],
                [("5551112233", "Ali"), ("5551112233", "Veli")]);

            WriteCsv(Path.Combine(root, "b.csv"), "PhoneNr,Ad\n5551112233,Ayşe\n5552223344,Fatma\n");

            var result = ExcelMergeService.Merge(root, ClassicOptions());
            Assert.NotNull(result.OutputPath);
            Assert.Equal(2, result.RowsWritten);
            Assert.Equal(2, result.DuplicatesSkipped);

            using var wb = new XLWorkbook(result.OutputPath!);
            var ws = wb.Worksheet("Birleştirilmiş");
            Assert.Equal("5551112233", ws.Cell(2, 1).GetString());
            Assert.Equal("5552223344", ws.Cell(3, 1).GetString());
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_ParsesSemicolonSeparatedCsv_IntoColumns()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciSemicolonCsvTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteCsv(Path.Combine(root, "semi.csv"), "Ad;Şehir;PhoneNr\nAli;Ankara;5551112233\nVeli;İzmir;5552223344\n");

            var result = ExcelMergeService.Merge(root, ClassicOptions());
            Assert.NotNull(result.OutputPath);
            Assert.Equal(2, result.RowsWritten);

            using var wb = new XLWorkbook(result.OutputPath!);
            var ws = wb.Worksheet("Birleştirilmiş");
            Assert.Equal("Ad", ws.Cell(1, 1).GetString());
            Assert.Equal("Şehir", ws.Cell(1, 2).GetString());
            Assert.Equal("PhoneNr", ws.Cell(1, 3).GetString());
            Assert.Equal("Ali", ws.Cell(2, 1).GetString());
            Assert.Equal("Ankara", ws.Cell(2, 2).GetString());
            Assert.Equal("5551112233", ws.Cell(2, 3).GetString());
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_MergesFirstNSheets_WhenHeadersMatch()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciMultiSheetTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbookWithSheets(
                Path.Combine(root, "book.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ali", 30)]),
                ("Sayfa2", ["Ad", "Yaş"], [("Veli", 40)]));

            var result = ExcelMergeService.Merge(root, ClassicOptions(sheetsToMerge: 2));

            Assert.NotNull(result.OutputPath);
            Assert.Equal(2, result.RowsWritten);
            Assert.Contains("book.xlsx / Sayfa1", result.MergedFiles);
            Assert.Contains("book.xlsx / Sayfa2", result.MergedFiles);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_Aborts_WhenSheetHeadersDiffer()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciHeaderMismatchTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbookWithSheets(
                Path.Combine(root, "book.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ali", 30)]),
                ("Sayfa2", ["Ad", "Şehir"], [("Veli", "Ankara")]));

            var validation = ExcelMergeService.ValidateSheetHeaders(root, ClassicOptions(sheetsToMerge: 2));
            Assert.False(validation.IsValid);
            Assert.NotEmpty(validation.Errors);

            var result = ExcelMergeService.Merge(root, ClassicOptions(sheetsToMerge: 2));
            Assert.Null(result.OutputPath);
            Assert.True(result.AbortedDueToHeaderMismatch);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_WorkbookProtectedMode_MergesBySheetIndex()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciProtectedByIndex_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbookWithSheets(
                Path.Combine(root, "a.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ali", 30)]),
                ("Sayfa2", ["Ad", "Yaş"], [("Veli", 40)]));
            WriteWorkbookWithSheets(
                Path.Combine(root, "b.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ayşe", 25)]),
                ("Sayfa2", ["Ad", "Yaş"], [("Can", 35)]));

            var result = ExcelMergeService.Merge(root, new MergeOptions { SheetsToMerge = 2, WorkbookProtectedMode = true });

            Assert.NotNull(result.OutputPath);
            Assert.Equal(4, result.RowsWritten);

            using var wb = new XLWorkbook(result.OutputPath!);
            Assert.Equal(2, wb.Worksheets.Count);
            Assert.Equal("Birleştirilmiş 1", wb.Worksheet(1).Name);
            Assert.Equal("Birleştirilmiş 2", wb.Worksheet(2).Name);
            Assert.Equal("Ali", wb.Worksheet(1).Cell(2, 1).GetString());
            Assert.Equal("Ayşe", wb.Worksheet(1).Cell(3, 1).GetString());
            Assert.Equal("Veli", wb.Worksheet(2).Cell(2, 1).GetString());
            Assert.Equal("Can", wb.Worksheet(2).Cell(3, 1).GetString());
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_WorkbookProtectedMode_ValidatesHeadersPerSheetIndex()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciProtectedHeader_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbookWithSheets(
                Path.Combine(root, "a.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ali", 30)]),
                ("Sayfa2", ["Ad", "Yaş"], [("Veli", 40)]));
            WriteWorkbookWithSheets(
                Path.Combine(root, "b.xlsx"),
                ("Sayfa1", ["Ad", "Yaş"], [("Ayşe", 25)]),
                ("Sayfa2", ["Ad", "Şehir"], [("Can", "Ankara")]));

            var validation = ExcelMergeService.ValidateSheetHeaders(root, new MergeOptions { SheetsToMerge = 2, WorkbookProtectedMode = true });
            Assert.False(validation.IsValid);
            Assert.Contains(validation.Errors, e => e.Contains("Sayfa 2", StringComparison.Ordinal));

            var result = ExcelMergeService.Merge(root, new MergeOptions { SheetsToMerge = 2, WorkbookProtectedMode = true });
            Assert.Null(result.OutputPath);
            Assert.True(result.AbortedDueToHeaderMismatch);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_WorkbookProtectedMode_SkipsCsvFiles()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciProtectedCsv_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş"], [("Ali", 30)]);
            WriteCsv(Path.Combine(root, "data.csv"), "Ad,Yaş\nVeli,40\n");

            var result = ExcelMergeService.Merge(root, new MergeOptions { WorkbookProtectedMode = true });

            Assert.NotNull(result.OutputPath);
            Assert.DoesNotContain(result.MergedFiles, f => f.Contains("data.csv", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(result.FileErrors, e => e.Contains("korumalı modda CSV atlandı", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(1, result.RowsWritten);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_WorkbookProtectedMode_DefaultIsTrue()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciProtectedDefault_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş"], [("Ali", 30)]);
            WriteWorkbook(Path.Combine(root, "b.xlsx"), ["Ad", "Yaş"], [("Ayşe", 25)]);

            var result = ExcelMergeService.Merge(root);

            Assert.NotNull(result.OutputPath);
            Assert.Equal(2, result.RowsWritten);

            using var wb = new XLWorkbook(result.OutputPath!);
            Assert.Single(wb.Worksheets);
            Assert.Equal("Birleştirilmiş 1", wb.Worksheet(1).Name);
            Assert.Equal("Ali", wb.Worksheet(1).Cell(2, 1).GetString());
            Assert.Equal("Ayşe", wb.Worksheet(1).Cell(3, 1).GetString());
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    private static void WriteWorkbook(string path, string[] headers, (object c1, object c2, object c3)[] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sayfa1");
        for (var c = 0; c < headers.Length; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        for (var r = 0; r < rows.Length; r++)
        {
            SetCellValue(ws.Cell(r + 2, 1), rows[r].c1);
            SetCellValue(ws.Cell(r + 2, 2), rows[r].c2);
            SetCellValue(ws.Cell(r + 2, 3), rows[r].c3);
        }

        wb.SaveAs(path);
    }

    private static void WriteWorkbook(string path, string[] headers, (object c1, object c2)[] rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sayfa1");
        for (var c = 0; c < headers.Length; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        for (var r = 0; r < rows.Length; r++)
        {
            SetCellValue(ws.Cell(r + 2, 1), rows[r].c1);
            SetCellValue(ws.Cell(r + 2, 2), rows[r].c2);
        }

        wb.SaveAs(path);
    }

    private static void WriteWorkbookWithSheets(
        string path,
        params (string SheetName, string[] Headers, (object c1, object c2)[] Rows)[] sheets)
    {
        using var wb = new XLWorkbook();
        foreach (var sheet in sheets)
        {
            var ws = wb.AddWorksheet(sheet.SheetName);
            for (var c = 0; c < sheet.Headers.Length; c++)
                ws.Cell(1, c + 1).Value = sheet.Headers[c];

            for (var r = 0; r < sheet.Rows.Length; r++)
            {
                SetCellValue(ws.Cell(r + 2, 1), sheet.Rows[r].c1);
                SetCellValue(ws.Cell(r + 2, 2), sheet.Rows[r].c2);
            }
        }

        wb.SaveAs(path);
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case int i:
                cell.Value = i;
                break;
            case long l:
                cell.Value = l;
                break;
            case double d:
                cell.Value = d;
                break;
            case string s:
                cell.Value = s;
                break;
            case null:
                cell.Clear();
                break;
            default:
                cell.Value = value.ToString() ?? string.Empty;
                break;
        }
    }

    [Fact]
    public void Merge_SkipInvalidRows_SkipsRowsWithInvalidDate()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciSkipInvalidTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteInvalidDateWorkbook(Path.Combine(root, "bad-date.xlsx"));

            var result = ExcelMergeService.Merge(root, ClassicOptions(skipInvalidRows: true));

            Assert.NotNull(result.OutputPath);
            Assert.Equal(2, result.RowsWritten);
            Assert.Single(result.SkippedRows);
            Assert.Empty(result.FixedRows);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    [Fact]
    public void Merge_StrictMode_RepairsInvalidDateRows()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciFixInvalidTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteInvalidDateWorkbook(Path.Combine(root, "bad-date.xlsx"));

            var result = ExcelMergeService.Merge(root, ClassicOptions());

            Assert.NotNull(result.OutputPath);
            Assert.Equal(3, result.RowsWritten);
            Assert.True(result.FixedRows.Count >= 1);
            Assert.False(result.AbortedDueToInvalidRow);
        }
        finally
        {
            try
            {
                Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore cleanup failures on locked temp files
            }
        }
    }

    private static void WriteInvalidDateWorkbook(string path)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sayfa1");
        ws.Cell(1, 1).Value = "Ad";
        ws.Cell(1, 2).Value = "Tarih";
        ws.Cell(2, 1).Value = "Ali";
        ws.Cell(2, 2).Value = new DateTime(2024, 1, 1);

        ws.Cell(3, 1).Value = "Veli";
        var badDateCell = ws.Cell(3, 2);
        badDateCell.SetValue(2_958_466.0);
        badDateCell.Style.DateFormat.Format = "dd.MM.yyyy";

        ws.Cell(4, 1).Value = "Ayşe";
        ws.Cell(4, 2).Value = new DateTime(2024, 2, 1);

        wb.SaveAs(path);
    }

    private static void WriteCsv(string path, string content)
    {
        File.WriteAllText(path, content);
    }
}
