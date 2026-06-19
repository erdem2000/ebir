using ClosedXML.Excel;

namespace ExcelBirlestirici.Tests;

public class MergeIntegrationTests
{
    [Fact]
    public void Merge_AlignsColumnsByName_AndWritesDatedOutput()
    {
        var root = Path.Combine(Path.GetTempPath(), "ExcelBirlestiriciTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        try
        {
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş"], [("Ali", 30), ("Ayşe", 25)]);
            WriteWorkbook(Path.Combine(root, "b.xlsx"), ["Yaş", "Şehir"], [(40, "Ankara"), (22, "İzmir")]);

            var result = ExcelMergeService.Merge(root);

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
            WriteWorkbook(Path.Combine(root, "a.xlsx"), ["Ad", "Yaş"], [("Ali", 30)]);
            WriteCsv(Path.Combine(root, "c.csv"), "Şehir,Yaş\nAnkara,40\nİzmir,22\n");

            var result = ExcelMergeService.Merge(root);

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

            WriteCsv(Path.Combine(root, "b.csv"), "Ad,PhoneNr\nAyşe,5551112233\nFatma,5552223344\n");

            var result = ExcelMergeService.Merge(root);
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

            var result = ExcelMergeService.Merge(root);
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

            var result = ExcelMergeService.Merge(root, new MergeOptions { SkipInvalidRows = true });

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

            var result = ExcelMergeService.Merge(root, new MergeOptions { SkipInvalidRows = false });

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
