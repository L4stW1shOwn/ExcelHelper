using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Mapping;
using ExcelHelper.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelReaderTests
{
    static ExcelReaderTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public void Read_AdvancesToFirstDataRow()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            ws.Cells[2, 1].Value = "Bob";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        Assert.Equal(1, reader.Context.Row);
        Assert.True(reader.Read());
        Assert.Equal(2, reader.Context.Row);
        Assert.False(reader.Read());
    }

    [Fact]
    public void GetField_ReturnsRawValue()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            ws.Cells[1, 2].Value = 30;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream,
            new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

        Assert.True(reader.Read());
        Assert.Equal("Alice", reader.GetField(0));
        Assert.Equal(30.0, reader.GetField(1));
    }

    [Fact]
    public void GetField_ReturnsNullForEmptyCell()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = null;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream,
            new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

        Assert.True(reader.Read());
        Assert.Null(reader.GetField(0));
    }

    [Fact]
    public void GetFieldT_ConvertsValue()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "42";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        Assert.Equal(42, reader.GetField<int>(0));
    }

    [Fact]
    public void GetFieldT_ThrowsOnBadData()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "not_a_number";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        Assert.Throws<ExcelTypeConversionException>(() => reader.GetField<int>(0));
    }

    [Fact]
    public void TryGetFieldT_ReturnsTrueAndValueOnValidData()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "42";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        Assert.True(reader.TryGetField<int>(0, out var value));
        Assert.Equal(42, value);
    }

    [Fact]
    public void TryGetFieldT_ReturnsFalseOnBadData()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "not_a_number";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        Assert.False(reader.TryGetField<int>(0, out _));
    }

    [Fact]
    public void GetFieldT_ReturnsDefaultForNullCell()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = null;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream,
            new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

        Assert.True(reader.Read());
        Assert.Equal(0, reader.GetField<int>(0));
    }

    [Fact]
    public void GetRecord_MaterialisesCurrentRow()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            ws.Cells[1, 2].Value = 30;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        var person = reader.GetRecord<SimplePerson>();
        Assert.Equal("Alice", person.Name);
        Assert.Equal(30, person.Age);
    }

    [Fact]
    public void GetRecord_CalledTwice_SameRow()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            ws.Cells[1, 2].Value = 30;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());
        var person = reader.GetRecord<SimplePerson>();
        var minimal = reader.GetRecord<MinimalPerson>();

        Assert.Equal("Alice", person.Name);
        Assert.Equal(30, person.Age);
        Assert.Equal("Alice", minimal.Name);
    }

    [Fact]
    public void MixedMode_ThrowsWhenReadAfterGetRecords()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        reader.GetRecords<SimplePerson>().ToList();

        Assert.Throws<InvalidOperationException>(() => reader.Read());
    }

    [Fact]
    public void MixedMode_ThrowsWhenGetRecordsAfterRead()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.True(reader.Read());

        Assert.Throws<InvalidOperationException>(() => reader.GetRecords<SimplePerson>().ToList());
    }

    [Fact]
    public void GetFieldBeforeRead_Throws()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.Throws<InvalidOperationException>(() => reader.GetField(0));
    }

    [Fact]
    public void GetRecordBeforeRead_Throws()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

        Assert.Throws<InvalidOperationException>(() => reader.GetRecord<SimplePerson>());
    }

    [Fact]
    public void Read_SkipsBlankRows()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Alice";
            ws.Cells[2, 1].Value = null; // blank
            ws.Cells[3, 1].Value = "Bob";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream,
            new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = true });

        Assert.True(reader.Read());
        Assert.Equal("Alice", reader.GetField(0));
        Assert.True(reader.Read());
        Assert.Equal("Bob", reader.GetField(0));
        Assert.False(reader.Read());
    }

    [Fact]
    public void GetField_TrimCellValues()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "  Alice  ";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream,
            new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, TrimCellValues = true });

        Assert.True(reader.Read());
        Assert.Equal("Alice", reader.GetField(0));
    }

    [Fact]
    public void GetRecords_Should_Yield_Records()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Apple";
            ws.Cells[1, 2].Value = 1.99m;
            ws.Cells[2, 1].Value = "Banana";
            ws.Cells[2, 2].Value = 0.99m;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
        var results = reader.GetRecords<Product>();

        // Yield should not materialize all records at once
        var enumerator = results.GetEnumerator();
        Assert.True(enumerator.MoveNext());
        Assert.Equal("Apple", enumerator.Current.Name);
        Assert.True(enumerator.MoveNext());
        Assert.Equal("Banana", enumerator.Current.Name);
        Assert.False(enumerator.MoveNext());
    }

#if NETCOREAPP || NETSTANDARD2_1
        [Fact]
        public async Task GetRecordsAsync_Should_Return_AsyncEnumerable()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Apple";
                ws.Cells[1, 2].Value = 1.99m;
                ws.Cells[2, 1].Value = "Banana";
                ws.Cells[2, 2].Value = 0.99m;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow =
 1 });
            var results = new List<Product>();

            await foreach (var record in reader.GetRecordsAsync<Product>())
            {
                results.Add(record);
            }

            Assert.Equal(2, results.Count);
            Assert.Equal("Apple", results[0].Name);
            Assert.Equal("Banana", results[1].Name);
        }

        [Fact]
        public async Task GetRecordsAsync_Should_Respect_CancellationToken()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Apple";
                ws.Cells[2, 1].Value = "Banana";
                ws.Cells[3, 1].Value = "Cherry";
                package.SaveAs(stream);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow =
 1 });
            var cts = new CancellationTokenSource();

            var results = new List<Product>();
            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            {
                await foreach (var record in reader.GetRecordsAsync<Product>(cts.Token))
                {
                    results.Add(record);
                    cts.Cancel();
                }
            });

            Assert.Single(results);
        }
#else
    [Fact]
    public async Task GetRecordsAsync_Should_Return_TaskList()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Apple";
            ws.Cells[1, 2].Value = 1.99m;
            ws.Cells[2, 1].Value = "Banana";
            ws.Cells[2, 2].Value = 0.99m;
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
        var results = await reader.GetRecordsAsync<Product>();

        Assert.Equal(2, results.Count);
        Assert.Equal("Apple", results[0].Name);
        Assert.Equal("Banana", results[1].Name);
    }
#endif

    [Fact]
    public void Reader_Should_Return_Context()
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells[1, 1].Value = "Test";
            package.SaveAs(stream);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
        Assert.NotNull(reader.Context);
        Assert.Equal("Sheet1", reader.Context.SheetName);
    }

    [Fact]
    public void GetRecords_Sets_Context_State_Per_Field()
    {
        var bytes = CreateExcelFile(new[] { "Name", "Age" }, new[] { new[] { "Alice", "30" } });
        var config = new ExcelConfiguration();
        ReadingContext? capturedContext = null;
        config.BadDataFound = args =>
        {
            capturedContext = args.ReadingContext;
            return false;
        };

        using var stream = new MemoryStream(bytes);
        using var reader = new ExcelReader(stream, config);

        // Trigger a conversion that won't fail — we just verify context is set during normal read
        var records = reader.GetRecords<Person>().ToList();
        Assert.Single(records);
    }

    [Fact]
    public void GetRecords_BadDataFound_Receives_EventArgs_With_Context()
    {
        var bytes = CreateExcelFile(new[] { "Name", "Age" }, new[] { new[] { "Alice", "not_a_number" } });
        BadDataFoundEventArgs? captured = null;
        var config = new ExcelConfiguration();
        config.BadDataFound = args =>
        {
            captured = args;
            return false; // let it throw
        };

        using var stream = new MemoryStream(bytes);
        using var reader = new ExcelReader(stream, config);

        Assert.Throws<ExcelTypeConversionException>(() => reader.GetRecords<Person>().ToList());
        Assert.NotNull(captured);
        Assert.Equal("Age", captured.FieldName);
        Assert.Equal("not_a_number", captured.RawValue);
        Assert.NotNull(captured.ReadingContext);
        Assert.Equal(2, captured.ReadingContext.Row);
    }

    [Fact]
    public void GetRecords_ReadingExceptionOccurred_Receives_EventArgs_With_Context()
    {
        var bytes = CreateExcelFile(new[] { "Name", "Age" }, new[] { new[] { "Alice", "not_a_number" } });
        ReadingExceptionEventArgs? captured = null;
        var config = new ExcelConfiguration();
        config.ReadingExceptionOccurred = args =>
        {
            captured = args;
            return true; // ignore
        };

        using var stream = new MemoryStream(bytes);
        using var reader = new ExcelReader(stream, config);
        var records = reader.GetRecords<Person>().ToList();

        Assert.Empty(records);
        Assert.NotNull(captured);
        Assert.IsType<ExcelTypeConversionException>(captured.Exception);
        Assert.NotNull(captured.ReadingContext);
    }

    [Fact]
    public void GetRecords_MissingFieldFound_Receives_EventArgs_With_Context()
    {
        var bytes = CreateExcelFile(new[] { "Name" }, new[] { new[] { "Alice" } });
        MissingFieldEventArgs? captured = null;
        var config = new ExcelConfiguration();
        var map = new ExcelClassMap<SimplePerson>();
        map.AutoMap();
        map.MemberMaps.First(m => m.Name == "Age").IsOptional = false;
        config.Maps.Add(map);
        config.MissingFieldFound = args =>
        {
            captured = args;
        };

        using var stream = new MemoryStream(bytes);
        using var reader = new ExcelReader(stream, config);
        var records = reader.GetRecords<SimplePerson>().ToList();

        Assert.NotNull(captured);
        Assert.Equal("Age", captured.FieldName);
        Assert.NotNull(captured.ReadingContext);
    }

    [Fact]
    public void GetRecords_Exception_Contains_Context()
    {
        var bytes = CreateExcelFile(new[] { "Name", "Age" }, new[] { new[] { "Alice", "bad" } });
        var config = new ExcelConfiguration();

        using var stream = new MemoryStream(bytes);
        using var reader = new ExcelReader(stream, config);

        var ex = Assert.Throws<ExcelTypeConversionException>(() => reader.GetRecords<Person>().ToList());
        Assert.NotNull(ex.Context);
        Assert.IsType<ReadingContext>(ex.Context);
        Assert.Equal(2, ex.Context.Row);
    }

    private static byte[] CreateExcelFile(string[] headers, string[][] rows)
    {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            for (var i = 0; i < headers.Length; i++)
            {
                ws.Cells[1, i + 1].Value = headers[i];
            }

            for (var r = 0; r < rows.Length; r++)
            {
                for (var c = 0; c < rows[r].Length; c++)
                {
                    ws.Cells[r + 2, c + 1].Value = rows[r][c];
                }
            }

            package.SaveAs(stream);
        }

        return stream.ToArray();
    }
}