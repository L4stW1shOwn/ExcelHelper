using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests
{
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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = false });

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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = true });

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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, TrimCellValues = true });

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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
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
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
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
    }
}
