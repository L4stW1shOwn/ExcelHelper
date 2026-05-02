using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Mapping;
using ExcelHelper.Tests.Models;
using ExcelHelper.Validation;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelWriterTests
{
    static ExcelWriterTests()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    [Fact]
    public void WriteRecords_Should_Support_Enumerable()
    {
        IEnumerable<Product> GetProducts()
        {
            yield return new Product { Name = "A", Price = 1m, InStock = true };
            yield return new Product { Name = "B", Price = 2m, InStock = false };
        }

        using var stream = new MemoryStream();
        using (var writer =
               new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, true))
        {
            writer.WriteRecords(GetProducts());
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
        var results = reader.GetRecords<Product>().ToList();

        Assert.Equal(2, results.Count);
    }

    [Fact]
    public void Writer_Should_Return_Context()
    {
        using var stream = new MemoryStream();
        using var writer = new ExcelWriter(stream,
            new ExcelConfiguration { SheetName = "Output", HasHeaderRecord = false, StartRow = 1 });
        Assert.NotNull(writer.Context);
        Assert.Equal("Output", writer.Context.SheetName);
    }

#if NETCOREAPP || NETSTANDARD2_1
        [Fact]
        public async Task WriteRecordsAsync_Should_Accept_AsyncEnumerable()
        {
            async IAsyncEnumerable<Product> GetProductsAsync()
            {
                await Task.Yield();
                yield return new Product { Name = "Apple", Price = 1.99m };
                await Task.Yield();
                yield return new Product { Name = "Banana", Price = 0.99m };
            }

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow =
 1 }, leaveOpen: true))
            {
                await writer.WriteRecordsAsync(GetProductsAsync());
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow =
 1 });
            var results = reader.GetRecords<Product>().ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Apple", results[0].Name);
            Assert.Equal("Banana", results[1].Name);
        }
#else
    [Fact]
    public async Task WriteRecordsAsync_Should_Write_Records()
    {
        var products = new List<Product>
        {
            new() { Name = "Apple", Price = 1.99m },
            new() { Name = "Banana", Price = 0.99m }
        };

        using var stream = new MemoryStream();
        using (var writer =
               new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, true))
        {
            await writer.WriteRecordsAsync(products);
        }

        stream.Position = 0;
        using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
        var results = reader.GetRecords<Product>().ToList();

        Assert.Equal(2, results.Count);
        Assert.Equal("Apple", results[0].Name);
        Assert.Equal("Banana", results[1].Name);
    }
#endif

    [Fact]
    public void WriteRecords_Sets_Context_State_Per_Field()
    {
        var config = new ExcelConfiguration();
        using var stream = new MemoryStream();
        using var writer = new ExcelWriter(stream, config);

        var records = new[] { new SimplePerson { Name = "Alice", Age = 30 } };
        writer.WriteRecords(records);

        Assert.NotNull(writer.Context);
        Assert.IsType<WritingContext>(writer.Context);
        // After writing one record, the context should reflect the last processed field state
        Assert.Equal("Age", writer.Context.CurrentFieldName);
        Assert.Equal(30, writer.Context.CurrentFieldValue);
    }

    [Fact]
    public void WriteRecords_WritingExceptionOccurred_Receives_EventArgs_With_Context()
    {
        var config = new ExcelConfiguration();
        WritingExceptionEventArgs? captured = null;
        config.WritingExceptionOccurred = args =>
        {
            captured = args;
            return true; // ignore
        };

        // Register a class map with a validator that always fails to trigger the exception path
        var map = new ExcelClassMap<Person>();
        map.AutoMap();
        map.MemberMaps.First(m => m.Name == "Age").Validators.Add(new AlwaysFailValidator());
        config.Maps.Add(map);

        using var stream = new MemoryStream();
        using var writer = new ExcelWriter(stream, config);
        var records = new[] { new Person { Name = "Alice", Age = 30 } };
        writer.WriteRecords(records);

        Assert.NotNull(captured);
        Assert.IsType<ExcelValidationException>(captured.Exception);
        Assert.NotNull(captured.WritingContext);
    }

    [Fact]
    public void WriteRecord_Single_Exception_Contains_WritingContext()
    {
        var config = new ExcelConfiguration();
        // Register a class map with a validator that always fails
        var map = new ExcelClassMap<Person>();
        map.AutoMap();
        map.MemberMaps.First(m => m.Name == "Age").Validators.Add(new AlwaysFailValidator());
        config.Maps.Add(map);

        using var stream = new MemoryStream();
        using var writer = new ExcelWriter(stream, config);

        var ex = Assert.Throws<ExcelValidationException>(() =>
            writer.WriteRecord(new Person { Name = "Alice", Age = 30 }));

        Assert.NotNull(ex.Context);
        Assert.IsType<WritingContext>(ex.Context);
    }
}

public class AlwaysFailValidator : IExcelFieldValidator
{
    public ValidationResult Validate(ValidateArgs args)
    {
        return ValidationResult.Failed("Always fails for testing.");
    }
}