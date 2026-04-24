using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelHelper.Core;
using ExcelHelper.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests
{
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
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
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
            using var writer = new ExcelWriter(stream, new ExcelConfiguration { SheetName = "Output", HasHeaderRecord = false, StartRow = 1 });
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
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
            {
                await writer.WriteRecordsAsync(GetProductsAsync());
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
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
                new Product { Name = "Apple", Price = 1.99m },
                new Product { Name = "Banana", Price = 0.99m }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
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
    }
}
