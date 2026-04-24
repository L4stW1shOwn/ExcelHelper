using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelHelper.Core;
using ExcelHelper.Tests.Models;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests
{
    public class ExcelReaderWriterIntegrationTests
    {
        static ExcelReaderWriterIntegrationTests()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void WriteAndRead_SimpleRecords()
        {
            var people = new List<Models.Person>
            {
                new Models.Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1994, 5, 15) },
                new Models.Person { Name = "Bob", Age = 25, BirthDate = new DateTime(1999, 8, 22) }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
            {
                writer.WriteRecords(people);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
            var results = reader.GetRecords<Models.Person>().ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal(30, results[0].Age);
            Assert.Equal(new DateTime(1994, 5, 15), results[0].BirthDate);
            Assert.Equal("Bob", results[1].Name);
            Assert.Equal(25, results[1].Age);
        }

        [Fact]
        public void WriteAndRead_WithHeaders()
        {
            var products = new List<Product>
            {
                new Product { Name = "Widget", Price = 9.99m, InStock = true },
                new Product { Name = "Gadget", Price = 19.99m, InStock = false }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = true, HeaderRow = 1, StartRow = 2 }, leaveOpen: true))
            {
                writer.WriteRecords(products);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = true, HeaderRow = 1, StartRow = 2 });
            var results = reader.GetRecords<Product>().ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Widget", results[0].Name);
            Assert.Equal(9.99m, results[0].Price);
            Assert.True(results[0].InStock);
            Assert.Equal("Gadget", results[1].Name);
            Assert.Equal(19.99m, results[1].Price);
            Assert.False(results[1].InStock);
        }

        [Fact]
        public void WriteAndRead_WithCustomMapping()
        {
            var config = new ExcelConfiguration();
            config.RegisterClassMap(new PersonMap());

            var people = new List<Models.Person>
            {
                new Models.Person { Name = "Charlie", Age = 35, BirthDate = new DateTime(1989, 3, 10) }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, config, leaveOpen: true))
            {
                writer.WriteRecords(people);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, config);
            var results = reader.GetRecords<Models.Person>().ToList();

            Assert.Single(results);
            Assert.Equal("Charlie", results[0].Name);
            Assert.Equal(35, results[0].Age);
        }

        [Fact]
        public void WriteAndRead_WithIgnoredProperty()
        {
            var people = new List<PersonWithIgnore>
            {
                new PersonWithIgnore { Name = "Alice", Age = 30, Secret = "TopSecret" }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
            {
                writer.WriteRecords(people);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
            var results = reader.GetRecords<PersonWithIgnore>().ToList();

            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal(30, results[0].Age);
            Assert.Empty(results[0].Secret); // Should not be written/read, remains default
        }

        [Fact]
        public void WriteAndRead_Dates()
        {
            var dates = new List<DateRecord>
            {
                new DateRecord { EventDate = new DateTime(2024, 1, 15), EventName = "Launch" }
            };

            using var stream = new MemoryStream();
            using (var writer = new ExcelWriter(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 }, leaveOpen: true))
            {
                writer.WriteRecords(dates);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });
            var results = reader.GetRecords<DateRecord>().ToList();

            Assert.Single(results);
            Assert.Equal(new DateTime(2024, 1, 15), results[0].EventDate);
            Assert.Equal("Launch", results[0].EventName);
        }

        [Fact]
        public void WriteAndRead_MultiSheet()
        {
            var people = new List<Models.Person>
            {
                new Models.Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1994, 5, 15) }
            };

            using var stream = new MemoryStream();

            // Write to "People" sheet
            var writeConfig = new ExcelConfiguration { SheetName = "People", HasHeaderRecord = false, StartRow = 1 };
            using (var writer = new ExcelWriter(stream, writeConfig, leaveOpen: true))
            {
                writer.WriteRecords(people);
            }

            stream.Position = 0;

            // Read from "People" sheet
            var readConfig = new ExcelConfiguration { SheetName = "People", HasHeaderRecord = false, StartRow = 1 };
            using var reader = new ExcelReader(stream, readConfig);
            var results = reader.GetRecords<Models.Person>().ToList();

            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
        }

        [Fact]
        public void Read_BlankRows_ShouldBeIgnored()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 30;
                ws.Cells[2, 1].Value = null; // blank row
                ws.Cells[2, 2].Value = null;
                ws.Cells[3, 1].Value = "Bob";
                ws.Cells[3, 2].Value = 25;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1, IgnoreBlankRows = true });
            var results = reader.GetRecords<SimplePerson>().ToList();

            Assert.Equal(2, results.Count);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal("Bob", results[1].Name);
        }
    }
}
