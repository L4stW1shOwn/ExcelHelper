using System;
using System.IO;
using System.Linq;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Mapping;
using ExcelHelper.Tests.Models;
using ExcelHelper.Validation;
using OfficeOpenXml;
using Xunit;

namespace ExcelHelper.Tests
{
    public class ExcelValidationTests
    {
        static ExcelValidationTests()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void FieldValidator_NotNull_Should_Throw_When_Null()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = null; // null name
                ws.Cells[1, 2].Value = 25;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var ex = Assert.Throws<ExcelValidationException>(() => reader.GetRecords<ValidatedPerson>().ToList());
            Assert.Contains("cannot be null", ex.Message);
        }

        [Fact]
        public void FieldValidator_NotEmpty_Should_Throw_When_Empty()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "   "; // whitespace
                ws.Cells[1, 2].Value = 25;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var ex = Assert.Throws<ExcelValidationException>(() => reader.GetRecords<ValidatedPerson>().ToList());
            Assert.Contains("cannot be empty", ex.Message);
        }

        [Fact]
        public void FieldValidator_GreaterThan_Should_Throw_When_Out_Of_Range()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 0; // Age must be > 0
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var ex = Assert.Throws<ExcelValidationException>(() => reader.GetRecords<ValidatedPerson>().ToList());
            Assert.Contains("greater than 0", ex.Message);
        }

        [Fact]
        public void FieldValidator_Range_Should_Throw_When_Out_Of_Range()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = 150; // must be 0-100
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new AgeRangeMap());

            using var reader = new ExcelReader(stream, config);
            var ex = Assert.Throws<ExcelValidationException>(() => reader.GetRecords<AgeRangeRecord>().ToList());
            Assert.Contains("between 0 and 100", ex.Message);
        }

        [Fact]
        public void FieldValidator_Should_Pass_When_Valid()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 25;
                ws.Cells[1, 3].Value = 50000;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var results = reader.GetRecords<ValidatedPerson>().ToList();

            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal(25, results[0].Age);
        }

        [Fact]
        public void RecordValidator_Should_Throw_When_Invalid()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 25;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new RecordValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var ex = Assert.Throws<ExcelValidationException>(() => reader.GetRecords<ValidatedPerson>().ToList());
            Assert.Contains("Salary must be greater than 0", ex.Message);
        }

        [Fact]
        public void RecordValidator_Should_Pass_When_Valid()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 25;
                ws.Cells[1, 3].Value = 50000;
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 };
            config.RegisterClassMap(new RecordValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var results = reader.GetRecords<ValidatedPerson>().ToList();

            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
        }

        [Fact]
        public void BadDataFound_Callback_Should_Be_Invoked()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = "NotANumber"; // invalid int for Age
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var callbackInvoked = false;
            var config = new ExcelConfiguration
            {
                HasHeaderRecord = false,
                StartRow = 1,
                BadDataFound = args =>
                {
                    callbackInvoked = true;
                    return true; // ignore and continue
                }
            };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            var results = reader.GetRecords<ValidatedPerson>().ToList();

            Assert.True(callbackInvoked);
            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
            Assert.Equal(0, results[0].Age); // default value because bad data was ignored
        }

        [Fact]
        public void BadDataFound_Callback_Should_Throw_When_Returns_False()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = "NotANumber"; // invalid int for Age
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration
            {
                HasHeaderRecord = false,
                StartRow = 1,
                BadDataFound = args => false // do not ignore
            };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            Assert.Throws<ExcelTypeConversionException>(() => reader.GetRecords<ValidatedPerson>().ToList());
        }

        [Fact]
        public void MissingFieldFound_Callback_Should_Be_Invoked()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Name"; // header
                ws.Cells[1, 2].Value = "Age";  // header
                ws.Cells[2, 1].Value = "Alice";
                // Age missing (column 2 empty)
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var callbackInvoked = false;
            var config = new ExcelConfiguration
            {
                HasHeaderRecord = true,
                HeaderRow = 1,
                StartRow = 2,
                MissingFieldFound = (headers, row, context) =>
                {
                    callbackInvoked = true;
                }
            };
            config.RegisterClassMap(new ValidatedPersonMap());

            using var reader = new ExcelReader(stream, config);
            reader.GetRecords<ValidatedPerson>().ToList();

            Assert.True(callbackInvoked);
        }

        [Fact]
        public void ReadingExceptionOccurred_Should_Continue_When_Returns_True()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "Alice";
                ws.Cells[1, 2].Value = 25;
                ws.Cells[2, 1].Value = "Bob";
                ws.Cells[2, 2].Value = "Invalid"; // invalid int
                package.SaveAs(stream);
            }

            stream.Position = 0;
            var config = new ExcelConfiguration
            {
                HasHeaderRecord = false,
                StartRow = 1,
                ReadingExceptionOccurred = ex => true // ignore all exceptions
            };

            using var reader = new ExcelReader(stream, config);
            var results = reader.GetRecords<ValidatedPerson>().ToList();

            Assert.Single(results);
            Assert.Equal("Alice", results[0].Name);
        }

        [Fact]
        public void ValidationResult_Success_Should_Be_Valid()
        {
            var result = ValidationResult.Success();
            Assert.True(result.IsValid);
            Assert.Null(result.ErrorMessage);
        }

        [Fact]
        public void ValidationResult_Failed_Should_Be_Invalid()
        {
            var result = ValidationResult.Failed("Error");
            Assert.False(result.IsValid);
            Assert.Equal("Error", result.ErrorMessage);
        }
    }
}
