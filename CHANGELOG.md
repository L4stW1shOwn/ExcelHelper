# Changelog

All notable changes to this project will be documented in this file.

## [1.0.0] - 2026-04-23

### Added
- Initial release of ExcelHelper
- `ExcelReader` and `ExcelWriter` for XLSX read/write operations
- `ExcelClassMap<T>` and `ExcelMemberMap<TClass, TMember>` for fluent mapping
- Auto-mapping via reflection with attribute support (`[ExcelColumn]`, `[ExcelIgnore]`, `[ExcelIndex]`, `[ExcelName]`, `[ExcelDefault]`)
- Type conversion pipeline with default converters (string, int, long, double, decimal, DateTime, bool, enum, Guid, Nullable<T>)
- OADate conversion for Excel serial dates
- Field-level validation (`IExcelFieldValidator`) with built-in validators (NotNull, NotEmpty, GreaterThan, Range)
- Record-level validation (`IExcelRecordValidator`)
- Async support: `IAsyncEnumerable<T>` on .NET Core+ / netstandard2.1, `Task<IReadOnlyList<T>>` on .NET Framework / netstandard2.0
- Hooks: `ReadingExceptionOccurred`, `WritingExceptionOccurred`, `MissingFieldFound`, `BadDataFound`
- Multi-sheet support via `SheetName` and `SheetIndex`
- Header handling (`HasHeaderRecord`, `HeaderRow`)
- Blank row skipping (`IgnoreBlankRows`)
- Culture-aware conversion (`CultureInfo`)
- `CompiledExpressionCache` for performant property access
- Multi-target support: `net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0`
- Comprehensive xUnit test suite
