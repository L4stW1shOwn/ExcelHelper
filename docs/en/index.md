# ExcelHelper Documentation

Welcome to the **ExcelHelper** documentation, the .NET library for reading and writing Excel (XLSX) files with a fluent API inspired by CsvHelper.

## Table of Contents

### User Guides

- [Getting Started](getting-started.md) — Installation and first example
- [Reading Excel Files](guides/reading.md) — `ExcelReader`, row cursor, streaming
- [Writing Excel Files](guides/writing.md) — `ExcelWriter`, headers, batch writing
- [Mapping](guides/mapping.md) — `ExcelClassMap`, attributes, auto-mapping
- [Validation](guides/validation.md) — Field and record validators
- [Type Conversion](guides/type-converters.md) — Built-in and custom converters
- [Configuration](guides/configuration.md) — `ExcelConfiguration`, culture, sheets
- [Async](guides/async.md) — `IAsyncEnumerable<T>` and `Task<IReadOnlyList<T>>`
- [Error Handling](guides/error-handling.md) — Exceptions, hooks, `BadDataFound`

### API Reference

- [ExcelReader](api-reference/ExcelReader.md)
- [ExcelWriter](api-reference/ExcelWriter.md)
- [ExcelConfiguration](api-reference/ExcelConfiguration.md)
- [ExcelClassMap&lt;T&gt;](api-reference/ExcelClassMap.md)
- [ExcelMemberMap&lt;TClass, TMember&gt;](api-reference/ExcelMemberMap.md)
- [Mapping Attributes](api-reference/attributes.md)
- [Validation](api-reference/validation.md)
- [Exceptions](api-reference/exceptions.md)
- [Contexts](api-reference/contexts.md)
- [Type Converters](api-reference/type-converters.md)

## Feature Overview

| Feature | Description |
|---|---|
| XLSX Read / Write | Based on EPPlus 7+ |
| Object Mapping | Via fluent code (`ExcelClassMap`) or attributes |
| Auto-mapping | Automatic mapping of public properties |
| Type Conversion | `string`, `int`, `long`, `double`, `decimal`, `bool`, `DateTime`, `Guid`, `enum`, nullables |
| Validation | Per field and per record |
| Streaming | `IAsyncEnumerable<T>` (.NET Core+) or `Task<IReadOnlyList<T>>` (.NET Framework) |
| Multi-target | `net10.0`, `net9.0`, `net8.0`, `net48`, `net47`, `netstandard2.1`, `netstandard2.0` |
| Multi-sheet | Read/write by sheet name or index |
| OADate | Automatic Excel date conversion |

## Installation

```bash
dotnet add package ExcelHelper
```

> **Important**: EPPlus 7+ requires a license context. Set it before using ExcelHelper:
>
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or Commercial
> ```
