# ExcelHelper Architecture

## Overview

ExcelHelper is an Excel (XLSX) reading and writing library for .NET, architecturally inspired by CsvHelper but adapted to the Excel domain. It is built on top of EPPlus 7+.

The library follows **SOLID principles** with a strict separation of concerns between mapping, type conversion, validation, and I/O.

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                        Consumer Code                        │
├─────────────────────────────────────────────────────────────┤
│  ExcelReader / ExcelWriter                                  │
│  ├─ ExcelConfiguration (culture, headers, callbacks)        │
│  ├─ ExcelContext (ReadingContext / WritingContext)          │
│  └─ ExcelClassMap<T> (mapping definitions)                  │
├─────────────────────────────────────────────────────────────┤
│  Pipeline                                                   │
│  ├─ Mapping Layer     (ExcelMemberMap, AutoMap)             │
│  ├─ Conversion Layer  (IExcelTypeConverter, Cache)          │
│  ├─ Validation Layer  (IExcelFieldValidator, etc.)          │
│  └─ I/O Layer         (EPPlus ExcelWorksheet)               │
├─────────────────────────────────────────────────────────────┤
│  Infrastructure                                             │
│  ├─ ReflectionHelper                                        │
│  ├─ CompiledExpressionCache                                 │
│  └─ SharedStringTableHelper                                 │
└─────────────────────────────────────────────────────────────┘
```

## Data Flow

### Reading Pipeline

```
ExcelWorksheet
    ↓
Enumerate rows (yield)
    ↓
Extract cell value (Text / Value / Formula)
    ↓
OADate conversion (if DateTime and UseOADate=true)
    ↓
TypeConverter (raw value → target .NET type)
    ↓
Field Validation
    ↓
Property Assignment (compiled expression)
    ↓
Record Validation
    ↓
yield return T
```

### Writing Pipeline

```
T object
    ↓
Property read (compiled expression getter)
    ↓
TypeConverter (.NET type → Excel-compatible value)
    ↓
Cell assignment (Value / Formula / Style)
    ↓
Optional row flush
    ↓
Save to stream
```

## Core Modules

### 1. Configuration (`Core/`)

`ExcelConfiguration` is the central immutable-like configuration object. It holds:
- Culture info for conversions
- Header options (row, presence)
- Sheet targeting (name or index)
- Callbacks (`ReadingExceptionOccurred`, `BadDataFound`, etc.)
- Behavior flags (`UseOADate`, `TrimCellValues`, `IgnoreBlankRows`)

### 2. Mapping (`Mapping/`)

`ExcelClassMap<T>` defines how a .NET type maps to Excel columns.
- `Map(m => m.Property)` creates an `ExcelMemberMap`
- Member maps support: `Index()`, `Name()`, `Default()`, `TypeConverter<>()`, `Ignore()`, `Optional()`
- `AutoMap` generates maps via reflection, respecting `[ExcelColumn]`, `[ExcelIgnore]`, etc.
- `CompiledExpressionCache` stores compiled getters/setters for performance

### 3. Type Conversion (`TypeConversion/`)

`IExcelTypeConverter<T>` handles conversion between Excel cell values and .NET types.
- Default converters for: `string`, `int`, `long`, `double`, `decimal`, `DateTime`, `bool`, `enum`, `Guid`, `Nullable<T>`
- `ExcelTypeConverterCache` caches converter instances
- `OADateConverter` handles Excel's serial date format

### 4. Validation (`Validation/`)

Two-level validation:
- **Field validation**: per-cell, after conversion
- **Record validation**: per-row, after object materialization

### 5. Exceptions (`Exceptions/`)

Hierarchical exceptions:
- `ExcelHelperException` (base)
- `ExcelValidationException`
- `ExcelTypeConversionException`
- `ExcelMappingException`

## Extension Points

| Extension Point | Interface / Base Class | Registration |
|---|---|---|
| Custom type converter | `IExcelTypeConverter<T>` | `ExcelConfiguration.TypeConverterCache` |
| Custom field validator | `IExcelFieldValidator` | Fluent API on `ExcelMemberMap` |
| Custom record validator | `IExcelRecordValidator` | Fluent API on `ExcelClassMap` |
| Custom object resolver | `IObjectResolver` | `ExcelConfiguration.ObjectResolver` |
| Custom mapping | Inherit `ExcelClassMap<T>` | Register via `ExcelConfiguration.RegisterClassMap` |

## Multi-Framework Strategy

```xml
<TargetFrameworks>net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0</TargetFrameworks>
```

### Compatibility Matrix

| TFM | `IAsyncEnumerable<T>` | `Span<T>` | `init` properties | Notes |
|---|---|---|---|---|
| `net10.0` | ✅ Native | ✅ Native | ✅ Native | Latest |
| `net9.0` | ✅ Native | ✅ Native | ✅ Native | |
| `net8.0` | ✅ Native | ✅ Native | ✅ Native | LTS |
| `netstandard2.1` | ✅ Native | ❌ No | ❌ No | |
| `netstandard2.0` | ❌ Fallback `Task<IReadOnlyList<T>>` | ❌ No | ❌ No | Requires `Microsoft.Bcl.AsyncInterfaces` |
| `net48` / `net47` | ❌ Fallback `Task<IReadOnlyList<T>>` | ❌ No | ❌ No | Requires `Microsoft.Bcl.AsyncInterfaces` |

### Compilation Symbols

Use `#if NETCOREAPP` for .NET Core+ specific APIs (Span, IAsyncEnumerable).
Use `#if NETFRAMEWORK` for .NET Framework specific workarounds.

## Performance Strategy

1. **Reflection cache**: `ConcurrentDictionary` for property info lookups
2. **Compiled expressions**: `Expression.Lambda` for getters/setters, cached per type
3. **Converter cache**: Singleton converters stored in `ExcelTypeConverterCache`
4. **Batch processing**: Read/write in configurable batch sizes
5. **Allocation reduction**: Reuse `StringBuilder`, avoid boxing in converters

## Key Technical Decisions

1. **EPPlus 7+**: Modern API, active maintenance. The caller is responsible for `ExcelPackage.LicenseContext`.
2. **DOM-based processing**: EPPlus loads the entire workbook. Streaming is simulated via `yield return` on records, not on the underlying file.
3. **Dual async API**: Different return types per TFM to maximize compatibility without forcing polyfills on modern runtimes.
4. **Fluent API**: All mapping and validation configuration supports method chaining.
5. **Separation of context**: `ReadingContext` and `WritingContext` carry state separately to avoid cross-contamination.

## File Structure

```
src/ExcelHelper/
├── Core/              # Configuration, context, pipeline args
├── Mapping/           # Class maps, member maps, auto-mapping, attributes
├── TypeConversion/    # Converters, factory, cache, OADate
├── Validation/        # Field validators, record validators, fluent extensions
├── Exceptions/        # Exception hierarchy
├── Internal/          # Reflection helpers, expression cache, shared strings
└── Extensions/        # Extension methods for EPPlus types
```

