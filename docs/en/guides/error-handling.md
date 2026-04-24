# Guide — Error Handling

ExcelHelper provides a structured mechanism for handling read and write errors through dedicated exceptions and callback hooks.

## Exception Hierarchy

```
ExcelHelperException (base)
├── ExcelMappingException
├── ExcelTypeConversionException
└── ExcelValidationException
```

All inherit from `ExcelHelperException`, allowing them to be caught globally:

```csharp
try
{
    var records = reader.GetRecords<Person>().ToList();
}
catch (ExcelHelperException ex)
{
    Console.WriteLine($"ExcelHelper error: {ex.Message}");
}
```

## `ExcelMappingException`

Thrown when the mapping configuration is invalid or when assigning a value fails.

| Property | Description |
|---|---|
| `MappedType` | Type currently being mapped. |
| `MemberName` | Name of the affected member. |

```csharp
catch (ExcelMappingException ex)
{
    Console.WriteLine($"Mapping failed on {ex.MemberName} of {ex.MappedType?.Name}");
}
```

## `ExcelTypeConversionException`

Thrown when converting a cell to a .NET type fails.

| Property | Description |
|---|---|
| `FieldName` | Field name. |
| `FieldValue` | Raw cell value. |
| `TargetType` | Target conversion type. |
| `Row` | 1-based row. |
| `Column` | 1-based column. |

```csharp
catch (ExcelTypeConversionException ex)
{
    Console.WriteLine($"Cannot convert '{ex.FieldValue}' to {ex.TargetType.Name} at [{ex.Row},{ex.Column}]");
}
```

## `ExcelValidationException`

Thrown when a validation rule fails.

| Property | Description |
|---|---|
| `FieldName` | Field name (null for record validation). |
| `Row` | 1-based row. |
| `Column` | 1-based column. |

```csharp
catch (ExcelValidationException ex)
{
    Console.WriteLine($"Validation failed at [{ex.Row},{ex.Column}]: {ex.Message}");
}
```

## Error Handling Hooks

Instead of catching exceptions after the fact, you can intercept issues during processing.

### `ReadingExceptionOccurred`

Return `true` to skip the row and continue; `false` to propagate the exception.

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        switch (ex)
        {
            case ExcelTypeConversionException tce:
                Console.WriteLine($"Bad value '{tce.FieldValue}' at row {tce.Row}, col {tce.Column}");
                return true; // skip the row

            case ExcelValidationException ve:
                Console.WriteLine($"Validation failed at row {ve.Row}: {ve.Message}");
                return true; // skip the row

            default:
                return false; // propagate
        }
    }
};
```

### `WritingExceptionOccurred`

Same principle for writing:

```csharp
var config = new ExcelConfiguration
{
    WritingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Write error: {ex.Message}");
        return true; // ignore and continue
    }
};
```

### `BadDataFound`

Called **before** `ExcelTypeConversionException` is thrown. Return `true` to ignore the field and continue the row.

```csharp
var config = new ExcelConfiguration
{
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data in '{args.Field}' at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignore the field, continue the row
    }
};
```

> Returning `false` in `BadDataFound` causes `ExcelTypeConversionException` to be thrown normally.

### `MissingFieldFound`

Called when a required field is empty:

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Row {row} is missing a required field.");
    }
};
```

## Full Example — Robust Read Pipeline

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var errors = new List<string>();

var config = new ExcelConfiguration
{
    TrimCellValues = true,
    IgnoreBlankRows = true,
    ReadingExceptionOccurred = ex =>
    {
        errors.Add($"[{ex.GetType().Name}] {ex.Message}");
        return true; // continue anyway
    },
    BadDataFound = args =>
    {
        errors.Add($"Bad data at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignore the field and continue the row
    },
    MissingFieldFound = (headers, row, context) =>
    {
        errors.Add($"Missing field at row {row}");
    }
};

using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream, config);

var records = reader.GetRecords<Person>().ToList();

if (errors.Count > 0)
{
    Console.WriteLine($"{errors.Count} error(s) occurred:");
    foreach (var error in errors)
        Console.WriteLine($"  - {error}");
}

Console.WriteLine($"Successfully loaded {records.Count} records.");
```
