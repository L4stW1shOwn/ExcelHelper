# API Reference — Contexts

Namespace: `ExcelHelper.Core`

## `ExcelContext`

Abstract base class providing common context for read and write operations.

```csharp
public abstract class ExcelContext
```

### Properties

| Property | Type | Description |
|---|---|---|
| `Configuration` | `ExcelConfiguration` | Configuration of the current operation. |
| `Worksheet` | `ExcelWorksheet?` | Currently processed Excel worksheet. |
| `Row` | `int` | Current 1-based row. |
| `Column` | `int` | Current 1-based column. |
| `SheetName` | `string?` | Sheet name (shortcut). |

---

## `ReadingContext`

Context specific to read operations.

```csharp
public sealed class ReadingContext : ExcelContext
```

### Properties

| Property | Type | Description |
|---|---|---|
| `HeaderRecord` | `string[]?` | Array of header names if `HasHeaderRecord` is `true`. |
| `Record` | `string[]?` | Raw values of the current record. |
| `CurrentIndex` | `int` | 0-based index of the current record in the data. |
| `RowCount` | `int` | Total number of rows in the worksheet. |

---

## `WritingContext`

Context specific to write operations.

```csharp
public sealed class WritingContext : ExcelContext
```

### Properties

| Property | Type | Description |
|---|---|---|
| `HasHeaderBeenWritten` | `bool` | Indicates whether the header has already been written. |

---

## `BadDataFoundArgs`

Arguments passed to the `ExcelConfiguration.BadDataFound` callback.

```csharp
public sealed class BadDataFoundArgs
```

### Properties

| Property | Type | Description |
|---|---|---|
| `Field` | `string?` | Name of the field concerned. |
| `RawCellValue` | `string?` | Raw value of the cell. |
| `Row` | `int` | 1-based row. |
| `Column` | `int` | 1-based column. |

### Constructor

```csharp
public BadDataFoundArgs(string? field, string? rawCellValue, int row, int column)
```

---

## Examples

### Accessing the read context

```csharp
using var reader = new ExcelReader(stream);

foreach (var person in reader.GetRecords<Person>())
{
    Console.WriteLine($"Processing row {reader.Context.Row} / {reader.Context.RowCount}");
}
```

### Accessing the write context

```csharp
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);

Console.WriteLine($"Last written row: {writer.Context.Row}");
```

### Usage in `MissingFieldFound`

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Sheet: {context.SheetName}, Row: {row}");
    }
};
```
