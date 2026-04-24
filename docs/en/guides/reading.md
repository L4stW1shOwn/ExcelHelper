# Guide — Reading Excel Files

This guide covers all reading methods provided by `ExcelReader`.

## Creating a Reader

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);
```

You can pass a configuration and control stream disposal:

```csharp
var config = new ExcelConfiguration { SheetName = "Products" };
using var reader = new ExcelReader(stream, config, leaveOpen: false);
```

## Reading All Records

The most common method is `GetRecords<T>()` which returns a lazy `IEnumerable<T>`:

```csharp
var people = reader.GetRecords<Person>().ToList();
```

```csharp
// Streaming processing without loading everything into memory
foreach (var person in reader.GetRecords<Person>())
{
    Console.WriteLine(person.Name);
}
```

## Reading with the Row Cursor

For granular control, use cursor mode (`Read()` + `GetRecord<T>()`).

```csharp
using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

while (reader.Read())
{
    var person = reader.GetRecord<Person>();
    Console.WriteLine($"Row {reader.Context.Row}: {person.Name}");
}
```

### Cursor Methods

| Method | Description |
|---|---|
| `Read()` | Advances to the next row. Returns `false` if there are no more rows. |
| `GetField(int index)` | Returns the raw cell value (0-based index). |
| `GetField<T>(int index)` | Returns the value converted to type `T`. |
| `TryGetField<T>(int index, out T? value)` | Attempts conversion; returns `false` without throwing. |
| `GetRecord<T>()` | Materializes the current row as object `T`. |

> **⚠️** You cannot mix `Read()` and `GetRecords<T>()` on the same reader.

### Example with `GetField`

```csharp
using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

while (reader.Read())
{
    var rawValue = reader.GetField(0);          // object?
    var age = reader.GetField<int>(1);          // int
    var success = reader.TryGetField<DateTime>(2, out var date);

    if (success)
        Console.WriteLine($"Date: {date}");
}
```

## Async Reading

### .NET Core+ (`IAsyncEnumerable<T>`)

```csharp
await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}
```

### .NET Framework (`Task<IReadOnlyList<T>>`)

```csharp
var people = await reader.GetRecordsAsync<Person>();
foreach (var person in people)
{
    Console.WriteLine(person.Name);
}
```

## Ignoring Empty Rows

By default, `IgnoreBlankRows = true`.

```csharp
var config = new ExcelConfiguration { IgnoreBlankRows = true };
using var reader = new ExcelReader(stream, config);

// Entirely empty rows are automatically skipped
var people = reader.GetRecords<Person>().ToList();
```

## Trimming Values

```csharp
var config = new ExcelConfiguration { TrimCellValues = true };
using var reader = new ExcelReader(stream, config);

// "  Alice  " becomes "Alice"
```

## Reading from a Specific Sheet

```csharp
// By name
var config = new ExcelConfiguration { SheetName = "Q1-Sales" };

// By index (0-based)
var config = new ExcelConfiguration { SheetIndex = 2 };
```

## Header Management

By default, ExcelHelper assumes row 1 contains headers and data starts at row 2.

```csharp
// File without header
var config = new ExcelConfiguration
{
    HasHeaderRecord = false,
    StartRow = 1
};
```

```csharp
// Header at row 3, data at row 5
var config = new ExcelConfiguration
{
    HeaderRow = 3,
    StartRow = 5
};
```

## Full Example

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;
using System.Globalization;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var config = new ExcelConfiguration
{
    SheetName = "Employees",
    TrimCellValues = true,
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR"),
    IgnoreBlankRows = true
};

using var stream = File.OpenRead("employees.xlsx");
using var reader = new ExcelReader(stream, config);

var employees = reader.GetRecords<Employee>().ToList();

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public decimal Salary { get; set; }
    public DateTime HireDate { get; set; }
}
```
