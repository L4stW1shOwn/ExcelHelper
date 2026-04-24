# Guide — Configuration

`ExcelConfiguration` centralizes all read and write settings.

## Main Properties

### Sheet Management

```csharp
var config = new ExcelConfiguration
{
    SheetName = "Sales",   // Read/write to the named sheet
    SheetIndex = 0         // Or use the 0-based index (default: 0)
};
```

> `SheetName` takes priority over `SheetIndex`.

### Headers and Data Rows

```csharp
var config = new ExcelConfiguration
{
    HasHeaderRecord = true,   // The sheet has a header row (default: true)
    HeaderRow = 1,            // 1-based header row (default: 1)
    StartRow = 2              // 1-based row where data starts (default: 2)
};
```

### Data Cleaning

```csharp
var config = new ExcelConfiguration
{
    TrimCellValues = true,      // Trim whitespace around strings (default: false)
    IgnoreBlankRows = true,     // Ignore entirely empty rows (default: true)
    IgnoreReferences = false    // Ignore Excel error references (#REF!, etc.)
};
```

### Culture and Dates

```csharp
var config = new ExcelConfiguration
{
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR"), // Culture for conversions (default: Invariant)
    UseOADate = true                                    // Interpret numbers as OADate (default: true)
};
```

## Hooks (callbacks)

### `ReadingExceptionOccurred`

Intercepts exceptions during reading. Return `true` to ignore and continue.

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Read error: {ex.Message}");
        return true; // skip the row
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
        return false; // propagate the exception
    }
};
```

### `MissingFieldFound`

Called when a required field is missing:

```csharp
var config = new ExcelConfiguration
{
    MissingFieldFound = (headers, row, context) =>
    {
        Console.WriteLine($"Missing field at row {row}");
    }
};
```

### `BadDataFound`

Called when a conversion fails before the exception is thrown:

```csharp
var config = new ExcelConfiguration
{
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data in field '{args.Field}' at [{args.Row},{args.Column}]: '{args.RawCellValue}'");
        return true; // ignore and continue
    }
};
```

### `PrepareHeaderForMatch`

Transforms header names before comparison:

```csharp
var config = new ExcelConfiguration
{
    PrepareHeaderForMatch = header => header?.Trim().ToUpperInvariant() ?? string.Empty
};
```

## Object Resolver

Customize instance creation during mapping:

```csharp
public class ServiceResolver : IObjectResolver
{
    private readonly IServiceProvider _services;
    public ServiceResolver(IServiceProvider services) => _services = services;

    public object Resolve(Type type)
    {
        return _services.GetService(type) ?? Activator.CreateInstance(type)!;
    }
}

var config = new ExcelConfiguration
{
    ObjectResolver = new ServiceResolver(serviceProvider)
};
```

## Registering Class Maps

```csharp
var config = new ExcelConfiguration();
config.RegisterClassMap(new PersonMap());
config.RegisterClassMap(new OrderMap());
```

The method returns `this` for chaining:

```csharp
var config = new ExcelConfiguration()
    .RegisterClassMap(new PersonMap())
    .RegisterClassMap(new OrderMap());
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
    SheetName = "Data",
    HasHeaderRecord = true,
    HeaderRow = 1,
    StartRow = 2,
    TrimCellValues = true,
    IgnoreBlankRows = true,
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR"),
    UseOADate = true,
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine($"Skipped row: {ex.Message}");
        return true;
    },
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data at {args.Row},{args.Column}: {args.RawCellValue}");
        return true;
    }
};

using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream, config);

var records = reader.GetRecords<MyRecord>().ToList();
```
