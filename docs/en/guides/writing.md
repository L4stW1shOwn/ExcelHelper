# Guide — Writing Excel Files

This guide covers all writing methods provided by `ExcelWriter`.

## Creating a Writer

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);
```

With configuration:

```csharp
var config = new ExcelConfiguration { SheetName = "Products" };
using var writer = new ExcelWriter(stream, config);
```

## Writing a Collection

```csharp
var people = new List<Person>
{
    new() { Name = "Alice", Age = 30 },
    new() { Name = "Bob", Age = 25 }
};

writer.WriteRecords(people);
```

The header is written automatically if `HasHeaderRecord = true` (default).

## Writing a Single Record

```csharp
writer.WriteRecord(new Person { Name = "Charlie", Age = 35 });
writer.WriteRecord(new Person { Name = "Diana", Age = 28 });
```

Each call writes a row and advances the internal cursor.

## Async Writing

### .NET Core+ (`IAsyncEnumerable<T>`)

```csharp
async IAsyncEnumerable<Product> GetProductsAsync()
{
    yield return new Product { Name = "Widget", Price = 9.99m };
    yield return new Product { Name = "Gadget", Price = 19.99m };
}

await writer.WriteRecordsAsync(GetProductsAsync());
```

### .NET Framework (`IEnumerable<T>`)

```csharp
await writer.WriteRecordsAsync(products);
```

## Custom Headers

Column names in the header come from the mapping:

- Default property name
- `ExcelNameAttribute` or `ExcelColumnAttribute`
- `.Name("...")` in `ExcelClassMap`

```csharp
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(m => m.Name).Index(0).Name("Product Name");
        Map(m => m.Price).Index(1).Name("Unit Price");
    }
}
```

The generated header will be:

| Product Name | Unit Price |
|---|---|

## Multiple Sheets

By default, ExcelWriter creates a sheet named "Sheet1". You can specify:

```csharp
var config = new ExcelConfiguration { SheetName = "Sales" };
using var writer = new ExcelWriter(stream, config);
```

## Manual Header Writing

```csharp
writer.WriteHeader<Product>();
```

This writes the header immediately without waiting for the first `WriteRecord`. Useful if you need to insert rows before the data.

## Full Example

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Core;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var config = new ExcelConfiguration
{
    SheetName = "Invoices",
    StartRow = 2,
    HasHeaderRecord = true
};

var invoices = new List<Invoice>
{
    new() { Number = "INV-001", Amount = 1500.00m, Date = new DateTime(2024, 1, 15) },
    new() { Number = "INV-002", Amount = 2300.50m, Date = new DateTime(2024, 2, 10) }
};

using var stream = File.OpenWrite("invoices.xlsx");
using var writer = new ExcelWriter(stream, config);

writer.WriteRecords(invoices);

public class Invoice
{
    public string Number { get; set; } = string.Empty;
    public decimal Amount { get; set; }
    public DateTime Date { get; set; }
}
```

## Important Points

- The Excel package is saved to the stream when `Dispose()` is called.
- If `leaveOpen = false` (default), the stream is closed automatically.
- Exceptions during writing can be intercepted via `ExcelConfiguration.WritingExceptionOccurred`.
