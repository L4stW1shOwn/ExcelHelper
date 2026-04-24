# Guide — Mapping

ExcelHelper offers several strategies for mapping Excel columns to .NET properties.

## 1. Auto-mapping (default)

If no `ExcelClassMap` is registered, ExcelHelper automatically maps public properties in declaration order.

```csharp
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public DateTime BirthDate { get; set; }
}

// Automatic mapping: Name=col0, Age=col1, BirthDate=col2
var people = reader.GetRecords<Person>().ToList();
```

## 2. Attribute Mapping

Decorate your properties with attributes from the `ExcelHelper.Mapping` namespace:

```csharp
using ExcelHelper.Mapping;

public class Person
{
    [ExcelName("Full Name")]
    public string Name { get; set; } = string.Empty;

    [ExcelIndex(2)]
    public int Age { get; set; }

    [ExcelIgnore]
    public string Secret { get; set; } = string.Empty;

    [ExcelDefault(42)]
    public int DefaultValue { get; set; }

    [ExcelColumn("Department", Index = 3)]
    public string Dept { get; set; } = string.Empty;
}
```

| Attribute | Description |
|---|---|
| `[ExcelName("...")]` | Sets the column name (used for the header). |
| `[ExcelIndex(n)]` | Sets the 0-based column index. |
| `[ExcelIgnore]` | Ignores the property. |
| `[ExcelDefault(value)]` | Default value if the cell is empty. |
| `[ExcelColumn("Name", Index = n)]` | Combined attribute (name, index, default, converter). |

## 3. Fluent Code Mapping (`ExcelClassMap`)

For full control without modifying model classes:

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name)
            .Index(0)
            .Name("Full Name")
            .Default("Unknown");

        Map(m => m.Age)
            .Index(1)
            .Optional();

        Map(m => m.BirthDate)
            .Index(2)
            .TypeConverter<DateTimeConverter>();

        Map(m => m.Secret).Ignore();
    }
}
```

Register the map in the configuration:

```csharp
var config = new ExcelConfiguration();
config.RegisterClassMap(new PersonMap());

using var reader = new ExcelReader(stream, config);
```

### Fluent Chaining

All methods return `this` to enable chaining:

```csharp
Map(m => m.Price)
    .Index(3)
    .Name("Unit Price")
    .Default(0.0m)
    .TypeConverter<DecimalConverter>()
    .Optional();
```

## 4. AutoMap with Refinement

You can start with auto-mapping and then override specific members:

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        AutoMap(); // Map all properties

        // Then customize
        Map(m => m.Name).Name("Full Name");
        Map(m => m.Secret).Ignore();
    }
}
```

> **Note**: `AutoMap()` creates `MemberMapData`. If you call `Map()` afterwards on the same member, you add a **second** mapping. To avoid duplicates, use either `AutoMap()` alone, or manual `Map()` without `AutoMap()`.

## Mapping Priority Order

1. Explicitly registered `ExcelClassMap` (highest priority)
2. Attributes on properties (if no class map)
3. Default auto-mapping (property declaration order)

## Full Example

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Mapping;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Model
public class Product
{
    public string Sku { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Stock { get; set; }
}

// Custom map
public class ProductMap : ExcelClassMap<Product>
{
    public ProductMap()
    {
        Map(m => m.Sku).Index(0).Name("SKU");
        Map(m => m.Label).Index(1).Name("Product Label");
        Map(m => m.Price).Index(2).Name("Price (EUR)").Default(0m);
        Map(m => m.Stock).Index(3).Name("Qty in Stock").Default(0);
    }
}

// Usage
var config = new ExcelConfiguration();
config.RegisterClassMap(new ProductMap());

using var stream = File.OpenRead("products.xlsx");
using var reader = new ExcelReader(stream, config);

var products = reader.GetRecords<Product>().ToList();
```
