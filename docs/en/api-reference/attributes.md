# API Reference — Mapping Attributes

Namespace: `ExcelHelper.Mapping`

## `[ExcelColumn]`

Combined attribute allowing you to define name, index, default value and converter.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelColumnAttribute : Attribute
```

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Name` | `string?` | `null` | Column name. |
| `Index` | `int` | `-1` | 0-based column index. |
| `Default` | `object?` | `null` | Default value if cell is empty. |
| `Converter` | `Type?` | `null` | Type of the `IExcelTypeConverter` to use. |

### Constructors

```csharp
public ExcelColumnAttribute()
public ExcelColumnAttribute(string name)
```

### Example

```csharp
[ExcelColumn("Product Name", Index = 0, Default = "N/A")]
public string Name { get; set; } = string.Empty;
```

---

## `[ExcelName]`

Defines the column name.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelNameAttribute : Attribute
```

### Constructor

```csharp
public ExcelNameAttribute(string name)
```

### Example

```csharp
[ExcelName("Full Name")]
public string Name { get; set; } = string.Empty;
```

---

## `[ExcelIndex]`

Defines the 0-based column index.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelIndexAttribute : Attribute
```

### Constructor

```csharp
public ExcelIndexAttribute(int index)
```

### Example

```csharp
[ExcelIndex(5)]
public int Priority { get; set; }
```

---

## `[ExcelIgnore]`

Ignores the property during mapping.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelIgnoreAttribute : Attribute
```

### Example

```csharp
[ExcelIgnore]
public string InternalId { get; set; } = string.Empty;
```

---

## `[ExcelDefault]`

Defines the default value if the cell is empty.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelDefaultAttribute : Attribute
```

### Constructor

```csharp
public ExcelDefaultAttribute(object? value)
```

### Example

```csharp
[ExcelDefault(42)]
public int Quantity { get; set; }
```

---

## Combined Example

```csharp
public class Product
{
    [ExcelName("SKU")]
    public string Code { get; set; } = string.Empty;

    [ExcelIndex(1)]
    public string Name { get; set; } = string.Empty;

    [ExcelDefault(0)]
    public int Stock { get; set; }

    [ExcelIgnore]
    public string TemporaryData { get; set; } = string.Empty;

    [ExcelColumn("Price", Index = 3, Default = 0.0)]
    public decimal Price { get; set; }
}
```
