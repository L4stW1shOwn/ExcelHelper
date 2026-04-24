# Référence API — Attributs de mapping

Namespace : `ExcelHelper.Mapping`

## `[ExcelColumn]`

Attribut combiné permettant de définir nom, index, valeur par défaut et convertisseur.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelColumnAttribute : Attribute
```

### Propriétés

| Propriété | Type | Défaut | Description |
|---|---|---|---|
| `Name` | `string?` | `null` | Nom de la colonne. |
| `Index` | `int` | `-1` | Index 0-based de la colonne. |
| `Default` | `object?` | `null` | Valeur par défaut si cellule vide. |
| `Converter` | `Type?` | `null` | Type du `IExcelTypeConverter` à utiliser. |

### Constructeurs

```csharp
public ExcelColumnAttribute()
public ExcelColumnAttribute(string name)
```

### Exemple

```csharp
[ExcelColumn("Product Name", Index = 0, Default = "N/A")]
public string Name { get; set; } = string.Empty;
```

---

## `[ExcelName]`

Définit le nom de la colonne.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelNameAttribute : Attribute
```

### Constructeur

```csharp
public ExcelNameAttribute(string name)
```

### Exemple

```csharp
[ExcelName("Full Name")]
public string Name { get; set; } = string.Empty;
```

---

## `[ExcelIndex]`

Définit l'index de colonne 0-based.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelIndexAttribute : Attribute
```

### Constructeur

```csharp
public ExcelIndexAttribute(int index)
```

### Exemple

```csharp
[ExcelIndex(5)]
public int Priority { get; set; }
```

---

## `[ExcelIgnore]`

Ignore la propriété lors du mapping.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelIgnoreAttribute : Attribute
```

### Exemple

```csharp
[ExcelIgnore]
public string InternalId { get; set; } = string.Empty;
```

---

## `[ExcelDefault]`

Définit la valeur par défaut si la cellule est vide.

```csharp
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ExcelDefaultAttribute : Attribute
```

### Constructeur

```csharp
public ExcelDefaultAttribute(object? value)
```

### Exemple

```csharp
[ExcelDefault(42)]
public int Quantity { get; set; }
```

---

## Exemple combiné

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
