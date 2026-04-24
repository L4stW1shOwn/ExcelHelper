# Guide — Mapping

ExcelHelper propose plusieurs stratégies pour mapper les colonnes Excel aux propriétés .NET.

## 1. Auto-mapping (par défaut)

Si aucun `ExcelClassMap` n'est enregistré, ExcelHelper mappe automatiquement les propriétés publiques dans l'ordre de déclaration.

```csharp
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public DateTime BirthDate { get; set; }
}

// Mapping automatique : Name=col0, Age=col1, BirthDate=col2
var people = reader.GetRecords<Person>().ToList();
```

## 2. Mapping par attributs

Décorez vos propriétés avec les attributs du namespace `ExcelHelper.Mapping` :

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

| Attribut | Description |
|---|---|
| `[ExcelName("...")]` | Définit le nom de colonne (utilisé pour l'en-tête). |
| `[ExcelIndex(n)]` | Définit l'index de colonne 0-based. |
| `[ExcelIgnore]` | Ignore la propriété. |
| `[ExcelDefault(value)]` | Valeur par défaut si la cellule est vide. |
| `[ExcelColumn("Name", Index = n)]` | Attribut combiné (nom, index, défaut, converter). |

## 3. Mapping par code fluent (`ExcelClassMap`)

Pour un contrôle total sans modifier les classes de modèle :

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

Enregistrez le map dans la configuration :

```csharp
var config = new ExcelConfiguration();
config.RegisterClassMap(new PersonMap());

using var reader = new ExcelReader(stream, config);
```

### Chaînage fluent

Toutes les méthodes retournent `this` pour permettre le chaînage :

```csharp
Map(m => m.Price)
    .Index(3)
    .Name("Unit Price")
    .Default(0.0m)
    .TypeConverter<DecimalConverter>()
    .Optional();
```

## 4. AutoMap avec raffinement

Vous pouvez commencer par l'auto-mapping puis surcharger certains membres :

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        AutoMap(); // Map toutes les propriétés

        // Puis personnaliser
        Map(m => m.Name).Name("Full Name");
        Map(m => m.Secret).Ignore();
    }
}
```

> **Note** : `AutoMap()` crée des `MemberMapData`. Si vous appelez `Map()` ensuite sur le même membre, vous ajoutez un **deuxième** mapping. Pour éviter les doublons, utilisez soit `AutoMap()` seul, soit `Map()` manuel sans `AutoMap()`.

## Ordre de priorité des mappings

1. `ExcelClassMap` enregistré explicitement (priorité maximale)
2. Attributs sur les propriétés (si pas de class map)
3. Auto-mapping par défaut (ordre de déclaration des propriétés)

## Exemple complet

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Mapping;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Modèle
public class Product
{
    public string Sku { get; set; } = string.Empty;
    public string Label { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Stock { get; set; }
}

// Map personnalisé
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

// Utilisation
var config = new ExcelConfiguration();
config.RegisterClassMap(new ProductMap());

using var stream = File.OpenRead("products.xlsx");
using var reader = new ExcelReader(stream, config);

var products = reader.GetRecords<Product>().ToList();
```
