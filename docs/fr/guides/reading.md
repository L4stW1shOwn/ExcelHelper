# Guide — Lecture de fichiers Excel

Ce guide couvre toutes les méthodes de lecture proposées par `ExcelReader`.

## Création d'un reader

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);
```

Vous pouvez passer une configuration et contrôler la fermeture du stream :

```csharp
var config = new ExcelConfiguration { SheetName = "Products" };
using var reader = new ExcelReader(stream, config, leaveOpen: false);
```

## Lecture de tous les enregistrements

La méthode la plus courante est `GetRecords<T>()` qui retourne un `IEnumerable<T>` lazy :

```csharp
var people = reader.GetRecords<Person>().ToList();
```

```csharp
// Traitement streaming sans charger tout en mémoire
foreach (var person in reader.GetRecords<Person>())
{
    Console.WriteLine(person.Name);
}
```

## Lecture avec le curseur de lignes

Pour un contrôle granulaire, utilisez le mode curseur (`Read()` + `GetRecord<T>()`).

```csharp
using var reader = new ExcelReader(stream, new ExcelConfiguration { HasHeaderRecord = false, StartRow = 1 });

while (reader.Read())
{
    var person = reader.GetRecord<Person>();
    Console.WriteLine($"Ligne {reader.Context.Row}: {person.Name}");
}
```

### Méthodes du curseur

| Méthode | Description |
|---|---|
| `Read()` | Avance à la ligne suivante. Retourne `false` s'il n'y a plus de lignes. |
| `GetField(int index)` | Retourne la valeur brute de la cellule (index 0-based). |
| `GetField<T>(int index)` | Retourne la valeur convertie en type `T`. |
| `TryGetField<T>(int index, out T? value)` | Essaie la conversion ; retourne `false` sans jeter d'exception. |
| `GetRecord<T>()` | Matérialise la ligne courante en objet `T`. |

> **⚠️** Vous ne pouvez pas mélanger `Read()` et `GetRecords<T>()` sur le même reader.

### Exemple avec `GetField`

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

## Lecture async

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

## Ignorer les lignes vides

Par défaut, `IgnoreBlankRows = true`.

```csharp
var config = new ExcelConfiguration { IgnoreBlankRows = true };
using var reader = new ExcelReader(stream, config);

// Les lignes entièrement vides sont automatiquement sautées
var people = reader.GetRecords<Person>().ToList();
```

## Trimer les valeurs

```csharp
var config = new ExcelConfiguration { TrimCellValues = true };
using var reader = new ExcelReader(stream, config);

// "  Alice  " devient "Alice"
```

## Lire depuis une feuille spécifique

```csharp
// Par nom
var config = new ExcelConfiguration { SheetName = "Q1-Sales" };

// Par index (0-based)
var config = new ExcelConfiguration { SheetIndex = 2 };
```

## Gestion des en-têtes

Par défaut, ExcelHelper considère que la ligne 1 contient les en-têtes et que les données commencent à la ligne 2.

```csharp
// Fichier sans en-tête
var config = new ExcelConfiguration
{
    HasHeaderRecord = false,
    StartRow = 1
};
```

```csharp
// En-tête à la ligne 3, données à la ligne 5
var config = new ExcelConfiguration
{
    HeaderRow = 3,
    StartRow = 5
};
```

## Exemple complet

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
