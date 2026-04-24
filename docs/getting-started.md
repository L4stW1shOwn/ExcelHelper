# Démarrage rapide

Ce guide vous aide à installer ExcelHelper et à exécuter votre premier exemple en quelques minutes.

## Prérequis

- .NET SDK 8.0+ (ou tout framework compatible)
- Un projet .NET Console, Web, ou Library

## Installation

```bash
dotnet add package ExcelHelper
```

> **⚠️ Licence EPPlus** : EPPlus 7+ exige que vous définissiez un contexte de licence avant toute utilisation.
>
> ```csharp
> using OfficeOpenXml;
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou LicenseContext.Commercial
> ```

## Premier exemple — Lecture

Créez un fichier Excel `people.xlsx` avec une feuille contenant :

| Name | Age | City |
|---|---|---|
| Alice | 30 | Paris |
| Bob | 25 | Lyon |

```csharp
using OfficeOpenXml;
using ExcelHelper;

// Définir la licence EPPlus (à faire une fois au démarrage)
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

using var stream = File.OpenRead("people.xlsx");
using var reader = new ExcelReader(stream);

var people = reader.GetRecords<Person>().ToList();

foreach (var person in people)
{
    Console.WriteLine($"{person.Name}, {person.Age}, {person.City}");
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string City { get; set; } = string.Empty;
}
```

Sortie :
```
Alice, 30, Paris
Bob, 25, Lyon
```

## Premier exemple — Écriture

```csharp
using OfficeOpenXml;
using ExcelHelper;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var people = new List<Person>
{
    new() { Name = "Alice", Age = 30, City = "Paris" },
    new() { Name = "Bob", Age = 25, City = "Lyon" }
};

using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);
```

Le fichier `output.xlsx` contiendra :

| Name | Age | City |
|---|---|---|
| Alice | 30 | Paris |
| Bob | 25 | Lyon |

## Prochaines étapes

- [Lecture de fichiers Excel](guides/reading.md)
- [Écriture de fichiers Excel](guides/writing.md)
- [Mapping](guides/mapping.md)
