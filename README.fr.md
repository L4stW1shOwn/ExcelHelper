# ExcelHelper

Une bibliothèque de lecture et d'écriture Excel (XLSX) pour .NET, architecturalement inspirée par **CsvHelper** mais construite nativement pour Excel en utilisant **EPPlus 7+**.

ExcelHelper fournit une API fluide et familière pour mapper les colonnes Excel vers des objets .NET, avec un support complet de la conversion de types, de la validation, du streaming asynchrone et de la compatibilité multi-framework.

---

## Installation

```bash
dotnet add package ExcelHelper
```

> **Note** : EPPlus 7+ nécessite un contexte de licence. Définissez-le avant d'utiliser ExcelHelper :
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // ou Commercial
> ```

---

## Démarrage rapide

### Lecture

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = reader.GetRecords<Person>().ToList();
```

### Écriture

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);
```

### Mappage

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Name("Full Name");
        Map(m => m.Age).Index(1).Default(-1);
        Map(m => m.BirthDate).Index(2).TypeConverter<DateTimeConverter>();
    }
}

var config = new ExcelConfiguration();
config.RegisterClassMap(new PersonMap());
```

### Mappage basé sur les attributs

```csharp
public class Person
{
    [ExcelName("Full Name")]
    public string Name { get; set; }

    [ExcelIndex(1)]
    public int Age { get; set; }

    [ExcelIgnore]
    public string Secret { get; set; }
}
```

### Validation

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(0, 120));
        Validate(new PersonRecordValidator());
    }
}
```

### Asynchrone

```csharp
// .NET Core+
await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}

// .NET Framework
var people = await reader.GetRecordsAsync<Person>();

// Écriture asynchrone
await writer.WriteRecordsAsync(GetProductsAsync());
```

---

## Compatibilité .NET

| TFM | `IAsyncEnumerable<T>` | `Task<IReadOnlyList<T>>` | EPPlus 7+ |
|---|---|---|---|
| `net10.0` | ✅ | ✅ | ✅ |
| `net9.0` | ✅ | ✅ | ✅ |
| `net8.0` | ✅ | ✅ | ✅ |
| `netstandard2.1` | ✅ | ✅ | ✅ |
| `netstandard2.0` | ❌ | ✅ | ✅ |
| `net48` | ❌ | ✅ | ✅ |
| `net47` | ❌ | ✅ | ✅ |

---

## Comparaison avec CsvHelper

| Fonctionnalité | CsvHelper | ExcelHelper |
|---|---|---|
| `ClassMap<T>` | ✅ | ✅ `ExcelClassMap<T>` |
| `MemberMap` | ✅ | ✅ `ExcelMemberMap<TClass, TMember>` |
| Auto-mapping | ✅ | ✅ `AutoMap` |
| Mappage par attributs | ✅ | ✅ `[ExcelColumn]`, `[ExcelIgnore]`, etc. |
| Convertisseurs de types | ✅ | ✅ `IExcelTypeConverter<T>` |
| Validation | ✅ | ✅ `IExcelFieldValidator`, `IExcelRecordValidator` |
| `IAsyncEnumerable<T>` | ✅ | ✅ (Core+) / `Task<IReadOnlyList<T>>` (Framework) |
| Hooks (`ReadingExceptionOccurred`) | ✅ | ✅ |
| En-têtes | ✅ | ✅ |
| Multi-feuilles | ❌ | ✅ |
| Formules / Styles | ❌ | ✅ (via EPPlus) |
| Conversion OADate | ❌ | ✅ |

---

## Architecture

Voir [ARCHITECTURE.md](ARCHITECTURE.md) pour la documentation détaillée de la conception.

## Contribuer

Voir [CONTRIBUTING.md](CONTRIBUTING.md) pour les directives.

## Feuille de route

Voir [ROADMAP.md](ROADMAP.md) pour les phases d'implémentation.

## Licence

Licence MIT. Voir LICENSE pour les détails.

EPPlus est sous licence PolyForm Noncommercial 1.0.0 (gratuit pour un usage non commercial) ou sous une licence commerciale.
