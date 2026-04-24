# Guide — Validation

ExcelHelper permet de valider les données **par champ** et **par enregistrement** lors de la lecture.

## Validation par champ (`IExcelFieldValidator`)

Créez un validateur en implémentant `IExcelFieldValidator` :

```csharp
using ExcelHelper.Validation;

public class NotEmptyValidator : IExcelFieldValidator
{
    public ValidationResult Validate(object? value, string? fieldName, int row, int column)
    {
        if (value == null || (value is string s && string.IsNullOrWhiteSpace(s)))
            return ValidationResult.Failed($"Field '{fieldName}' cannot be empty.");

        return ValidationResult.Success();
    }
}
```

### Validateurs intégrés

ExcelHelper fournit plusieurs validateurs prêts à l'emploi :

| Validateur | Description |
|---|---|
| `NotNullValidator` | Vérifie que la valeur n'est pas `null`. |
| `NotEmptyValidator` | Vérifie que la chaîne n'est pas vide ou blanche. |
| `GreaterThanValidator(double min)` | Valeur numérique > `min`. |
| `RangeValidator(double min, double max)` | Valeur numérique dans [`min`, `max`]. |

### Utilisation dans un `ExcelClassMap`

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(0, 120));
    }
}
```

### Utilisation avec un validateur instance

Utile quand le validateur a besoin de paramètres d'initialisation :

```csharp
Map(m => m.Email)
    .Index(2)
    .Validate(new RegexValidator(@"^[^@]+@[^@]+$"));
```

## Validation par enregistrement (`IExcelRecordValidator<T>`)

Pour valider la cohérence globale d'une ligne (plusieurs champs ensemble) :

```csharp
public class PersonRecordValidator : IExcelRecordValidator<Person>
{
    public ValidationResult Validate(Person record, int row)
    {
        if (record.Age < 18 && record.Country == "US")
            return ValidationResult.Failed($"Row {row}: Must be 18+ in the US.");

        return ValidationResult.Success();
    }
}
```

Enregistrez-le dans le class map :

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0);
        Map(m => m.Age).Index(1);
        Map(m => m.Country).Index(2);

        Validate<PersonRecordValidator>();       // par type
        Validate(new PersonRecordValidator());   // par instance
    }
}
```

## `ValidationResult`

```csharp
// Succès
return ValidationResult.Success();

// Échec
return ValidationResult.Failed("Message d'erreur détaillé");
```

En cas d'échec, une `ExcelValidationException` est levée avec :
- `Message` : le message d'erreur
- `FieldName` : le nom du champ (null pour la validation d'enregistrement)
- `Row` / `Column` : position 1-based dans la feuille

## Gestion des erreurs de validation

Vous pouvez intercepter les erreurs via le hook `ReadingExceptionOccurred` :

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        if (ex is ExcelValidationException vex)
        {
            Console.WriteLine($"Validation failed at row {vex.Row}, col {vex.Column}: {vex.Message}");
            return true; // ignorer et continuer
        }
        return false; // propager l'exception
    }
};
```

## Exemple complet

```csharp
using OfficeOpenXml;
using ExcelHelper;
using ExcelHelper.Mapping;
using ExcelHelper.Validation;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public decimal Salary { get; set; }
}

public class EmployeeMap : ExcelClassMap<Employee>
{
    public EmployeeMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(18, 65));
        Map(m => m.Salary).Index(2).Validate(new GreaterThanValidator(0));

        Validate(new SalaryAgeValidator());
    }
}

public class SalaryAgeValidator : IExcelRecordValidator<Employee>
{
    public ValidationResult Validate(Employee record, int row)
    {
        if (record.Age < 21 && record.Salary > 100000)
            return ValidationResult.Failed($"Row {row}: Salary too high for age {record.Age}.");

        return ValidationResult.Success();
    }
}

// Utilisation
var config = new ExcelConfiguration();
config.RegisterClassMap(new EmployeeMap());
config.ReadingExceptionOccurred = ex =>
{
    Console.WriteLine($"Error: {ex.Message}");
    return true; // continuer malgré les erreurs
};

using var stream = File.OpenRead("employees.xlsx");
using var reader = new ExcelReader(stream, config);

var employees = reader.GetRecords<Employee>().ToList();
```
