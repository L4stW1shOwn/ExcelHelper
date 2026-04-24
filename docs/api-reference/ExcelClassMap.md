# Référence API — ExcelClassMap&lt;T&gt;

Namespace : `ExcelHelper.Mapping`

## Description

Définit le mapping entre une classe .NET et les colonnes Excel.

## Propriétés

| Propriété | Type | Description |
|---|---|---|
| `MemberMaps` | `ExcelMemberMapCollection` | Collection des mappings de membres. |
| `RecordValidators` | `List<object>` | Validateurs d'enregistrement enregistrés. |

## Méthodes

### `Map<TMember>`

```csharp
public ExcelMemberMap<T, TMember> Map<TMember>(Expression<Func<T, TMember>> expression)
```

Déclare le mapping pour un membre (propriété ou champ) de la classe.

Retourne un `ExcelMemberMap<T, TMember>` pour la configuration fluent.

---

### `Validate<TValidator>`

```csharp
public ExcelClassMap<T> Validate<TValidator>() where TValidator : IExcelRecordValidator<T>, new()
```

Ajoute un validateur d'enregistrement par type.

---

### `Validate`

```csharp
public ExcelClassMap<T> Validate(IExcelRecordValidator<T> validator)
```

Ajoute un validateur d'enregistrement par instance.

---

### `AutoMap`

```csharp
public void AutoMap()
```

Mappe automatiquement toutes les propriétés publiques de la classe en respectant les attributs.

## Exemples

### Mapping manuel

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Name("Full Name");
        Map(m => m.Age).Index(1).Default(0);
        Map(m => m.BirthDate).Index(2).TypeConverter<DateTimeConverter>();
    }
}
```

### Avec validateurs

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        Map(m => m.Name).Index(0).Validate<NotEmptyValidator>();
        Map(m => m.Age).Index(1).Validate(new RangeValidator(0, 120));

        Validate<PersonRecordValidator>();
    }
}
```

### AutoMap + raffinement

```csharp
public class PersonMap : ExcelClassMap<Person>
{
    public PersonMap()
    {
        AutoMap();
        Map(m => m.Secret).Ignore();
    }
}
```
