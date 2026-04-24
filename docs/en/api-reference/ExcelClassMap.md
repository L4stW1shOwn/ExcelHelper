# API Reference — ExcelClassMap&lt;T&gt;

Namespace: `ExcelHelper.Mapping`

## Description

Defines the mapping between a .NET class and Excel columns.

## Properties

| Property | Type | Description |
|---|---|---|
| `MemberMaps` | `ExcelMemberMapCollection` | Collection of member mappings. |
| `RecordValidators` | `List<object>` | Registered record validators. |

## Methods

### `Map<TMember>`

```csharp
public ExcelMemberMap<T, TMember> Map<TMember>(Expression<Func<T, TMember>> expression)
```

Declares the mapping for a member (property or field) of the class.

Returns an `ExcelMemberMap<T, TMember>` for fluent configuration.

---

### `Validate<TValidator>`

```csharp
public ExcelClassMap<T> Validate<TValidator>() where TValidator : IExcelRecordValidator<T>, new()
```

Adds a record validator by type.

---

### `Validate`

```csharp
public ExcelClassMap<T> Validate(IExcelRecordValidator<T> validator)
```

Adds a record validator by instance.

---

### `AutoMap`

```csharp
public void AutoMap()
```

Automatically maps all public properties of the class respecting attributes.

## Examples

### Manual mapping

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

### With validators

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

### AutoMap + refinement

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
