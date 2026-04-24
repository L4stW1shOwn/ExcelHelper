# API Reference — ExcelMemberMap&lt;TClass, TMember&gt;

Namespace: `ExcelHelper.Mapping`

## Description

Represents the fluent mapping of a member (property or field) to an Excel column.

## Properties

| Property | Type | Description |
|---|---|---|
| `Data` | `MemberMapData` | Raw mapping data. |

## Methods

### `Index`

```csharp
public ExcelMemberMap<TClass, TMember> Index(int index)
```

Defines the 0-based column index.

---

### `Name`

```csharp
public ExcelMemberMap<TClass, TMember> Name(string name)
```

Defines the column name (used for the header).

---

### `Default`

```csharp
public ExcelMemberMap<TClass, TMember> Default(TMember defaultValue)
```

Defines the default value if the cell is empty.

---

### `TypeConverter<TConverter>`

```csharp
public ExcelMemberMap<TClass, TMember> TypeConverter<TConverter>() where TConverter : class
```

Defines the type converter to use.

---

### `Ignore`

```csharp
public ExcelMemberMap<TClass, TMember> Ignore()
```

Marks the member as ignored.

---

### `Optional`

```csharp
public ExcelMemberMap<TClass, TMember> Optional()
```

Marks the member as optional (no error if missing).

---

### `Validate<TValidator>`

```csharp
public ExcelMemberMap<TClass, TMember> Validate<TValidator>() where TValidator : IExcelFieldValidator, new()
```

Adds a field validator by type.

---

### `Validate`

```csharp
public ExcelMemberMap<TClass, TMember> Validate(IExcelFieldValidator validator)
```

Adds a field validator by instance.

---

### `GetGetter`

```csharp
public Func<TClass, TMember> GetGetter()
```

Returns a compiled getter for this member.

---

### `GetSetter`

```csharp
public Action<TClass, TMember> GetSetter()
```

Returns a compiled setter for this member.

## Example

```csharp
Map(m => m.Price)
    .Index(3)
    .Name("Unit Price (EUR)")
    .Default(0.0m)
    .TypeConverter<DecimalConverter>()
    .Optional()
    .Validate(new GreaterThanValidator(0));
```
