# ExcelHelper

An Excel (XLSX) reading and writing library for .NET, architecturally inspired by **CsvHelper** but built natively for Excel using **EPPlus 7+**.

ExcelHelper provides a familiar, fluent API for mapping Excel columns to .NET objects, with full support for type conversion, validation, async streaming, and multi-target framework compatibility.

---

## Installation

```bash
dotnet add package ExcelHelper
```

> **Note**: EPPlus 7+ requires a license context. Set it before using ExcelHelper:
> ```csharp
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or Commercial
> ```

---

## Quick Start

### Reading

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = reader.GetRecords<Person>().ToList();
```

### Writing

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

writer.WriteRecords(people);
```

### Mapping

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

### Attribute-based Mapping

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

### Async

```csharp
// .NET Core+
await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}

// .NET Framework
var people = await reader.GetRecordsAsync<Person>();

// Writing async
await writer.WriteRecordsAsync(GetProductsAsync());
```

---

## .NET Compatibility

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

## Comparison with CsvHelper

| Feature | CsvHelper | ExcelHelper |
|---|---|---|
| `ClassMap<T>` | ✅ | ✅ `ExcelClassMap<T>` |
| `MemberMap` | ✅ | ✅ `ExcelMemberMap<TClass, TMember>` |
| Auto-mapping | ✅ | ✅ `AutoMap` |
| Attribute mapping | ✅ | ✅ `[ExcelColumn]`, `[ExcelIgnore]`, etc. |
| Type converters | ✅ | ✅ `IExcelTypeConverter<T>` |
| Validation | ✅ | ✅ `IExcelFieldValidator`, `IExcelRecordValidator` |
| `IAsyncEnumerable<T>` | ✅ | ✅ (Core+) / `Task<IReadOnlyList<T>>` (Framework) |
| Hooks (`ReadingExceptionOccurred`) | ✅ | ✅ |
| Headers | ✅ | ✅ |
| Multi-sheet | ❌ | ✅ |
| Formulas / Styles | ❌ | ✅ (via EPPlus) |
| OADate conversion | ❌ | ✅ |

---

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for detailed design documentation.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## Roadmap

See [ROADMAP.md](ROADMAP.md) for the implementation phases.

## License

MIT License. See LICENSE for details.

EPPlus is licensed under PolyForm Noncommercial 1.0.0 (free for non-commercial use) or a commercial license.
