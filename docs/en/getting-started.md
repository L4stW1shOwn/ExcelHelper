# Getting Started

This guide helps you install ExcelHelper and run your first example in minutes.

## Prerequisites

- .NET SDK 8.0+ (or any compatible framework)
- A .NET Console, Web, or Library project

## Installation

```bash
dotnet add package ExcelHelper
```

> **⚠️ EPPlus License**: EPPlus 7+ requires you to set a license context before any use.
>
> ```csharp
> using OfficeOpenXml;
> ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial
> ```

## First Example — Reading

Create an Excel file `people.xlsx` with a sheet containing:

| Name | Age | City |
|---|---|---|
| Alice | 30 | Paris |
| Bob | 25 | Lyon |

```csharp
using OfficeOpenXml;
using ExcelHelper;

// Set the EPPlus license (do once at startup)
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

Output:
```
Alice, 30, Paris
Bob, 25, Lyon
```

## First Example — Writing

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

The `output.xlsx` file will contain:

| Name | Age | City |
|---|---|---|
| Alice | 30 | Paris |
| Bob | 25 | Lyon |

## Next Steps

- [Reading Excel Files](guides/reading.md)
- [Writing Excel Files](guides/writing.md)
- [Mapping](guides/mapping.md)
