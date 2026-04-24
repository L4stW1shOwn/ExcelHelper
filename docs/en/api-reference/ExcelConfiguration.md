# API Reference — ExcelConfiguration

Namespace: `ExcelHelper.Core`

## Description

Centralizes all configuration parameters for read and write operations.

## Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `Maps` | `ExcelClassMapCollection` | `new` | Collection of registered class maps. |
| `ObjectResolver` | `IObjectResolver` | `DefaultObjectResolver` | Instance resolver for mapping. |
| `CultureInfo` | `CultureInfo` | `InvariantCulture` | Culture used for conversions. |
| `HasHeaderRecord` | `bool` | `true` | Indicates whether the sheet has a header row. |
| `HeaderRow` | `int` | `1` | 1-based header row. |
| `StartRow` | `int` | `2` | 1-based row where data starts. |
| `SheetName` | `string?` | `null` | Name of the sheet to use. |
| `SheetIndex` | `int` | `0` | 0-based sheet index. |
| `UseOADate` | `bool` | `true` | Interprets numbers as OADate. |
| `TrimCellValues` | `bool` | `false` | Trims whitespace around strings. |
| `IgnoreBlankRows` | `bool` | `true` | Ignores completely empty rows. |
| `IgnoreReferences` | `bool` | `false` | Ignores Excel error references. |
| `PrepareHeaderForMatch` | `Func<string, string>` | `Trim` | Transforms headers before comparison. |
| `ReadingExceptionOccurred` | `Func<ExcelHelperException, bool>` | `_ => false` | Hook called when a read error occurs. |
| `WritingExceptionOccurred` | `Func<ExcelHelperException, bool>` | `_ => false` | Hook called when a write error occurs. |
| `MissingFieldFound` | `Action<string[], int, ReadingContext>?` | `null` | Hook called when a required field is missing. |
| `BadDataFound` | `Func<BadDataFoundArgs, bool>` | `_ => false` | Hook called when invalid data is detected. |

## Methods

### `RegisterClassMap<T>`

```csharp
public ExcelConfiguration RegisterClassMap<T>(ExcelClassMap<T> map)
```

Registers a class map for type `T`. Returns `this` for fluent chaining.

## Examples

### Basic configuration

```csharp
var config = new ExcelConfiguration
{
    SheetName = "Data",
    TrimCellValues = true,
    CultureInfo = CultureInfo.GetCultureInfo("fr-FR")
};
```

### Fluent chaining

```csharp
var config = new ExcelConfiguration()
    .RegisterClassMap(new PersonMap())
    .RegisterClassMap(new OrderMap());
```

### Hooks

```csharp
var config = new ExcelConfiguration
{
    ReadingExceptionOccurred = ex =>
    {
        Console.WriteLine(ex.Message);
        return true;
    },
    BadDataFound = args =>
    {
        Console.WriteLine($"Bad data: {args.RawCellValue}");
        return true;
    }
};
```
