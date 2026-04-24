# Guide — Async

ExcelHelper fournit une API asynchrone qui s'adapte automatiquement au framework cible.

## .NET Core / .NET 5+ / `netstandard2.1`

Sur ces plateformes, `GetRecordsAsync<T>()` retourne un `IAsyncEnumerable<T>` :

### Lecture

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}
```

Avec annulation :

```csharp
await foreach (var person in reader.GetRecordsAsync<Person>(cancellationToken))
{
    // ...
}
```

### Écriture

```csharp
async IAsyncEnumerable<Product> GenerateProductsAsync()
{
    for (int i = 0; i < 1000; i++)
    {
        yield return new Product { Name = $"Product {i}", Price = i * 1.5m };
        await Task.Delay(10);
    }
}

using var stream = File.OpenWrite("products.xlsx");
using var writer = new ExcelWriter(stream);

await writer.WriteRecordsAsync(GenerateProductsAsync());
```

## .NET Framework (`net48`, `net47`) / `netstandard2.0`

Sur ces plateformes, `GetRecordsAsync<T>()` retourne un `Task<IReadOnlyList<T>>` :

### Lecture

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = await reader.GetRecordsAsync<Person>();

foreach (var person in people)
{
    Console.WriteLine(person.Name);
}
```

### Écriture

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

await writer.WriteRecordsAsync(products);
```

## Tableau de compatibilité

| TFM | `IAsyncEnumerable<T>` | `Task<IReadOnlyList<T>>` |
|---|---|---|
| `net10.0` | ✅ | ✅ |
| `net9.0` | ✅ | ✅ |
| `net8.0` | ✅ | ✅ |
| `netstandard2.1` | ✅ | ✅ |
| `netstandard2.0` | ❌ | ✅ |
| `net48` | ❌ | ✅ |
| `net47` | ❌ | ✅ |

## Points importants

- L'implémentation `IAsyncEnumerable` utilise `Task.Yield()` pour ne pas bloquer le thread.
- L'implémentation .NET Framework délègue simplement à `Task.Run`.
- L'API est **identique** dans le code consommateur ; le compilateur résout la surcharge appropriée.

## Exemple complet cross-platform

```csharp
using OfficeOpenXml;
using ExcelHelper;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

public class LogEntry
{
    public DateTime Timestamp { get; set; }
    public string Level { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
}

// Ce code compile sur TOUTES les cibles supportées
public async Task ProcessLogsAsync(string path)
{
    using var stream = File.OpenRead(path);
    using var reader = new ExcelReader(stream);

#if NETSTANDARD2_1 || NETCOREAPP
    await foreach (var entry in reader.GetRecordsAsync<LogEntry>())
    {
        Console.WriteLine($"[{entry.Timestamp}] {entry.Level}: {entry.Message}");
    }
#else
    var entries = await reader.GetRecordsAsync<LogEntry>();
    foreach (var entry in entries)
    {
        Console.WriteLine($"[{entry.Timestamp}] {entry.Level}: {entry.Message}");
    }
#endif
}
```

> En pratique, vous n'avez pas besoin du `#if` ; utilisez directement `await foreach` sur .NET Core+ ou `await ...` sur .NET Framework. Le compilateur sélectionne la bonne signature automatiquement.
