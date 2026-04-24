# Guide — Async

ExcelHelper provides an async API that automatically adapts to the target framework.

## .NET Core / .NET 5+ / `netstandard2.1`

On these platforms, `GetRecordsAsync<T>()` returns an `IAsyncEnumerable<T>`:

### Reading

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

await foreach (var person in reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}
```

With cancellation:

```csharp
await foreach (var person in reader.GetRecordsAsync<Person>(cancellationToken))
{
    // ...
}
```

### Writing

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

On these platforms, `GetRecordsAsync<T>()` returns a `Task<IReadOnlyList<T>>`:

### Reading

```csharp
using var stream = File.OpenRead("data.xlsx");
using var reader = new ExcelReader(stream);

var people = await reader.GetRecordsAsync<Person>();

foreach (var person in people)
{
    Console.WriteLine(person.Name);
}
```

### Writing

```csharp
using var stream = File.OpenWrite("output.xlsx");
using var writer = new ExcelWriter(stream);

await writer.WriteRecordsAsync(products);
```

## Compatibility Table

| TFM | `IAsyncEnumerable<T>` | `Task<IReadOnlyList<T>>` |
|---|---|---|
| `net10.0` | ✅ | ✅ |
| `net9.0` | ✅ | ✅ |
| `net8.0` | ✅ | ✅ |
| `netstandard2.1` | ✅ | ✅ |
| `netstandard2.0` | ❌ | ✅ |
| `net48` | ❌ | ✅ |
| `net47` | ❌ | ✅ |

## Important Points

- The `IAsyncEnumerable` implementation uses `Task.Yield()` to avoid blocking the thread.
- The .NET Framework implementation simply delegates to `Task.Run`.
- The API is **identical** in consumer code; the compiler resolves the appropriate overload.

## Full Cross-Platform Example

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

// This code compiles on ALL supported targets
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

> In practice, you do not need the `#if`; use `await foreach` directly on .NET Core+ or `await ...` on .NET Framework. The compiler selects the correct signature automatically.
