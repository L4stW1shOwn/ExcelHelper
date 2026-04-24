# Contributing to ExcelHelper

Thank you for your interest in contributing to ExcelHelper!

## Getting Started

1. Fork the repository
2. Clone your fork
3. Create a feature branch (`git checkout -b feature/my-feature`)
4. Build the solution: `dotnet build ExcelHelper.sln`
5. Run tests: `dotnet test ExcelHelper.sln`

## Requirements

- All code must compile on **all target frameworks** (`net10.0;net9.0;net8.0;net48;net47;netstandard2.1;netstandard2.0`)
- **Zero warnings**: `TreatWarningsAsErrors` is enabled
- **XML documentation** required for all public APIs
- **Unit tests** required for all new features
- Follow the existing code style

## Coding Standards

1. Use `Nullable` annotations (`?`) and guard clauses
2. Use `CompiledExpressionCache` for property get/set in hot paths
3. Minimize allocations in loops
4. Support fluent API (method chaining)
5. Use `IAsyncEnumerable<T>` for .NET Core+, `Task<IReadOnlyList<T>>` for .NET Framework

## Pull Request Process

1. Every PR must target a single phase or feature
2. Build must pass on all target frameworks (`dotnet build`)
3. All tests must pass (`dotnet test`)
4. XML docs must be complete
5. Update `ARCHITECTURE.md` if adding new extension points
6. At least one reviewer approval required

## Reporting Issues

Please include:
- Target framework
- ExcelHelper version
- Minimal reproduction code
- Expected vs actual behavior
