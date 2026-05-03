using System.Globalization;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using ExcelHelper.Mapping;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelConfigurationTests
{
    [Fact]
    public void Configuration_Should_Have_Sensible_Defaults()
    {
        var config = new ExcelConfiguration();

        Assert.True(config.HasHeaderRecord);
        Assert.Equal(1, config.HeaderRow);
        Assert.Equal(2, config.StartRow);
        Assert.Equal(0, config.SheetIndex);
        Assert.True(config.UseOADate);
        Assert.False(config.TrimCellValues);
        Assert.True(config.IgnoreBlankRows);
        Assert.False(config.IgnoreReferences);
        Assert.Equal(CultureInfo.InvariantCulture, config.CultureInfo);
    }

    [Fact]
    public void Configuration_Can_Override_Defaults()
    {
        var config = new ExcelConfiguration
        {
            HasHeaderRecord = false,
            HeaderRow = 3,
            StartRow = 4,
            SheetName = "Data",
            SheetIndex = 1,
            UseOADate = false,
            TrimCellValues = true,
            IgnoreBlankRows = false,
            CultureInfo = CultureInfo.GetCultureInfo("fr-FR")
        };

        Assert.False(config.HasHeaderRecord);
        Assert.Equal(3, config.HeaderRow);
        Assert.Equal(4, config.StartRow);
        Assert.Equal("Data", config.SheetName);
        Assert.Equal(1, config.SheetIndex);
        Assert.False(config.UseOADate);
        Assert.True(config.TrimCellValues);
        Assert.False(config.IgnoreBlankRows);
        Assert.Equal("fr-FR", config.CultureInfo.Name);
    }

    [Fact]
    public void Default_ReadingExceptionOccurred_Returns_False()
    {
        var config = new ExcelConfiguration();
        var args = new ReadingExceptionEventArgs(
            new ReadingContext(config),
            new ExcelHelperException("test"));

        Assert.False(config.ReadingExceptionOccurred(args));
    }

    [Fact]
    public void Default_WritingExceptionOccurred_Returns_False()
    {
        var config = new ExcelConfiguration();
        var args = new WritingExceptionEventArgs(
            new WritingContext(config),
            new ExcelHelperException("test"));

        Assert.False(config.WritingExceptionOccurred(args));
    }

    [Fact]
    public void Default_BadDataFound_Returns_False()
    {
        var config = new ExcelConfiguration();
        var args = new BadDataFoundEventArgs(
            new ReadingContext(config), "F", "x", 1, 1);

        Assert.False(config.BadDataFound(args));
    }

    [Fact]
    public void Default_ValidationFailed_Returns_False()
    {
        var config = new ExcelConfiguration();
        var args = new ValidationFailedEventArgs(
            new ReadingContext(config), "F", 1, "err");

        Assert.False(config.ValidationFailed(args));
    }

    [Fact]
    public void MissingFieldFound_Can_Be_Assigned()
    {
        var config = new ExcelConfiguration();
        MissingFieldEventArgs? captured = null;
        config.MissingFieldFound = args => captured = args;

        var ctx = new ReadingContext(config);
        config.MissingFieldFound(new MissingFieldEventArgs(ctx, "Name", 2));

        Assert.NotNull(captured);
        Assert.Equal("Name", captured.FieldName);
    }

    private class TestClass { }

    [Fact]
    public void UnregisterClassMap_Removes_Registered_Map()
    {
        var config = new ExcelConfiguration();
        config.RegisterClassMap(new ExcelClassMap<TestClass>());

        var result = config.UnregisterClassMap<TestClass>();

        Assert.True(result);
        Assert.False(config.Maps.Contains<TestClass>());
    }

    [Fact]
    public void UnregisterClassMap_Returns_False_When_Not_Registered()
    {
        var config = new ExcelConfiguration();

        var result = config.UnregisterClassMap<TestClass>();

        Assert.False(result);
    }

    [Fact]
    public void UnregisterAllClassMaps_Removes_All_Maps()
    {
        var config = new ExcelConfiguration();
        config.RegisterClassMap(new ExcelClassMap<TestClass>());
        config.RegisterClassMap<AnotherTestClassMap>();

        config.UnregisterAllClassMaps();

        Assert.False(config.Maps.Contains<TestClass>());
        Assert.False(config.Maps.Contains<AnotherTestClass>());
    }

    private class AnotherTestClass { }
    private class AnotherTestClassMap : ExcelClassMap<AnotherTestClass> { }
}