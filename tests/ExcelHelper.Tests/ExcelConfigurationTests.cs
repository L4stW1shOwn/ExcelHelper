using System.Globalization;
using ExcelHelper.Core;
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

}