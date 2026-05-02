using Xunit;
using ExcelHelper.Core;

namespace ExcelHelper.Tests;

public class ExcelContextTests
{
    [Fact]
    public void ReadingContext_Has_New_Properties()
    {
        var config = new ExcelConfiguration();
        var ctx = new ReadingContext(config);

        ctx.RawRecord = new[] { "a", "b" };
        ctx.CurrentFieldValue = "test";
        ctx.CurrentFieldName = "Field1";
        ctx.CurrentFieldIndex = 2;
        ctx.CurrentRecord = new object();

        Assert.Equal(new[] { "a", "b" }, ctx.RawRecord);
        Assert.Equal("test", ctx.CurrentFieldValue);
        Assert.Equal("Field1", ctx.CurrentFieldName);
        Assert.Equal(2, ctx.CurrentFieldIndex);
        Assert.NotNull(ctx.CurrentRecord);
    }

    [Fact]
    public void ReadingContext_RawRecord_Replaces_Old_Record_Property()
    {
        var config = new ExcelConfiguration();
        var ctx = new ReadingContext(config);

        // Ensure the old 'Record' property no longer exists
        var type = ctx.GetType();
        Assert.Null(type.GetProperty("Record"));
        Assert.NotNull(type.GetProperty("RawRecord"));
    }

    [Fact]
    public void WritingContext_Has_New_Properties()
    {
        var config = new ExcelConfiguration();
        var ctx = new WritingContext(config);

        ctx.CurrentFieldValue = 42;
        ctx.CurrentFieldName = "Age";
        ctx.CurrentFieldIndex = 1;
        ctx.CurrentRecord = new object();

        Assert.Equal(42, ctx.CurrentFieldValue);
        Assert.Equal("Age", ctx.CurrentFieldName);
        Assert.Equal(1, ctx.CurrentFieldIndex);
        Assert.NotNull(ctx.CurrentRecord);
    }
}
