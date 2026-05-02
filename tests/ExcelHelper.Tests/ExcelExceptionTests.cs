using System;
using ExcelHelper.Core;
using ExcelHelper.Exceptions;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelExceptionTests
{
    [Fact]
    public void ExcelHelperException_Has_Context_Property()
    {
        var ex = new ExcelHelperException("msg");
        Assert.Null(ex.Context);
    }

    [Fact]
    public void ExcelTypeConversionException_Accepts_Context()
    {
        var config = new ExcelConfiguration();
        var ctx = new ReadingContext(config);
        var ex = new ExcelTypeConversionException("msg", "abc", typeof(int), "F1", 1, 2, ctx);

        Assert.Same(ctx, ex.Context);
        Assert.Equal("F1", ex.FieldName);
        Assert.Equal(1, ex.Row);
    }

    [Fact]
    public void ExcelValidationException_Accepts_Context()
    {
        var config = new ExcelConfiguration();
        var ctx = new WritingContext(config);
        var ex = new ExcelValidationException("msg", "F1", 3, 4, ctx);

        Assert.Same(ctx, ex.Context);
        Assert.Equal(3, ex.Row);
    }

    [Fact]
    public void ExcelMappingException_Accepts_Context()
    {
        var config = new ExcelConfiguration();
        var ctx = new ReadingContext(config);
        var ex = new ExcelMappingException("msg", typeof(string), "Prop", ctx);

        Assert.Same(ctx, ex.Context);
    }
}
