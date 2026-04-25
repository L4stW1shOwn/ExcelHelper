using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelHelper.Internal;
using ExcelHelper.Validation;

namespace ExcelHelper.Mapping;

/// <summary>
///     Provides the base class for mapping a class to Excel columns.
/// </summary>
public abstract class ExcelClassMap
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelClassMap" /> class.
    /// </summary>
    /// <param name="classType">The type of the class this map is for.</param>
    protected ExcelClassMap(Type classType)
    {
        ClassType = classType;
    }

    /// <summary>
    ///     The type of the class this map is for.
    /// </summary>
    public virtual Type ClassType { get; }

    /// <summary>
    ///     Gets the collection of member maps.
    /// </summary>
    public virtual ExcelMemberMapCollection MemberMaps { get; } = [];

    /// <summary>
    ///     Gets the collection of record validators.
    /// </summary>
    public virtual List<object> RecordValidators { get; } = [];
}

/// <summary>
///     Maps a class to Excel columns.
/// </summary>
/// <typeparam name="T">The type to map.</typeparam>
public class ExcelClassMap<T> : ExcelClassMap
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelClassMap{T}" /> class.
    /// </summary>
    public ExcelClassMap()
        : base(typeof(T))
    {
    }

    /// <summary>
    ///     Automatically maps all public properties of the class.
    /// </summary>
    public void AutoMap()
    {
        AutoMapper.Apply(this);
    }

    /// <summary>
    ///     Maps the specified member to an Excel column.
    /// </summary>
    /// <typeparam name="TMember">The member type.</typeparam>
    /// <param name="expression">The member expression.</param>
    /// <returns>The <see cref="ExcelMemberMap{T, TMember}" /> for configuration.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="expression" /> is null.</exception>
    public ExcelMemberMap<T, TMember> Map<TMember>(Expression<Func<T, TMember>> expression)
    {
        if (expression == null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        var member = ReflectionHelper.GetMember(expression);
        var map = new ExcelMemberMap<T, TMember>(member)
        {
            Data =
            {
                Index = GetMaxIndex() + 1
            }
        };
        MemberMaps.Add(map.Data);
        return map;
    }

    /// <summary>
    ///     Adds a record validator for this class map.
    /// </summary>
    /// <typeparam name="TValidator">The type of the record validator.</typeparam>
    /// <returns>This <see cref="ExcelClassMap{T}" /> instance.</returns>
    public ExcelClassMap<T> Validate<TValidator>()
        where TValidator : IExcelRecordValidator<T>, new()
    {
        RecordValidators.Add(new TValidator());
        return this;
    }

    /// <summary>
    ///     Adds a record validator for this class map.
    /// </summary>
    /// <param name="validator">The record validator instance.</param>
    /// <returns>This <see cref="ExcelClassMap{T}" /> instance.</returns>
    public ExcelClassMap<T> Validate(IExcelRecordValidator<T> validator)
    {
        if (validator == null)
        {
            throw new ArgumentNullException(nameof(validator));
        }

        RecordValidators.Add(validator);
        return this;
    }

    /// <summary>
    ///     Get the largest index for the members and references.
    /// </summary>
    /// <returns>The max index.</returns>
    private int GetMaxIndex()
    {
        return MemberMaps.Count == 0 ? 0 : MemberMaps.Max(m => m.Index);
    }
}