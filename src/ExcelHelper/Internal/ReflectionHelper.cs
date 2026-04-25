using System;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelHelper.Internal;

/// <summary>
///     Provides helper methods for reflection operations.
/// </summary>
public static class ReflectionHelper
{
    /// <summary>
    ///     Gets the <see cref="MemberInfo" /> from a member expression.
    /// </summary>
    /// <typeparam name="T">The type containing the member.</typeparam>
    /// <typeparam name="TMember">The member type.</typeparam>
    /// <param name="expression">The member expression.</param>
    /// <returns>The <see cref="MemberInfo" />.</returns>
    /// <exception cref="ArgumentException">Thrown when the expression is not a member expression.</exception>
    public static MemberInfo GetMember<T, TMember>(Expression<Func<T, TMember>> expression)
    {
        if (expression == null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        if (expression.Body is MemberExpression memberExpression)
        {
            return memberExpression.Member;
        }

        throw new ArgumentException("Expression must be a member expression.", nameof(expression));
    }

    /// <summary>
    ///     Determines whether the specified type is nullable.
    /// </summary>
    /// <param name="type">The type to check.</param>
    /// <returns><c>true</c> if the type is nullable; otherwise, <c>false</c>.</returns>
    public static bool IsNullable(Type type)
    {
        if (type == null)
        {
            throw new ArgumentNullException(nameof(type));
        }

        return Nullable.GetUnderlyingType(type) != null;
    }

    /// <summary>
    ///     Gets the underlying type of a nullable type, or the type itself if not nullable.
    /// </summary>
    /// <param name="type">The type.</param>
    /// <returns>The underlying type.</returns>
    public static Type GetUnderlyingType(Type type)
    {
        if (type == null)
        {
            throw new ArgumentNullException(nameof(type));
        }

        var underlying = Nullable.GetUnderlyingType(type);
        return underlying ?? type;
    }

    /// <summary>
    ///     Determines whether the specified type is a simple type (primitive, string, decimal, DateTime, Guid, enum, or
    ///     nullable thereof).
    /// </summary>
    /// <param name="type">The type to check.</param>
    /// <returns><c>true</c> if the type is simple; otherwise, <c>false</c>.</returns>
    public static bool IsSimpleType(Type type)
    {
        if (type == null)
        {
            throw new ArgumentNullException(nameof(type));
        }

        var underlying = GetUnderlyingType(type);

        return underlying.IsPrimitive
               || underlying == typeof(string)
               || underlying == typeof(decimal)
               || underlying == typeof(DateTime)
               || underlying == typeof(Guid)
               || underlying == typeof(TimeSpan)
               || underlying.IsEnum;
    }
}