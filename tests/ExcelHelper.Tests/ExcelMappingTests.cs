using System;
using ExcelHelper.Core;
using ExcelHelper.Internal;
using ExcelHelper.Mapping;
using ExcelHelper.Tests.Models;
using Xunit;

namespace ExcelHelper.Tests;

public class ExcelMappingTests
{
    [Fact]
    public void ExcelClassMap_Should_Create_MemberMap()
    {
        var map = new ExcelClassMap<SimplePerson>();
        map.Map(m => m.Name).Index(0);
        map.Map(m => m.Age).Index(1);

        Assert.Equal(2, map.MemberMaps.Count);

        var nameMap = map.MemberMaps.FindByIndex(0);
        Assert.NotNull(nameMap);
        Assert.Equal("Name", nameMap!.Name);
        Assert.Equal(0, nameMap.Index);

        var ageMap = map.MemberMaps.FindByIndex(1);
        Assert.NotNull(ageMap);
        Assert.Equal("Age", ageMap!.Name);
        Assert.Equal(1, ageMap.Index);
    }

    [Fact]
    public void ExcelMemberMap_Should_Support_Fluent_Configuration()
    {
        var map = new ExcelClassMap<SimplePerson>();
        map.Map(m => m.Name)
            .Index(2)
            .Name("PersonName")
            .Default("Unknown")
            .Optional();

        var memberMap = map.MemberMaps.FindByName("PersonName");
        Assert.NotNull(memberMap);
        Assert.Equal(2, memberMap!.Index);
        Assert.Equal("PersonName", memberMap.Name);
        Assert.Equal("Unknown", memberMap.Default);
        Assert.True(memberMap.IsOptional);
    }

    [Fact]
    public void ExcelMemberMap_Should_Support_Ignore()
    {
        var map = new ExcelClassMap<SimplePerson>();
        map.Map(m => m.Name).Index(0);
        map.Map(m => m.Age).Ignore();

        Assert.Single(map.MemberMaps);

        var nameMap = map.MemberMaps.FindByIndex(0);
        Assert.NotNull(nameMap);
        Assert.Equal("Name", nameMap!.Name);
    }

    [Fact]
    public void AutoMap_Should_Map_All_Properties()
    {
        var map = new ExcelClassMap<Person>();
        map.AutoMap();

        Assert.Equal(3, map.MemberMaps.Count);

        var nameMap = map.MemberMaps.FindByName("Name");
        Assert.NotNull(nameMap);
        Assert.Equal(0, nameMap!.Index);

        var ageMap = map.MemberMaps.FindByName("Age");
        Assert.NotNull(ageMap);
        Assert.Equal(1, ageMap!.Index);

        var birthMap = map.MemberMaps.FindByName("BirthDate");
        Assert.NotNull(birthMap);
        Assert.Equal(2, birthMap!.Index);
    }

    [Fact]
    public void AutoMap_Should_Respect_ExcelIgnoreAttribute()
    {
        var map = new ExcelClassMap<PersonWithAttributes>();
        map.AutoMap();

        Assert.Equal(3, map.MemberMaps.Count);
        Assert.Null(map.MemberMaps.FindByName("Secret"));
    }

    [Fact]
    public void AutoMap_Should_Respect_ExcelNameAttribute()
    {
        var map = new ExcelClassMap<PersonWithAttributes>();
        map.AutoMap();

        var nameMap = map.MemberMaps.FindByName("Full Name");
        Assert.NotNull(nameMap);
        Assert.Equal("Full Name", nameMap!.Name);
    }

    [Fact]
    public void AutoMap_Should_Respect_ExcelIndexAttribute()
    {
        var map = new ExcelClassMap<PersonWithAttributes>();
        map.AutoMap();

        var ageMap = map.MemberMaps.FindByIndex(5);
        Assert.NotNull(ageMap);
        Assert.Equal("Age", ageMap!.Member.Name);
    }

    [Fact]
    public void AutoMap_Should_Respect_ExcelDefaultAttribute()
    {
        var map = new ExcelClassMap<PersonWithAttributes>();
        map.AutoMap();

        var defaultMap = map.MemberMaps.FindByName("DefaultValue");
        Assert.NotNull(defaultMap);
        Assert.Equal(42, defaultMap!.Default);
    }

    [Fact]
    public void ExcelClassMapCollection_Should_Register_And_Retrieve_Maps()
    {
        var collection = new ExcelClassMapCollection();
        var map = new ExcelClassMap<SimplePerson>();
        map.Map(m => m.Name).Index(0);

        collection.Add(map);

        Assert.True(collection.Contains<SimplePerson>());
        var retrieved = collection.Get<SimplePerson>();
        Assert.NotNull(retrieved);
        Assert.Equal(1, retrieved!.MemberMaps.Count);
    }

    [Fact]
    public void ExcelContext_Should_Register_ClassMap_Fluent()
    {
        var config = new ExcelConfiguration();
        var map = new ExcelClassMap<SimplePerson>();
        map.Map(m => m.Name).Index(0);

        var context = new ReadingContext(config);
        context.RegisterClassMap(map);

        Assert.True(config.Maps.Contains<SimplePerson>());
    }

    [Fact]
    public void CompiledExpressionCache_Should_Get_And_Set_Property()
    {
        var person = new SimplePerson { Name = "Alice", Age = 30 };

        var getter = CompiledExpressionCache.GetGetter<SimplePerson, string>("Name");
        Assert.Equal("Alice", getter(person));

        var setter = CompiledExpressionCache.GetSetter<SimplePerson, string>("Name");
        setter(person, "Bob");
        Assert.Equal("Bob", person.Name);
    }

    [Fact]
    public void ReflectionHelper_IsNullable_Should_Detect_Nullable_Types()
    {
        Assert.True(ReflectionHelper.IsNullable(typeof(int?)));
        Assert.False(ReflectionHelper.IsNullable(typeof(int)));
        Assert.True(ReflectionHelper.IsNullable(typeof(DateTime?)));
        Assert.False(ReflectionHelper.IsNullable(typeof(string)));
    }

    [Fact]
    public void ReflectionHelper_IsSimpleType_Should_Detect_Simple_Types()
    {
        Assert.True(ReflectionHelper.IsSimpleType(typeof(int)));
        Assert.True(ReflectionHelper.IsSimpleType(typeof(string)));
        Assert.True(ReflectionHelper.IsSimpleType(typeof(DateTime)));
        Assert.True(ReflectionHelper.IsSimpleType(typeof(decimal)));
        Assert.True(ReflectionHelper.IsSimpleType(typeof(Guid)));
        Assert.True(ReflectionHelper.IsSimpleType(typeof(int?)));
        Assert.False(ReflectionHelper.IsSimpleType(typeof(SimplePerson)));
    }

    [Fact]
    public void DefaultObjectResolver_Should_Create_Instance()
    {
        var resolver = new DefaultObjectResolver();
        var instance = resolver.Resolve(typeof(SimplePerson));

        Assert.NotNull(instance);
        Assert.IsType<SimplePerson>(instance);
    }

    [Fact]
    public void ExcelClassMapCollection_Should_Remove_Map()
    {
        var collection = new ExcelClassMapCollection();
        var map = new ExcelClassMap<SimplePerson>();

        collection.Add(map);
        Assert.True(collection.Contains<SimplePerson>());

        collection.Remove<SimplePerson>();
        Assert.False(collection.Contains<SimplePerson>());
    }
}