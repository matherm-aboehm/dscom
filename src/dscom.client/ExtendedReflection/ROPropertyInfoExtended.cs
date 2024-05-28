using System.Globalization;
using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROPropertyInfoExtended : PropertyInfo
{
    private readonly PropertyInfo _roPropertyInfo;
    internal readonly ROTypeExtended _roTypeExtended;
    private ROTypeExtended? _roPropTypeExtended;
    public ROPropertyInfoExtended(ROTypeExtended roTypeExtended, PropertyInfo roPropertyInfo)
    {
        _roTypeExtended = roTypeExtended;
        _roPropertyInfo = roPropertyInfo;
    }

    public override PropertyAttributes Attributes => _roPropertyInfo.Attributes;
    public override bool CanRead => _roPropertyInfo.CanRead;
    public override bool CanWrite => _roPropertyInfo.CanWrite;

    public override Type PropertyType => _roPropTypeExtended ??=
        new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roPropertyInfo.PropertyType);

    public override Type? DeclaringType => _roPropertyInfo.DeclaringType is null ? null :
        (_roPropertyInfo.DeclaringType == _roTypeExtended._roType ? _roTypeExtended :
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roPropertyInfo.DeclaringType));

    public override int MetadataToken => _roPropertyInfo.MetadataToken;
    public override Module Module => _roTypeExtended.Module;
    public override string Name => _roPropertyInfo.Name;
    public override Type? ReflectedType => _roTypeExtended;

    public override MethodInfo[] GetAccessors(bool nonPublic)
        => _roPropertyInfo.GetAccessors(nonPublic)
        .Select(mi => new ROMethodInfoExtended(_roTypeExtended, mi)).ToArray();

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roPropertyInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes();
        if (inherit)
        {
            List<Attribute> resultList = new((Attribute[])result);
            var indexParamTypes = _roPropertyInfo.GetIndexParameters().Select(pi => pi.ParameterType).ToArray();
            var attrTypes = Extensions.GetAttributeTypes(resultList);
            for (var baseProp = _roPropertyInfo.GetParentDefinition(indexParamTypes);
                baseProp != null; baseProp = baseProp.GetParentDefinition(indexParamTypes))
            {
                result = baseProp.GetCustomAttributesData().GetCustomAttributes();
                Extensions.AddAttributesToList(resultList, (Attribute[])result, attrTypes);
            }
            result = resultList.ToArray();
        }
        return result;
    }

    public override object[] GetCustomAttributes(Type attributeType, bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes(attributeType);
        if (!inherit ||
            (attributeType.GetCustomAttribute<AttributeUsageAttribute>() is var attrUsage &&
            (attrUsage is { Inherited: false } ||
            (attrUsage is null or { Inherited: true, AllowMultiple: false } && result.Length != 0))))
        {
            return result;
        }

        List<Attribute> resultList = new((Attribute[])result);
        var indexParamTypes = _roPropertyInfo.GetIndexParameters().Select(pi => pi.ParameterType).ToArray();
        for (var baseProp = _roPropertyInfo.GetParentDefinition(indexParamTypes);
            baseProp != null; baseProp = baseProp.GetParentDefinition(indexParamTypes))
        {
            result = baseProp.GetCustomAttributesData().GetCustomAttributes(attributeType);
            if (!(attrUsage?.AllowMultiple ?? false) && result.Length != 0)
            {
                return result;
            }
            resultList.AddRange((Attribute[])result);
        }
        result = (Attribute[])Array.CreateInstance(attributeType, resultList.Count);
        resultList.CopyTo((Attribute[])result);
        return result;
    }

    public override bool IsDefined(Type attributeType, bool inherit)
    {
        var result = GetCustomAttributesData().IsDefined(attributeType);
        if (result || !inherit ||
            attributeType.GetCustomAttribute<AttributeUsageAttribute>() is { Inherited: false })
        {
            return result;
        }
        var indexParamTypes = _roPropertyInfo.GetIndexParameters().Select(pi => pi.ParameterType).ToArray();
        for (var baseProp = _roPropertyInfo.GetParentDefinition(indexParamTypes);
            baseProp != null; baseProp = baseProp.GetParentDefinition(indexParamTypes))
        {
            if (baseProp.GetCustomAttributesData().IsDefined(attributeType))
            {
                return true;
            }
        }
        return false;
    }

    public override MethodInfo? GetGetMethod(bool nonPublic)
    {
        var getmethod = _roPropertyInfo.GetGetMethod(nonPublic);
        return getmethod is null ? null : new ROMethodInfoExtended(_roTypeExtended, getmethod);
    }

    public override ParameterInfo[] GetIndexParameters()
        => _roPropertyInfo.GetIndexParameters()
            .Select(pi => new ROParameterInfoExtended(this, pi)).ToArray();

    public override MethodInfo? GetSetMethod(bool nonPublic)
    {
        var setmethod = _roPropertyInfo.GetSetMethod(nonPublic);
        return setmethod is null ? null : new ROMethodInfoExtended(_roTypeExtended, setmethod);
    }

    public override object? GetValue(object? obj, BindingFlags invokeAttr, Binder? binder, object?[]? index, CultureInfo? culture)
        => _roPropertyInfo.GetValue(obj, invokeAttr, binder, index, culture);

    public override void SetValue(object? obj, object? value, BindingFlags invokeAttr, Binder? binder, object?[]? index, CultureInfo? culture)
        => _roPropertyInfo.SetValue(obj, value, invokeAttr, binder, index, culture);

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROPropertyInfoExtended extended ?
            _roPropertyInfo.Equals(extended._roPropertyInfo) : _roPropertyInfo.Equals(obj));
    public override int GetHashCode() => _roPropertyInfo.GetHashCode();
    public override string? ToString() => _roPropertyInfo.ToString();
    #endregion
}
