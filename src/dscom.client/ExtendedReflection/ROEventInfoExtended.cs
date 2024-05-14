using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROEventInfoExtended : EventInfo
{
    private readonly EventInfo _roEventInfo;
    private readonly ROTypeExtended _roTypeExtended;

    public ROEventInfoExtended(ROTypeExtended roTypeExtended, EventInfo roEventInfo)
    {
        _roTypeExtended = roTypeExtended;
        _roEventInfo = roEventInfo;
    }

    public override EventAttributes Attributes => _roEventInfo.Attributes;

    public override Type? DeclaringType => _roEventInfo.DeclaringType is null ? null :
        (_roEventInfo.DeclaringType == _roTypeExtended._roType ? _roTypeExtended :
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roEventInfo.DeclaringType));
    public override int MetadataToken => _roEventInfo.MetadataToken;
    public override Module Module => _roTypeExtended.Module;
    public override string Name => _roEventInfo.Name;
    public override Type? ReflectedType => _roTypeExtended;

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roEventInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes();
        if (inherit)
        {
            List<Attribute> resultList = new((Attribute[])result);
            var attrTypes = Extensions.GetAttributeTypes(resultList);
            for (var baseEvent = _roEventInfo.GetParentDefinition();
                baseEvent != null; baseEvent = baseEvent.GetParentDefinition())
            {
                result = baseEvent.GetCustomAttributesData().GetCustomAttributes();
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
        for (var baseEvent = _roEventInfo.GetParentDefinition();
            baseEvent != null; baseEvent = baseEvent.GetParentDefinition())
        {
            result = baseEvent.GetCustomAttributesData().GetCustomAttributes(attributeType);
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
        for (var baseEvent = _roEventInfo.GetParentDefinition();
            baseEvent != null; baseEvent = baseEvent.GetParentDefinition())
        {
            if (baseEvent.GetCustomAttributesData().IsDefined(attributeType))
            {
                return true;
            }
        }
        return false;
    }

    public override MethodInfo? GetAddMethod(bool nonPublic)
    {
        var add = _roEventInfo.GetAddMethod(nonPublic);
        return add is null ? null : new ROMethodInfoExtended(_roTypeExtended, add);
    }

    public override MethodInfo? GetRaiseMethod(bool nonPublic)
    {
        var raise = _roEventInfo.GetRaiseMethod(nonPublic);
        return raise is null ? null : new ROMethodInfoExtended(_roTypeExtended, raise);
    }

    public override MethodInfo? GetRemoveMethod(bool nonPublic)
    {
        var remove = _roEventInfo.GetRemoveMethod(nonPublic);
        return remove is null ? null : new ROMethodInfoExtended(_roTypeExtended, remove);
    }

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROEventInfoExtended extended ?
            _roEventInfo.Equals(extended._roEventInfo) : _roEventInfo.Equals(obj));
    public override int GetHashCode() => _roEventInfo.GetHashCode();
    public override string? ToString() => _roEventInfo.ToString();
    #endregion
}
