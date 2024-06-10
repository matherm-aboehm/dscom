using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROParameterInfoExtended : ParameterInfo
{
    private readonly MemberInfo _roMemberExtended;
    private readonly ROAssemblyExtended _roAssemblyExtended;
    private readonly ParameterInfo _roParameterInfo;
    public ROParameterInfoExtended(ROMethodInfoExtended roMemberExtended, ParameterInfo roParameterInfo)
    {
        _roMemberExtended = roMemberExtended;
        _roAssemblyExtended = roMemberExtended._roTypeExtended._roAssemblyExtended;
        _roParameterInfo = roParameterInfo;
    }
    public ROParameterInfoExtended(ROPropertyInfoExtended roMemberExtended, ParameterInfo roParameterInfo)
    {
        _roMemberExtended = roMemberExtended;
        _roAssemblyExtended = roMemberExtended._roTypeExtended._roAssemblyExtended;
        _roParameterInfo = roParameterInfo;
    }
    public ROParameterInfoExtended(ROConstructorInfoExtended roMemberExtended, ParameterInfo roParameterInfo)
    {
        _roMemberExtended = roMemberExtended;
        _roAssemblyExtended = roMemberExtended._roTypeExtended._roAssemblyExtended;
        _roParameterInfo = roParameterInfo;
    }
    public override ParameterAttributes Attributes => _roParameterInfo.Attributes;
    public override object? DefaultValue
        => new CustomAttributeTypedArgument(
            _roParameterInfo.ParameterType, RawDefaultValue).ToRuntimeTypedValue();
#if NETCOREAPP
    public override bool HasDefaultValue => _roParameterInfo.HasDefaultValue;
#else
    public override bool HasDefaultValue
    {
        get
        {
            if (!_roAssemblyExtended.ReflectionOnly)
            {
                return _roParameterInfo.HasDefaultValue;
            }
            var rawdefault = RawDefaultValue;
            if (rawdefault == DBNull.Value || (IsOptional && rawdefault == Type.Missing))
            {
                return false;
            }
            return true;
        }
    }

#endif
    public override MemberInfo Member => _roMemberExtended;
    public override int MetadataToken => _roParameterInfo.MetadataToken;
    public override string? Name => _roParameterInfo.Name;
    public override Type ParameterType
        => new ROTypeExtended(_roAssemblyExtended, _roParameterInfo.ParameterType);
    public override int Position => _roParameterInfo.Position;
    public override object? RawDefaultValue => _roParameterInfo.RawDefaultValue;

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roParameterInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes();
        if (inherit)
        {
            List<Attribute> resultList = new((Attribute[])result);
            var attrTypes = Extensions.GetAttributeTypes(resultList);
            for (var baseParam = _roParameterInfo.GetParentDefinition();
                baseParam != null; baseParam = baseParam.GetParentDefinition())
            {
                result = baseParam.GetCustomAttributesData().GetCustomAttributes();
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
        for (var baseParam = _roParameterInfo.GetParentDefinition();
            baseParam != null; baseParam = baseParam.GetParentDefinition())
        {
            result = baseParam.GetCustomAttributesData().GetCustomAttributes(attributeType);
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
        for (var baseParam = _roParameterInfo.GetParentDefinition();
            baseParam != null; baseParam = baseParam.GetParentDefinition())
        {
            if (baseParam.GetCustomAttributesData().IsDefined(attributeType))
            {
                return true;
            }
        }
        return false;
    }

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROParameterInfoExtended extended ?
            _roParameterInfo.Equals(extended._roParameterInfo) : _roParameterInfo.Equals(obj));
    public override int GetHashCode() => _roParameterInfo.GetHashCode();
    public override string ToString() => _roParameterInfo.ToString();
    #endregion
}
