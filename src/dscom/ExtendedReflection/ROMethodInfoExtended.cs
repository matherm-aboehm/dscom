using System.Globalization;
using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROMethodInfoExtended : MethodInfo
{
    private readonly MethodInfo _roMethodInfo;
    internal readonly ROTypeExtended _roTypeExtended;
    private ROParameterInfoExtended? _roReturnParamExtended;
    public ROMethodInfoExtended(ROTypeExtended roTypeExtended, MethodInfo roMethodInfo)
    {
        _roTypeExtended = roTypeExtended;
        _roMethodInfo = roMethodInfo;
    }

    public override MethodAttributes Attributes => _roMethodInfo.Attributes;
    public override RuntimeMethodHandle MethodHandle => _roMethodInfo.MethodHandle;
    public override MemberTypes MemberType => _roMethodInfo.MemberType;
    public override int MetadataToken => _roMethodInfo.MetadataToken;
    public override Module Module => _roTypeExtended.Module;

    public override Type? DeclaringType => _roMethodInfo.DeclaringType is null ? null :
        (_roMethodInfo.DeclaringType == _roTypeExtended._roType ? _roTypeExtended :
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roMethodInfo.DeclaringType));

    public override string Name => _roMethodInfo.Name;
    public override Type? ReflectedType => _roTypeExtended;
    public override ICustomAttributeProvider ReturnTypeCustomAttributes => ReturnParameter;
    public override Type ReturnType => ReturnParameter.ParameterType;
    public override ParameterInfo ReturnParameter => _roReturnParamExtended ??=
        new ROParameterInfoExtended(this, _roMethodInfo.ReturnParameter);

    public override bool IsGenericMethod => _roMethodInfo.IsGenericMethod;
    public override bool IsGenericMethodDefinition => _roMethodInfo.IsGenericMethodDefinition;
    public override bool ContainsGenericParameters => _roMethodInfo.ContainsGenericParameters;

    public override Type[] GetGenericArguments()
        => _roMethodInfo.GetGenericArguments()
            .Select(t => new ROTypeExtended(_roTypeExtended._roAssemblyExtended, t)).ToArray();

    public override MethodInfo GetGenericMethodDefinition()
        => new ROMethodInfoExtended(_roTypeExtended, _roMethodInfo.GetGenericMethodDefinition());

    public override MethodInfo MakeGenericMethod(params Type[] typeArguments)
    {
        var roTypeArgs = typeArguments.Select(
            t => t is ROTypeExtended extended ? extended._roType : t).ToArray();
        return new ROMethodInfoExtended(_roTypeExtended, _roMethodInfo.MakeGenericMethod(roTypeArgs));
    }

    public override MethodInfo GetBaseDefinition()
    {
        var mibase = _roMethodInfo.GetBaseDefinition();
        return mibase == _roMethodInfo ? this : new ROMethodInfoExtended(
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, mibase.DeclaringType!),
            mibase);
    }

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roMethodInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes();
        if (inherit)
        {
            List<Attribute> resultList = new((Attribute[])result);
            var attrTypes = Extensions.GetAttributeTypes(resultList);
            for (var baseMethod = _roMethodInfo.GetParentDefinition();
                baseMethod != null; baseMethod = baseMethod.GetParentDefinition())
            {
                result = baseMethod.GetCustomAttributesData().GetCustomAttributes();
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
        for (var baseMethod = _roMethodInfo.GetParentDefinition();
            baseMethod != null; baseMethod = baseMethod.GetParentDefinition())
        {
            result = baseMethod.GetCustomAttributesData().GetCustomAttributes(attributeType);
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
        for (var baseMethod = _roMethodInfo.GetParentDefinition();
            baseMethod != null; baseMethod = baseMethod.GetParentDefinition())
        {
            if (baseMethod.GetCustomAttributesData().IsDefined(attributeType))
            {
                return true;
            }
        }
        return false;
    }

    public override MethodImplAttributes GetMethodImplementationFlags()
        => _roMethodInfo.MethodImplementationFlags;

    public override ParameterInfo[] GetParameters()
        => _roMethodInfo.GetParameters()
            .Select(pi => new ROParameterInfoExtended(this, pi)).ToArray();

    public override object? Invoke(object? obj, BindingFlags invokeAttr, Binder? binder, object?[]? parameters, CultureInfo? culture)
        => _roMethodInfo.Invoke(obj, invokeAttr, binder, parameters, culture);

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROMethodInfoExtended extended ?
            _roMethodInfo.Equals(extended._roMethodInfo) : _roMethodInfo.Equals(obj));
    public override int GetHashCode() => _roMethodInfo.GetHashCode();
    public override string? ToString() => _roMethodInfo.ToString();
    #endregion
}
