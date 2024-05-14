using System.Globalization;
using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROConstructorInfoExtended : ConstructorInfo
{
    private readonly ConstructorInfo _roConstructorInfo;
    internal readonly ROTypeExtended _roTypeExtended;
    public ROConstructorInfoExtended(ROTypeExtended roTypeExtended, ConstructorInfo roConstructorInfo)
    {
        _roTypeExtended = roTypeExtended;
        _roConstructorInfo = roConstructorInfo;
    }

    public override MethodAttributes Attributes => _roConstructorInfo.Attributes;
    public override RuntimeMethodHandle MethodHandle => _roConstructorInfo.MethodHandle;
    public override int MetadataToken => _roConstructorInfo.MetadataToken;
    public override Module Module => _roTypeExtended.Module;

    public override Type? DeclaringType => _roConstructorInfo.DeclaringType is null ? null :
        (_roConstructorInfo.DeclaringType == _roTypeExtended._roType ? _roTypeExtended :
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roConstructorInfo.DeclaringType));

    public override string Name => _roConstructorInfo.Name;
    public override Type? ReflectedType => _roTypeExtended;

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roConstructorInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
        => GetCustomAttributesData().GetCustomAttributes();

    public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        => GetCustomAttributesData().GetCustomAttributes(attributeType);

    public override bool IsDefined(Type attributeType, bool inherit)
        => GetCustomAttributesData().IsDefined(attributeType);

    public override MethodImplAttributes GetMethodImplementationFlags()
        => _roConstructorInfo.MethodImplementationFlags;

    public override ParameterInfo[] GetParameters()
        => _roConstructorInfo.GetParameters()
            .Select(pi => new ROParameterInfoExtended(this, pi)).ToArray();

    public override object Invoke(BindingFlags invokeAttr, Binder? binder, object?[]? parameters, CultureInfo? culture)
        => _roConstructorInfo.Invoke(invokeAttr, binder, parameters, culture);

    public override object? Invoke(object? obj, BindingFlags invokeAttr, Binder? binder, object?[]? parameters, CultureInfo? culture)
        => _roConstructorInfo.Invoke(obj, invokeAttr, binder, parameters, culture);

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROConstructorInfoExtended extended ?
            _roConstructorInfo.Equals(extended._roConstructorInfo) : _roConstructorInfo.Equals(obj));
    public override int GetHashCode() => _roConstructorInfo.GetHashCode();
    public override string? ToString() => _roConstructorInfo.ToString();
    #endregion
}
