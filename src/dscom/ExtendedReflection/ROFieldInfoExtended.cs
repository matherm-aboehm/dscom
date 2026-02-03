using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROFieldInfoExtended : FieldInfo
{
    private readonly FieldInfo _roFieldInfo;
    private readonly ROTypeExtended _roTypeExtended;
    private ROTypeExtended? _roFieldTypeExtended;

    public ROFieldInfoExtended(ROTypeExtended roTypeExtended, FieldInfo roFieldInfo)
    {
        _roTypeExtended = roTypeExtended;
        _roFieldInfo = roFieldInfo;
    }

    public override FieldAttributes Attributes => _roFieldInfo.Attributes;

    public override RuntimeFieldHandle FieldHandle => _roFieldInfo.FieldHandle;

    public override Type FieldType => _roFieldTypeExtended ??=
        new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roFieldInfo.FieldType);

    public override Type? DeclaringType => _roFieldInfo.DeclaringType is null ? null :
        (_roFieldInfo.DeclaringType == _roTypeExtended._roType ? _roTypeExtended :
            new ROTypeExtended(_roTypeExtended._roAssemblyExtended, _roFieldInfo.DeclaringType));
    public override int MetadataToken => _roFieldInfo.MetadataToken;
    public override Module Module => _roTypeExtended.Module;
    public override string Name => _roFieldInfo.Name;
    public override Type? ReflectedType => _roTypeExtended;

    private MarshalAsAttribute? ComputeMarshalAsAttribute()
    {
        if ((_roFieldInfo.Attributes & FieldAttributes.HasFieldMarshal) == 0)
        {
            return null;
        }
        else if (_roTypeExtended.Module is ROModuleExtended roModuleEx)
        {
            return roModuleEx.ComputeMarshalAsAttribute(MetadataToken);
        }

        throw new NotSupportedException($"{nameof(_roTypeExtended.Module)} must be from reflection-only load context.");
    }

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roFieldInfo.GetCustomAttributesData();

    public override object[] GetCustomAttributes(bool inherit)
        => GetCustomAttributesData().GetCustomAttributes(
            (Attributes & FieldAttributes.HasFieldMarshal) != 0 ?
            ComputeMarshalAsAttribute : null);

    public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        => GetCustomAttributesData().GetCustomAttributes(attributeType, ComputeMarshalAsAttribute);

    public override bool IsDefined(Type attributeType, bool inherit)
        => (attributeType.Equals(typeof(MarshalAsAttribute)) &&
            (Attributes & FieldAttributes.HasFieldMarshal) != 0) ||
            GetCustomAttributesData().IsDefined(attributeType);

    public override object? GetValue(object? obj)
        => _roFieldInfo.GetValue(obj);

    public override void SetValue(object? obj, object? value, BindingFlags invokeAttr, Binder? binder, CultureInfo? culture)
        => _roFieldInfo.SetValue(obj, value, invokeAttr, binder, culture);

    public override object? GetRawConstantValue()
        => _roFieldInfo.GetRawConstantValue();

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROFieldInfoExtended extended ?
            _roFieldInfo.Equals(extended._roFieldInfo) : _roFieldInfo.Equals(obj));
    public override int GetHashCode() => _roFieldInfo.GetHashCode();
    public override string? ToString() => _roFieldInfo.ToString();
    #endregion
}
