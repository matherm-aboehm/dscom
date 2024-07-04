#if NET5_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
#endif
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROTypeExtended : Type
{
    internal readonly ROAssemblyExtended _roAssemblyExtended;
    internal readonly Type _roType;
    private ROTypeExtended? _roBaseType;
    private ROTypeExtended? _roDeclaringType;
    private MethodBase? _roDeclaringMethod;
    private ROModuleExtended? _roModuleExtended;

    public ROTypeExtended(ROAssemblyExtended roAssemblyExtended, Type roType)
    {
        if (!roType.Assembly.ReflectionOnly
#if !NETCOREAPP
            && roType.Assembly != Extensions._mscorlib
#endif
        )
        {
            throw new ArgumentOutOfRangeException(nameof(roType), $"{nameof(roType)} must be from reflection-only load context.");
        }

        _roAssemblyExtended = roAssemblyExtended._roAssembly == roType.Assembly ?
            roAssemblyExtended : new ROAssemblyExtended(roType.Assembly);
        _roType = roType;
    }

    public override Assembly Assembly => _roAssemblyExtended;
    public override string? AssemblyQualifiedName => _roType.AssemblyQualifiedName;
    public override Type? BaseType => _roType.BaseType is null ? null :
        (_roBaseType ??= new ROTypeExtended(_roAssemblyExtended, _roType.BaseType));
    public override Type? DeclaringType => _roType.DeclaringType is null ? null :
        (_roDeclaringType ??= new ROTypeExtended(_roAssemblyExtended, _roType.DeclaringType));
    public override MethodBase? DeclaringMethod => _roType.DeclaringMethod is null ? null :
        (_roDeclaringMethod ??= _roType.DeclaringMethod switch
        {
            MethodInfo mi => new ROMethodInfoExtended(this, mi),
            ConstructorInfo ci => new ROConstructorInfoExtended(this, ci),
            _ => throw new InvalidOperationException($"No extension type for {_roType.DeclaringMethod.GetType()}")
        });
    public override string? FullName => _roType.FullName;
    public override Guid GUID => _roType.GUID;
    public override int MetadataToken => _roType.MetadataToken;
    public override Module Module => _roModuleExtended ??= new ROModuleExtended(_roAssemblyExtended, _roType.Module);
    public override string? Namespace => _roType.Namespace;
    public override Type UnderlyingSystemType => _roType.UnderlyingSystemType == _roType ?
        this : new ROTypeExtended(_roAssemblyExtended, _roType.UnderlyingSystemType);
    public override string Name => _roType.Name;

    protected override TypeAttributes GetAttributeFlagsImpl() => _roType.Attributes;
    protected override bool HasElementTypeImpl() => _roType.HasElementType;
    protected override bool IsArrayImpl() => _roType.IsArray;
    protected override bool IsByRefImpl() => _roType.IsByRef;
    protected override bool IsCOMObjectImpl()
    {
        // the following is not usable, because it is overriden to always return false:
        //return _roType.IsCOMObject;
        // so try to get same Type from runtime reflection context
        var rtType = _roType.ToRuntimeType();
        if (rtType != null)
        {
            return rtType.IsCOMObject;
        }
        // fallback to attribute/name checking if runtime type isn't available
        var type = _roType;
        while (type != null && !type.IsImport &&
            type.FullName != "System.__ComObject" &&
            type.FullName != "System.Runtime.InteropServices.WindowsRuntime.RuntimeClass")
        {
            type = type.BaseType;
        }
        return type != null;
    }
    public override bool IsConstructedGenericType => _roType.IsConstructedGenericType;
    public override bool IsGenericParameter => _roType.IsGenericParameter;
    public override bool IsGenericType => _roType.IsGenericType;
    public override bool IsGenericTypeDefinition => _roType.IsGenericTypeDefinition;
    protected override bool IsPointerImpl() => _roType.IsPointer;
    protected override bool IsPrimitiveImpl() => _roType.IsPrimitive;
    public override bool IsEnum => _roType.IsEnum;
    protected override bool IsValueTypeImpl() => _roType.IsValueType;
    public override StructLayoutAttribute? StructLayoutAttribute => _roType.StructLayoutAttribute;

    public override object[] GetCustomAttributes(bool inherit)
    {
        var result = GetCustomAttributesData().GetCustomAttributes();
        if (inherit)
        {
            List<Attribute> resultList = new((Attribute[])result);
            var attrTypes = Extensions.GetAttributeTypes(resultList);
            for (var baseType = BaseType; baseType != null; baseType = baseType.BaseType)
            {
                result = baseType.GetCustomAttributesData().GetCustomAttributes();
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
        for (var baseType = BaseType; baseType != null; baseType = baseType.BaseType)
        {
            result = baseType.GetCustomAttributesData().GetCustomAttributes(attributeType);
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
        for (var baseType = BaseType; baseType != null; baseType = baseType.BaseType)
        {
            if (baseType.GetCustomAttributesData().IsDefined(attributeType))
            {
                return true;
            }
        }
        return false;
    }

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roType.GetCustomAttributesData();

    public override bool ContainsGenericParameters => _roType.ContainsGenericParameters;
    public override GenericParameterAttributes GenericParameterAttributes
        => _roType.GenericParameterAttributes;
    public override int GenericParameterPosition => _roType.GenericParameterPosition;

    public override Type GetGenericTypeDefinition()
        => new ROTypeExtended(_roAssemblyExtended, _roType.GetGenericTypeDefinition());

    public override Type[] GetGenericArguments()
        => _roType.GetGenericArguments()
            .Select(t => new ROTypeExtended(_roAssemblyExtended, t)).ToArray();

    public override Type[] GetGenericParameterConstraints()
        => _roType.GetGenericParameterConstraints()
            .Select(t => new ROTypeExtended(_roAssemblyExtended, t)).ToArray();

    public override Type MakeGenericType(params Type[] typeArguments)
    {
        var roTypeArgs = typeArguments.Select(
            t => t is ROTypeExtended extended ? extended._roType : t).ToArray();
        return new ROTypeExtended(_roAssemblyExtended, _roType.MakeGenericType(roTypeArgs));
    }

    public override Type? GetElementType()
    {
        var elementType = _roType.GetElementType();
        return elementType is null ? null : new ROTypeExtended(_roAssemblyExtended, elementType);
    }
    public override ConstructorInfo[] GetConstructors(BindingFlags bindingAttr)
        => _roType.GetConstructors(bindingAttr)
            .Select(ci => new ROConstructorInfoExtended(this, ci)).ToArray();

    public override EventInfo? GetEvent(string name, BindingFlags bindingAttr)
    {
        var ei = _roType.GetEvent(name, bindingAttr);
        return ei is null ? null : new ROEventInfoExtended(this, ei);
    }

    public override EventInfo[] GetEvents(BindingFlags bindingAttr)
        => _roType.GetEvents(bindingAttr)
            .Select(ei => new ROEventInfoExtended(this, ei)).ToArray();

    public override FieldInfo? GetField(string name, BindingFlags bindingAttr)
    {
        var fi = _roType.GetField(name, bindingAttr);
        return fi is null ? null : new ROFieldInfoExtended(this, fi);
    }

    public override FieldInfo[] GetFields(BindingFlags bindingAttr)
        => _roType.GetFields(bindingAttr)
            .Select(fi => new ROFieldInfoExtended(this, fi)).ToArray();

#if NET5_0_OR_GREATER
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.Interfaces)]
#endif
    public override Type? GetInterface(string name, bool ignoreCase)
    {
        var intf = _roType.GetInterface(name, ignoreCase);
        return intf is null ? null : new ROTypeExtended(_roAssemblyExtended, intf);
    }

    public override Type[] GetInterfaces()
        => _roType.GetInterfaces()
            .Select(t => new ROTypeExtended(_roAssemblyExtended, t)).ToArray();

    public override MemberInfo[] GetMembers(BindingFlags bindingAttr)
        => _roType.GetMembers(bindingAttr)
            .Select(m => (MemberInfo)(m switch
            {
                ConstructorInfo ci => new ROConstructorInfoExtended(this, ci),
                MethodInfo mi => new ROMethodInfoExtended(this, mi),
                FieldInfo fi => new ROFieldInfoExtended(this, fi),
                PropertyInfo pi => new ROPropertyInfoExtended(this, pi),
                EventInfo ei => new ROEventInfoExtended(this, ei),
                _ => throw new NotImplementedException()
            })).ToArray();

    public override MethodInfo[] GetMethods(BindingFlags bindingAttr)
        => _roType.GetMethods(bindingAttr)
            .Select(mi => new ROMethodInfoExtended(this, mi)).ToArray();

    public override Type? GetNestedType(string name, BindingFlags bindingAttr)
    {
        var nestedType = _roType.GetNestedType(name, bindingAttr);
        return nestedType is null ? null : new ROTypeExtended(_roAssemblyExtended, nestedType);
    }

    public override Type[] GetNestedTypes(BindingFlags bindingAttr)
        => _roType.GetNestedTypes(bindingAttr)
            .Select(t => new ROTypeExtended(_roAssemblyExtended, t)).ToArray();

    public override PropertyInfo[] GetProperties(BindingFlags bindingAttr)
        => _roType.GetProperties(bindingAttr)
            .Select(pi => new ROPropertyInfoExtended(this, pi)).ToArray();

    public override object? InvokeMember(string name, BindingFlags invokeAttr, Binder? binder, object? target, object?[]? args, ParameterModifier[]? modifiers, CultureInfo? culture, string[]? namedParameters)
        => _roType.InvokeMember(name, invokeAttr, binder, target, args, modifiers, culture, namedParameters);

    protected override ConstructorInfo? GetConstructorImpl(BindingFlags bindingAttr, Binder? binder, CallingConventions callConvention, Type[] types, ParameterModifier[]? modifiers)
    {
        var ctor = _roType.GetConstructor(bindingAttr, binder, callConvention, types, modifiers);
        return ctor is null ? null : new ROConstructorInfoExtended(this, ctor);
    }

    protected override MethodInfo? GetMethodImpl(string name, BindingFlags bindingAttr, Binder? binder, CallingConventions callConvention, Type[]? types, ParameterModifier[]? modifiers)
    {
        var mi = _roType.GetMethod(name, bindingAttr, binder, callConvention, types ?? EmptyTypes, modifiers);
        return mi is null ? null : new ROMethodInfoExtended(this, mi);
    }

    protected override PropertyInfo? GetPropertyImpl(string name, BindingFlags bindingAttr, Binder? binder, Type? returnType, Type[]? types, ParameterModifier[]? modifiers)
    {
        var pi = _roType.GetProperty(name, bindingAttr, binder, returnType, types ?? EmptyTypes, modifiers);
        return pi is null ? null : new ROPropertyInfoExtended(this, pi);
    }

    #region Object overrides
    public override bool Equals(Type? o)
        => o is not null && (o is ROTypeExtended extended ?
            _roType.Equals(extended._roType) :
            (o.Assembly.ReflectionOnly ? _roType.Equals(o) : _roType.EqualsToRuntimeType(o)));

    public override int GetHashCode() => _roType.GetHashCode();
    public override string ToString() => _roType.ToString();
    #endregion
}
