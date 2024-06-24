using System.Reflection;
using System.Runtime.CompilerServices;

namespace dSPACE.Runtime.InteropServices;

internal static partial class Extensions
{
    private static readonly ILookup<string, Assembly> _mappingForwardedTypesFrom;
#if !NETCOREAPP
    internal static readonly Assembly _mscorlib;
#endif
    private static readonly AttributeUsageAttribute _defaultAttrUsage = new(AttributeTargets.All);

    static Extensions()
    {
#if NETCOREAPP
        var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies().AsEnumerable();
        if (!loadedAssemblies.Any(a => a.GetName().Name == "mscorlib"))
        {
            // explicitly load mscorlib for compatibility with .NET Framework
            loadedAssemblies = loadedAssemblies.Prepend(AppDomain.CurrentDomain.Load("mscorlib"));
        }
        _mappingForwardedTypesFrom = (from a in loadedAssemblies
                                      from t in a.GetForwardedTypes()
                                      select (t, a))
                                      .ToLookup(ta => ta.t.FullName!, ta => ta.a);
#else
        _mscorlib = typeof(object).Assembly;
        var miGetCustomAttribute = typeof(TypeForwardedToAttribute).GetMethod(
            "GetCustomAttribute", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.DeclaredOnly)
            ?? throw new MissingMemberException(typeof(TypeForwardedToAttribute).FullName, "GetCustomAttribute");
        var delType = typeof(Func<,>).MakeGenericType(typeof(Extensions).Assembly.GetType(), typeof(TypeForwardedToAttribute[]));
        var delGetCustomAttribute = miGetCustomAttribute.CreateDelegate(delType);
        _mappingForwardedTypesFrom = (from a in AppDomain.CurrentDomain.GetAssemblies()
                                      from ta in (TypeForwardedToAttribute[])delGetCustomAttribute.DynamicInvoke(a)!
                                      select (t: ta.Destination, a))
                                      .ToLookup(ta => ta.t.FullName!, ta => ta.a);
#endif
    }

    public static bool EqualsToRuntimeType(this Type roType, Type rtType)
    {
        if (!roType.Assembly.ReflectionOnly
#if !NETCOREAPP
        // mscorlib can't be loaded into reflection-only context
        // https://learn.microsoft.com/en-us/dotnet/framework/reflection-and-codedom/how-to-load-assemblies-into-the-reflection-only-context
            && roType.Assembly != _mscorlib
#endif
        )
        {
            throw new ArgumentOutOfRangeException(nameof(roType));
        }
        if (rtType.Assembly.ReflectionOnly)
        {
            throw new ArgumentOutOfRangeException(nameof(rtType));
        }
        if (roType.FullName != rtType.FullName)
        {
            return false;
        }

        var roAssemblyName = roType.Assembly.GetName();
        var rtAssemblyNames = rtType.Assembly.GetName().Yield();
        // For our cases assume roType always comes from ref assembly, if there is one
        var forwardedFrom = rtType.GetCustomAttribute<TypeForwardedFromAttribute>(inherit: false);
        if (forwardedFrom != null)
        {
            rtAssemblyNames = rtAssemblyNames.Append(new AssemblyName(forwardedFrom.AssemblyFullName));
        }
        else if (_mappingForwardedTypesFrom[rtType.FullName!] is var forwardedFromAssemblies)
        {
            rtAssemblyNames = rtAssemblyNames.Concat(forwardedFromAssemblies.Select(a => a.GetName()));
        }

        return rtAssemblyNames.Any(an => IsEqualForReflectionOnly(an, roAssemblyName));
    }

    public static bool IsAssignableFromReflectionOnly(this Type rtTypeTo, Type? roTypeFrom)
    {
        if (roTypeFrom is null) //same as (object?)roTypeFrom == null
        {
            return false;
        }
        if (EqualsToRuntimeType(roTypeFrom, rtTypeTo))
        {
            return true;
        }

        if (IsSubclassOfRuntimeType(roTypeFrom, rtTypeTo))
        {
            return true;
        }
        if (rtTypeTo.IsInterface)
        {
            return roTypeFrom.ImplementInterface(rtTypeTo);
        }
        else if (rtTypeTo.IsGenericParameter)
        {
            var constraints = rtTypeTo.GetGenericParameterConstraints();
            for (var i = 0; i < constraints.Length; i++)
            {
                if (!constraints[i].IsAssignableFromReflectionOnly(roTypeFrom))
                {
                    return false;
                }
            }

            return true;
        }

        return false;
    }

    public static bool ImplementInterface(this Type roType, Type ifaceType)
    {
        if (!roType.Assembly.ReflectionOnly
#if !NETCOREAPP
            && roType.Assembly != _mscorlib
#endif
        )
        {
            throw new ArgumentOutOfRangeException(nameof(roType));
        }
        if (ifaceType.Assembly.ReflectionOnly || !ifaceType.IsInterface)
        {
            throw new ArgumentOutOfRangeException(nameof(ifaceType), $"nameof(ifaceType) must be an runtime interface type");
        }

        var t = roType;
        while (t != null)
        {
            var interfaces = t.GetInterfaces();
            if (interfaces != null)
            {
                for (var i = 0; i < interfaces.Length; i++)
                {
                    // Interfaces don't derive from other interfaces, they implement them.
                    // So instead of IsSubclassOf, we should use ImplementInterface instead.
                    if (EqualsToRuntimeType(interfaces[i], ifaceType) ||
                        (interfaces[i] != null && interfaces[i].ImplementInterface(ifaceType)))
                    {
                        return true;
                    }
                }
            }

            t = t.BaseType;
        }

        return false;
    }

    public static bool IsSubclassOfRuntimeType(this Type roType, Type rtType)
    {
        if (!roType.Assembly.ReflectionOnly
#if !NETCOREAPP
            && roType.Assembly != _mscorlib
#endif
        )
        {
            throw new ArgumentOutOfRangeException(nameof(roType));
        }
        if (rtType.Assembly.ReflectionOnly)
        {
            throw new ArgumentOutOfRangeException(nameof(rtType));
        }
        var p = roType;
        if (EqualsToRuntimeType(p, rtType))
        {
            return false;
        }

        while (p != null)
        {
            if (EqualsToRuntimeType(p, rtType))
            {
                return true;
            }
            p = p.BaseType;
        }
        return false;
    }

    public static bool IsEqualForReflectionOnly(this AssemblyName anRuntime, AssemblyName anReflectionOnly)
    {
        // Compare each AssemblyName component except Version when comparing assemblies
        // from command line parameters to runtime assemblies
        if (anReflectionOnly.Name != anRuntime.Name)
        {
            return false;
        }
        if (anReflectionOnly.CultureName != anRuntime.CultureName)
        {
            return false;
        }

        var roPKToken = anReflectionOnly.GetPublicKeyToken();
        var rtPKToken = anRuntime.GetPublicKeyToken();
        if (roPKToken != rtPKToken &&
            (roPKToken == null || rtPKToken == null || !roPKToken.SequenceEqual(rtPKToken)))
        {
            return false;
        }

        return true;
    }

    public static Type? ToRuntimeType(this Type roType)
    {
#if !NETCOREAPP
        if (!roType.Assembly.ReflectionOnly)
        {
            return roType;
        }
#endif
        var roAssemblyName = roType.Assembly.GetName();
        var runtimeAssemblies = AppDomain.CurrentDomain.GetAssemblies().Where(a => IsEqualForReflectionOnly(a.GetName(), roAssemblyName));
        return runtimeAssemblies.Select(a => a.GetType(roType.FullName!)).FirstOrDefault();
    }

    public static object[] GetCustomAttributes(this IList<CustomAttributeData> attrData)
    {
        List<Attribute> result = new(attrData.Count);
        foreach (var ad in attrData)
        {
            var rtAttrType = ad.AttributeType.ToRuntimeType()
                ?? throw new InvalidOperationException($"Can't create an instance of reflection-only type {ad.AttributeType}");
            var ctorArgs = ad.ConstructorArguments.Select(arg => arg.ToRuntimeTypedValue()).ToArray();
            var attrObj = (Attribute?)Activator.CreateInstance(rtAttrType, ctorArgs)
                ?? throw new InvalidOperationException($"Can't create an instance of {rtAttrType}");
            foreach (var namedArg in ad.NamedArguments)
            {
                MemberInfo? fieldOrProperty = namedArg.IsField ?
                    rtAttrType.GetField(namedArg.MemberName) : rtAttrType.GetProperty(namedArg.MemberName);
                if (fieldOrProperty is FieldInfo field)
                {
                    field.SetValue(attrObj, namedArg.TypedValue.ToRuntimeTypedValue());
                }
                else if (fieldOrProperty is PropertyInfo property)
                {
                    property.SetValue(attrObj, namedArg.TypedValue.ToRuntimeTypedValue());
                }
                else
                {
                    throw new InvalidOperationException($"Can't find public {(namedArg.IsField ? "field" : "property")} \"{namedArg.MemberName}\" in {rtAttrType}");
                }
            }
            result.Add(attrObj);
        }
        return result.ToArray();
    }

    public static object[] GetCustomAttributes(this IList<CustomAttributeData> attrData, Type attributeType)
    {
        List<Attribute> resultList = new(attrData.Count);
        foreach (var ad in attrData.Where(ad => attributeType.IsAssignableFromReflectionOnly(ad.AttributeType)))
        {
            var ctorArgs = ad.ConstructorArguments.Select(arg => arg.ToRuntimeTypedValue()).ToArray();
            var attrObj = (Attribute?)Activator.CreateInstance(attributeType, ctorArgs)
                ?? throw new InvalidOperationException($"Can't create an instance of {attributeType}");
            foreach (var namedArg in ad.NamedArguments)
            {
                MemberInfo? fieldOrProperty = namedArg.IsField ?
                    attributeType.GetField(namedArg.MemberName) : attributeType.GetProperty(namedArg.MemberName);
                if (fieldOrProperty is FieldInfo field)
                {
                    field.SetValue(attrObj, namedArg.TypedValue.ToRuntimeTypedValue());
                }
                else if (fieldOrProperty is PropertyInfo property)
                {
                    property.SetValue(attrObj, namedArg.TypedValue.ToRuntimeTypedValue());
                }
                else
                {
                    throw new InvalidOperationException($"Can't find public {(namedArg.IsField ? "field" : "property")} \"{namedArg.MemberName}\" in {attributeType}");
                }
            }
            resultList.Add(attrObj);
        }
        var result = (Attribute[])Array.CreateInstance(attributeType, resultList.Count);
        resultList.CopyTo(result);
        return result;
    }

    public static bool IsDefined(this IList<CustomAttributeData> attrData, Type attributeType)
    {
        return attrData.Any(ad => attributeType.IsAssignableFromReflectionOnly(ad.AttributeType));
    }

    public static object? ToRuntimeTypedValue(this CustomAttributeTypedArgument arg)
    {
        var rtType = arg.ArgumentType.ToRuntimeType();
        if (rtType is null || rtType == typeof(void) || arg.Value is null)
        {
            return arg.Value;
        }
        if (rtType.IsEnum)
        {
            return Enum.ToObject(rtType, arg.Value);
        }
        else if (rtType.IsAssignableFrom(arg.Value.GetType()))
        {
            return arg.Value is Type typeValue && typeValue.Assembly.ReflectionOnly &&
                typeValue is not ROTypeExtended ?
                new ROTypeExtended(new ROAssemblyExtended(typeValue.Assembly), typeValue) :
                arg.Value;
        }
        return Convert.ChangeType(arg.Value, rtType);
    }

    public static Dictionary<Type, AttributeUsageAttribute> GetAttributeTypes(IEnumerable<Attribute> attributes)
    {
        return attributes
            .Select(attr => attr.GetType()).Distinct()
            .ToDictionary(t => t, t => t.GetCustomAttribute<AttributeUsageAttribute>() ?? _defaultAttrUsage);
    }

    public static void AddAttributesToList(List<Attribute> attributeList, Attribute[] attributes, Dictionary<Type, AttributeUsageAttribute> types)
    {
        for (var i = 0; i < attributes.Length; i++)
        {
            var attrType = attributes[i].GetType();

            if (!types.TryGetValue(attrType, out var usage))
            {
                // the type has never been seen before if it's inheritable add it to the list
                usage = attrType.GetCustomAttribute<AttributeUsageAttribute>() ?? _defaultAttrUsage;
                types[attrType] = usage;

                if (usage.Inherited)
                {
                    attributeList.Add(attributes[i]);
                }
            }
            else if (usage.Inherited && usage.AllowMultiple)
            {
                // we saw this type already add it only if it is inheritable and it does allow multiple 
                attributeList.Add(attributes[i]);
            }
        }
    }

    public static MethodInfo? GetParentDefinition(this MethodInfo mi)
    {
        // from RuntimeMethodInfo.GetParentDefinition()
        if (mi.IsVirtual && mi.DeclaringType != null &&
            !mi.DeclaringType.IsInterface &&
            mi.DeclaringType.BaseType is var parent and not null)
        {
            // not exactly the same as lookup by slot #, but good enough for most cases
            // and it is only using public API.
            var all = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            return parent.GetMethod(mi.Name, all, null, mi.GetParameters().Select(pi => pi.ParameterType).ToArray(), null);
        }
        return null;
    }

    public static PropertyInfo? GetParentDefinition(this PropertyInfo property, Type[] propertyParameters)
    {
        // for the current property get the base class of the getter and the setter, they might be different
        // note that this only works for RuntimeMethodInfo
        var propAccessor = property.GetGetMethod(true) ?? property.GetSetMethod(true);

        propAccessor = propAccessor?.GetParentDefinition();

        if (propAccessor != null)
        {
            // There is a public overload of Type.GetProperty that takes both a BingingFlags enum and a return type.
            // However, we cannot use that because it doesn't accept null for "types".
            return propAccessor.DeclaringType!.GetProperty(
                property.Name,
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly,
                null, //will use default binder
                property.PropertyType,
                propertyParameters, //used for index properties
                null);
        }

        return null;
    }

    public static EventInfo? GetParentDefinition(this EventInfo ev)
    {
        var add = ev.GetAddMethod(true)?.GetParentDefinition();

        return add?.DeclaringType!.GetEvent(ev.Name);
    }

    public static ParameterInfo? GetParentDefinition(this ParameterInfo param)
    {
        // note that this only works for MethodInfo
        var method = param.Member as MethodInfo;

        if (method != null)
        {
            method = method.GetParentDefinition();

            if (method != null)
            {
                // Find the ParameterInfo on this method
                var parameters = method.GetParameters();
                return parameters[param.Position]; // Point to the correct ParameterInfo of the method
            }
        }
        return null;
    }

    public static IEnumerable<T> Yield<T>(this T source)
    {
        yield return source;
    }

}
