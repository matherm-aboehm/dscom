using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

//https://stackoverflow.com/questions/10679515/winrt-projected-types-documentation
[ExcludeFromCodeCoverage]
internal static class WinRTExtensions
{
    private enum WinMDTypeKind
    {
        Attribute,
        Enum,
        Delegate,
        Interface,
        PDelegate,
        PInterface,
        Struct,
        Runtimeclass,
    };

    private const string Mscorlib = "mscorlib";
    private const string System = "System";
    private const string SystemRuntimeWindowsRuntime = "System.Runtime.WindowsRuntime";
    private const string SystemRuntimeWindowsRuntimeUIXaml = "System.Runtime.WindowsRuntime.UI.Xaml";
    private const string SystemNumericsVectors = "System.Numerics.Vectors";
    private const string InternalUri = "Internal.Uri";


    // this is hardcoded into runtime using macros with winrtprojectedtypes.h
    // and has no public API, so hardcode this here too.
    private static readonly Dictionary<(string ns, string name), (string ns, string name, string assembly, WinMDTypeKind kind)> _redirectedTypeNames =
    new()
    {
        { ("Windows.Foundation.Metadata", "AttributeUsageAttribute"),("System", "AttributeUsageAttribute", Mscorlib, WinMDTypeKind.Attribute) },
        { ("Windows.Foundation.Metadata", "AttributeTargets"),("System", "AttributeTargets", Mscorlib, WinMDTypeKind.Enum) },
        { ("Windows.UI", "Color"),("Windows.UI", "Color", SystemRuntimeWindowsRuntime, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "DateTime"),("System", "DateTimeOffset", Mscorlib, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "EventHandler`1"),("System", "EventHandler`1", Mscorlib, WinMDTypeKind.PDelegate) },
        { ("Windows.Foundation", "EventRegistrationToken"),("System.Runtime.InteropServices.WindowsRuntime", "EventRegistrationToken", Mscorlib, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "HResult"),("System", "Exception", Mscorlib, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "IReference`1"),("System", "Nullable`1", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation", "Point"),("Windows.Foundation", "Point", SystemRuntimeWindowsRuntime, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "Rect"),("Windows.Foundation", "Rect", SystemRuntimeWindowsRuntime, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "Size"),("Windows.Foundation", "Size", SystemRuntimeWindowsRuntime, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "TimeSpan"),("System", "TimeSpan", Mscorlib, WinMDTypeKind.Struct) },
        { ("Windows.Foundation", "Uri"),("System", "Uri", InternalUri, WinMDTypeKind.Runtimeclass) },
        { ("Windows.Foundation", "IClosable"),("System", "IDisposable", Mscorlib, WinMDTypeKind.Interface) },
        { ("Windows.Foundation.Collections", "IIterable`1"),("System.Collections.Generic", "IEnumerable`1", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation.Collections", "IVector`1"),("System.Collections.Generic", "IList`1", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation.Collections", "IVectorView`1"),("System.Collections.Generic", "IReadOnlyList`1", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation.Collections", "IMap`2"),("System.Collections.Generic", "IDictionary`2", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation.Collections", "IMapView`2"),("System.Collections.Generic", "IReadOnlyDictionary`2", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.Foundation.Collections", "IKeyValuePair`2"),("System.Collections.Generic", "KeyValuePair`2", Mscorlib, WinMDTypeKind.PInterface) },
        { ("Windows.UI.Xaml.Input", "ICommand"),("System.Windows.Input", "ICommand", System, WinMDTypeKind.Interface) },
        { ("Windows.UI.Xaml.Interop", "IBindableIterable"),("System.Collections", "IEnumerable", Mscorlib, WinMDTypeKind.Interface) },
        { ("Windows.UI.Xaml.Interop", "IBindableVector"),("System.Collections", "IList", Mscorlib, WinMDTypeKind.Interface) },
        { ("Windows.UI.Xaml.Interop", "INotifyCollectionChanged"),("System.Collections.Specialized", "INotifyCollectionChanged", System, WinMDTypeKind.Interface) },
        { ("Windows.UI.Xaml.Interop", "NotifyCollectionChangedEventHandler"),("System.Collections.Specialized", "NotifyCollectionChangedEventHandler", System, WinMDTypeKind.Delegate) },
        { ("Windows.UI.Xaml.Interop", "NotifyCollectionChangedEventArgs"),("System.Collections.Specialized", "NotifyCollectionChangedEventArgs", System, WinMDTypeKind.Runtimeclass) },
        { ("Windows.UI.Xaml.Interop", "NotifyCollectionChangedAction"),("System.Collections.Specialized", "NotifyCollectionChangedAction", System, WinMDTypeKind.Enum) },
        { ("Windows.UI.Xaml.Data", "INotifyPropertyChanged"),("System.ComponentModel", "INotifyPropertyChanged", System, WinMDTypeKind.Interface) },
        { ("Windows.UI.Xaml.Data", "PropertyChangedEventHandler"),("System.ComponentModel", "PropertyChangedEventHandler", System, WinMDTypeKind.Delegate) },
        { ("Windows.UI.Xaml.Data", "PropertyChangedEventArgs"),("System.ComponentModel", "PropertyChangedEventArgs", System, WinMDTypeKind.Runtimeclass) },
        { ("Windows.UI.Xaml", "CornerRadius"),("Windows.UI.Xaml", "CornerRadius", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml", "Duration"),("Windows.UI.Xaml", "Duration", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml", "DurationType"),("Windows.UI.Xaml", "DurationType", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Enum) },
        { ("Windows.UI.Xaml", "GridLength"),("Windows.UI.Xaml", "GridLength", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml", "GridUnitType"),("Windows.UI.Xaml", "GridUnitType", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Enum) },
        { ("Windows.UI.Xaml", "Thickness"),("Windows.UI.Xaml", "Thickness", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Interop", "TypeName"),("System", "Type", Mscorlib, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Controls.Primitives", "GeneratorPosition"),("Windows.UI.Xaml.Controls.Primitives", "GeneratorPosition", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Media", "Matrix"),("Windows.UI.Xaml.Media", "Matrix", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Media.Animation", "KeyTime"),("Windows.UI.Xaml.Media.Animation", "KeyTime", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Media.Animation", "RepeatBehavior"),("Windows.UI.Xaml.Media.Animation", "RepeatBehavior", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.UI.Xaml.Media.Animation", "RepeatBehaviorType"),("Windows.UI.Xaml.Media.Animation", "RepeatBehaviorType", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Enum) },
        { ("Windows.UI.Xaml.Media.Media3D", "Matrix3D"),("Windows.UI.Xaml.Media.Media3D", "Matrix3D", SystemRuntimeWindowsRuntimeUIXaml, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Vector2"),("System.Numerics", "Vector2", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Vector3"),("System.Numerics", "Vector3", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Vector4"),("System.Numerics", "Vector4", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Matrix3x2"),("System.Numerics", "Matrix3x2", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Matrix4x4"),("System.Numerics", "Matrix4x4", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Plane"),("System.Numerics", "Plane", SystemNumericsVectors, WinMDTypeKind.Struct) },
        { ("Windows.Foundation.Numerics", "Quaternion"),("System.Numerics", "Quaternion", SystemNumericsVectors, WinMDTypeKind.Struct) },
    };

    private static readonly Dictionary<string, Assembly?> _assemblies = new();
    private static readonly Dictionary<(string, string), Type?> _redirectedResolvedTypes =
        _redirectedTypeNames.ToDictionary(kvp => kvp.Key, kvp => GetRuntimeAssemblyForSimpleName(kvp.Value.assembly)?.GetType($"{kvp.Value.ns}.{kvp.Value.name}", false));

    private static Assembly? GetRuntimeAssemblyForSimpleName(string simpleName)
    {
        if (_assemblies.TryGetValue(simpleName, out var assembly))
        {
            return assembly;
        }
        try
        {
#if NETCOREAPP
            assembly = AppDomain.CurrentDomain.Load(simpleName);
#else
            assembly = Assembly.LoadWithPartialName(simpleName);
#endif
        }
        catch (FileNotFoundException)
        {
            assembly = null;
        }
        catch (FileLoadException ex) when (ex.HResult == HRESULT.E_POINTER)
        {
            assembly = null;
        }
        _assemblies.Add(simpleName, assembly);
        return assembly;
    }

    internal static bool ResolveRedirectedInterface(Type typeWinRT, out Type? clrType)
    {
        if (ResolveRedirectedType(typeWinRT, out clrType, out var kind))
        {
            if (kind is WinMDTypeKind.Interface or WinMDTypeKind.PInterface &&
                (clrType == null ||
                // filter out KeyValuePair and Nullable which are structures projected from WinRT interfaces
                (clrType != typeof(KeyValuePair<,>) &&
                clrType != typeof(Nullable<>))))
            {
                return true;
            }
        }
        return false;
    }

    internal static bool ResolveRedirectedDelegate(Type typeWinRT, out Type? clrType)
    {
        if (ResolveRedirectedType(typeWinRT, out clrType, out var kind))
        {
            if (kind is WinMDTypeKind.Delegate or WinMDTypeKind.PDelegate)
            {
                return true;
            }
        }

        return false;
    }

    private static bool ResolveRedirectedType(Type typeWinRT, out Type? clrType, out WinMDTypeKind? kind)
    {
        var key = (typeWinRT.Namespace ?? string.Empty, typeWinRT.Name);
        clrType = null;
        kind = null;
        if (!_redirectedTypeNames.TryGetValue(key, out var mapping))
        {
            return false;
        }
        kind = mapping.kind;
        return _redirectedResolvedTypes.TryGetValue(key, out clrType);
    }

    private static readonly Func<Type, bool>? _getIsWindowsRuntimeObject =
        typeof(Type).GetProperty("IsWindowsRuntimeObject",
            BindingFlags.Instance | BindingFlags.NonPublic)?.GetGetMethod(true)?
        .CreateDelegate<Func<Type, bool>>(null);

    internal static bool IsWinRTObjectType(this Type type)
    {
        try
        {
            if (_getIsWindowsRuntimeObject != null && !type.Assembly.ReflectionOnly)
            {
                if (type is ROTypeExtended extended)
                {
                    type = extended._roType;
                }
                return _getIsWindowsRuntimeObject(type);
            }
        }
        catch (NotImplementedException) { }

        // Try to determine if this object represents a WindowsRuntime object - i.e. is either
        // ProjectedFromWinRT or derived from a class that is

        if (!type.IsCOMObject)
        {
            return false;
        }

        var parentType = type;
        do
        {
            if (parentType.IsProjectedFromWinRT())
            {
                // Found a WinRT COM object
                return true;
            }
            if (parentType.IsImport)
            {
                // Found a class that is actually imported from COM but not WinRT
                // this is definitely a non-WinRT COM object
                return false;
            }
            parentType = parentType.BaseType;
        } while (parentType != null);

        return false;
    }

    private static readonly Func<Type, bool>? _getIsExportedToWindowsRuntime =
        typeof(Type).GetProperty("IsExportedToWindowsRuntime",
            BindingFlags.Instance | BindingFlags.NonPublic)?.GetGetMethod(true)?
        .CreateDelegate<Func<Type, bool>>(null);

    internal static bool IsExportedToWinRT(this Type type)
    {
        try
        {
            if (type is ROTypeExtended extended)
            {
                type = extended._roType;
            }
            return _getIsExportedToWindowsRuntime?.Invoke(type) ?? false;
        }
        catch (NotImplementedException) { return false; }
    }

    internal static bool IsProjectedFromWinRT(this Type type)
    {
        // checks for winmd specific internals are not tested
        return (type.Attributes & TypeAttributes.WindowsRuntime) != 0 ||
            type.FullName == "System.Runtime.InteropServices.WindowsRuntime.RuntimeClass";
        //return type.Assembly.GetName().ContentType == AssemblyContentType.WindowsRuntime;
    }

    internal enum InteropKind
    {
        ManagedToNative, // use for RCW-related queries
        NativeToManaged, // use for CCW-related queries
    };

    internal static bool IsWinRTRedirectedInterface(this Type type, InteropKind interopKind)
    {
        if (!type.IsInterface)
        {
            return false;
        }
        // checks for winmd specific internals are not tested
        if (ResolveRedirectedInterface(type, out _))
        {
            return true;
        }
        else if (interopKind == InteropKind.ManagedToNative)
        {
            var genType = !type.IsGenericType ? null :
                (type.IsGenericTypeDefinition ? type : type.GetGenericTypeDefinition());

            if (type.Equals(typeof(ICollection)) ||
                (genType != null && (genType.Equals(typeof(ICollection<>)) ||
                genType.Equals(typeof(IReadOnlyCollection<>)))))
            {
                return true;
            }
        }
        return false;
    }

    internal static bool IsWinRTRedirectedDelegate(this Type type)
    {
        if (!type.IsDelegate())
        {
            return false;
        }
        // checks for winmd specific internals are not tested
        return ResolveRedirectedDelegate(type, out _);
    }

    [Flags]
    internal enum Mode
    {
        Projected = 0x1,
        Redirected = 0x2,
        All = Projected | Redirected
    };
    internal static bool SupportsGenericInterop(this Type type, InteropKind interopKind,
                        Mode mode = Mode.All)
    {
        return (type.IsInterface || type.IsDelegate()) &&    // interface or delegate
                type.IsGenericType &&                 // generic
                !type.IsSharedByGenericInstantiations() && // unshared
                !type.ContainsGenericParameters &&        // closed over concrete types
                                                          // defined in .winmd or one of the redirected mscorlib interfaces
                ((((mode & Mode.Projected) != 0) && type.IsProjectedFromWinRT()) ||
                (((mode & Mode.Redirected) != 0) && (type.IsWinRTRedirectedInterface(interopKind) || type.IsWinRTRedirectedDelegate())));
    }
}
