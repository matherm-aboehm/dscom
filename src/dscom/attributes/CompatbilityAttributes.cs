#if !NETSTANDARD2_0
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

[assembly: TypeForwardedTo(typeof(TypeLibImportClassAttribute))]
[assembly: TypeForwardedTo(typeof(ImportedFromTypeLibAttribute))]
[assembly: TypeForwardedTo(typeof(TypeLibTypeFlags))]
[assembly: TypeForwardedTo(typeof(TypeLibFuncFlags))]
[assembly: TypeForwardedTo(typeof(TypeLibVarFlags))]
[assembly: TypeForwardedTo(typeof(TypeLibTypeAttribute))]
[assembly: TypeForwardedTo(typeof(TypeLibFuncAttribute))]
[assembly: TypeForwardedTo(typeof(TypeLibVarAttribute))]
[assembly: TypeForwardedTo(typeof(AutomationProxyAttribute))]
[assembly: TypeForwardedTo(typeof(TypeLibVersionAttribute))]
[assembly: TypeForwardedTo(typeof(ManagedToNativeComInteropStubAttribute))]

#else

namespace System.Runtime.InteropServices;

/// <summary>Specifies which <see cref="T:System.Type" /> exclusively uses an interface. This class cannot be inherited.</summary>
[AttributeUsage(AttributeTargets.Interface, Inherited = false)]
[ComVisible(true)]
public sealed class TypeLibImportClassAttribute : Attribute
{
    internal string _importClassName;
    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.TypeLibImportClassAttribute" /> class specifying the <see cref="T:System.Type" /> that exclusively uses an interface.</summary>
    /// <param name="importClass">The <see cref="T:System.Type" /> object that exclusively uses an interface.</param>
    public TypeLibImportClassAttribute(Type importClass)
    {
        _importClassName = importClass.ToString();
    }
    /// <summary>Gets the name of a <see cref="T:System.Type" /> object that exclusively uses an interface.</summary>
    /// <returns>The name of a <see cref="T:System.Type" /> object that exclusively uses an interface.</returns>
    public string Value { get { return _importClassName; } }
}

/// <summary>Indicates that the types defined within an assembly were originally defined in a type library.</summary>
[AttributeUsage(AttributeTargets.Assembly, Inherited = false)]
[ComVisible(true)]
public sealed class ImportedFromTypeLibAttribute : Attribute
{
    internal string _val;
    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.ImportedFromTypeLibAttribute" /> class with the name of the original type library file.</summary>
    /// <param name="tlbFile">The location of the original type library file.</param>
    public ImportedFromTypeLibAttribute(string tlbFile)
    {
        _val = tlbFile;
    }
    /// <summary>Gets the name of the original type library file.</summary>
    /// <returns>The name of the original type library file.</returns>
    public string Value { get { return _val; } }
}

/// <summary>Describes the original settings of the <see cref="T:System.Runtime.InteropServices.TYPEFLAGS" /> in the COM type library from which the type was imported.</summary>
[Serializable]
[Flags()]
[ComVisible(true)]
public enum TypeLibTypeFlags
{
    /// <summary>A type description that describes an <see langword="Application" /> object.</summary>
    FAppObject = 0x0001,
    /// <summary>Instances of the type can be created by <see langword="ITypeInfo::CreateInstance" />.</summary>
    FCanCreate = 0x0002,
    /// <summary>The type is licensed.</summary>
    FLicensed = 0x0004,
    /// <summary>The type is predefined. The client application should automatically create a single instance of the object that has this attribute. The name of the variable that points to the object is the same as the class name of the object.</summary>
    FPreDeclId = 0x0008,
    /// <summary>The type should not be displayed to browsers.</summary>
    FHidden = 0x0010,
    /// <summary>The type is a control from which other types will be derived, and should not be displayed to users.</summary>
    FControl = 0x0020,
    /// <summary>The interface supplies both <see langword="IDispatch" /> and V-table binding.</summary>
    FDual = 0x0040,
    /// <summary>The interface cannot add members at run time.</summary>
    FNonExtensible = 0x0080,
    /// <summary>The types used in the interface are fully compatible with Automation, including vtable binding support.</summary>
    FOleAutomation = 0x0100,
    /// <summary>This flag is intended for system-level types or types that type browsers should not display.</summary>
    FRestricted = 0x0200,
    /// <summary>The class supports aggregation.</summary>
    FAggregatable = 0x0400,
    /// <summary>The object supports <see langword="IConnectionPointWithDefault" />, and has default behaviors.</summary>
    FReplaceable = 0x0800,
    /// <summary>Indicates that the interface derives from <see langword="IDispatch" />, either directly or indirectly.</summary>
    FDispatchable = 0x1000,
    /// <summary>Indicates base interfaces should be checked for name resolution before checking child interfaces. This is the reverse of the default behavior.</summary>
    FReverseBind = 0x2000,
}

/// <summary>Describes the original settings of the <see langword="FUNCFLAGS" /> in the COM type library from where this method was imported.</summary>
[Serializable]
[Flags()]
[ComVisible(true)]
public enum TypeLibFuncFlags
{
    /// <summary>This flag is intended for system-level functions or functions that type browsers should not display.</summary>
    FRestricted = 0x0001,
    /// <summary>The function returns an object that is a source of events.</summary>
    FSource = 0x0002,
    /// <summary>The function that supports data binding.</summary>
    FBindable = 0x0004,
    /// <summary>When set, any call to a method that sets the property results first in a call to <see langword="IPropertyNotifySink::OnRequestEdit" />.</summary>
    FRequestEdit = 0x0008,
    /// <summary>The function that is displayed to the user as bindable. <see cref="F:System.Runtime.InteropServices.TypeLibFuncFlags.FBindable" /> must also be set.</summary>
    FDisplayBind = 0x0010,
    /// <summary>The function that best represents the object. Only one function in a type information can have this attribute.</summary>
    FDefaultBind = 0x0020,
    /// <summary>The function should not be displayed to the user, although it exists and is bindable.</summary>
    FHidden = 0x0040,
    /// <summary>The function supports <see langword="GetLastError" />.</summary>
    FUsesGetLastError = 0x0080,
    /// <summary>Permits an optimization in which the compiler looks for a member named "xyz" on the type "abc". If such a member is found and is flagged as an accessor function for an element of the default collection, then a call is generated to that member function.</summary>
    FDefaultCollelem = 0x0100,
    /// <summary>The type information member is the default member for display in the user interface.</summary>
    FUiDefault = 0x0200,
    /// <summary>The property appears in an object browser, but not in a properties browser.</summary>
    FNonBrowsable = 0x0400,
    /// <summary>Tags the interface as having default behaviors.</summary>
    FReplaceable = 0x0800,
    /// <summary>The function is mapped as individual bindable properties.</summary>
    FImmediateBind = 0x1000,
}

/// <summary>Describes the original settings of the <see cref="T:System.Runtime.InteropServices.VARFLAGS" /> in the COM type library from which the variable was imported.</summary>
[Serializable]
[Flags()]
[ComVisible(true)]
public enum TypeLibVarFlags
{
    /// <summary>Assignment to the variable should not be allowed.</summary>
    FReadOnly = 0x0001,
    /// <summary>The variable returns an object that is a source of events.</summary>
    FSource = 0x0002,
    /// <summary>The variable supports data binding.</summary>
    FBindable = 0x0004,
    /// <summary>Indicates that the property supports the COM <see langword="OnRequestEdit" /> notification.</summary>
    FRequestEdit = 0x0008,
    /// <summary>The variable is displayed as bindable. <see cref="F:System.Runtime.InteropServices.TypeLibVarFlags.FBindable" /> must also be set.</summary>
    FDisplayBind = 0x0010,
    /// <summary>The variable is the single property that best represents the object. Only one variable in a type info can have this value.</summary>
    FDefaultBind = 0x0020,
    /// <summary>The variable should not be displayed in a browser, though it exists and is bindable.</summary>
    FHidden = 0x0040,
    /// <summary>This flag is intended for system-level functions or functions that type browsers should not display.</summary>
    FRestricted = 0x0080,
    /// <summary>Permits an optimization in which the compiler looks for a member named "xyz" on the type "abc". If such a member is found and is flagged as an accessor function for an element of the default collection, then a call is generated to that member function.</summary>
    FDefaultCollelem = 0x0100,
    /// <summary>The default display in the user interface.</summary>
    FUiDefault = 0x0200,
    /// <summary>The variable appears in an object browser, but not in a properties browser.</summary>
    FNonBrowsable = 0x0400,
    /// <summary>Tags the interface as having default behaviors.</summary>
    FReplaceable = 0x0800,
    /// <summary>The variable is mapped as individual bindable properties.</summary>
    FImmediateBind = 0x1000,
}

/// <summary>Contains the <see cref="T:System.Runtime.InteropServices.TYPEFLAGS" /> that were originally imported for this type from the COM type library.</summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface | AttributeTargets.Enum | AttributeTargets.Struct, Inherited = false)]
[ComVisible(true)]
public sealed class TypeLibTypeAttribute : Attribute
{
    internal TypeLibTypeFlags _val;
    /// <summary>Initializes a new instance of the <see langword="TypeLibTypeAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value for the attributed type as found in the type library it was imported from.</param>
    public TypeLibTypeAttribute(TypeLibTypeFlags flags)
    {
        _val = flags;
    }
    /// <summary>Initializes a new instance of the <see langword="TypeLibTypeAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value for the attributed type as found in the type library it was imported from.</param>
    public TypeLibTypeAttribute(short flags)
    {
        _val = (TypeLibTypeFlags)flags;
    }
    /// <summary>Gets the <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value for this type.</summary>
    /// <returns>The <see cref="T:System.Runtime.InteropServices.TypeLibTypeFlags" /> value for this type.</returns>
    public TypeLibTypeFlags Value { get { return _val; } }
}

/// <summary>Contains the <see cref="T:System.Runtime.InteropServices.FUNCFLAGS" /> that were originally imported for this method from the COM type library.</summary>
[AttributeUsage(AttributeTargets.Method, Inherited = false)]
[ComVisible(true)]
public sealed class TypeLibFuncAttribute : Attribute
{
    internal TypeLibFuncFlags _val;
    /// <summary>Initializes a new instance of the <see langword="TypeLibFuncAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value for the attributed method as found in the type library it was imported from.</param>
    public TypeLibFuncAttribute(TypeLibFuncFlags flags)
    {
        _val = flags;
    }
    /// <summary>Initializes a new instance of the <see langword="TypeLibFuncAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value for the attributed method as found in the type library it was imported from.</param>
    public TypeLibFuncAttribute(short flags)
    {
        _val = (TypeLibFuncFlags)flags;
    }
    /// <summary>Gets the <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value for this method.</summary>
    /// <returns>The <see cref="T:System.Runtime.InteropServices.TypeLibFuncFlags" /> value for this method.</returns>
    public TypeLibFuncFlags Value { get { return _val; } }
}

/// <summary>Contains the <see cref="T:System.Runtime.InteropServices.VARFLAGS" /> that were originally imported for this field from the COM type library.</summary>
[AttributeUsage(AttributeTargets.Field, Inherited = false)]
[ComVisible(true)]
public sealed class TypeLibVarAttribute : Attribute
{
    internal TypeLibVarFlags _val;
    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.TypeLibVarAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value for the attributed field as found in the type library it was imported from.</param>
    public TypeLibVarAttribute(TypeLibVarFlags flags)
    {
        _val = flags;
    }
    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.TypeLibVarAttribute" /> class with the specified <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value.</summary>
    /// <param name="flags">The <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value for the attributed field as found in the type library it was imported from.</param>
    public TypeLibVarAttribute(short flags)
    {
        _val = (TypeLibVarFlags)flags;
    }
    /// <summary>Gets the <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value for this field.</summary>
    /// <returns>The <see cref="T:System.Runtime.InteropServices.TypeLibVarFlags" /> value for this field.</returns>
    public TypeLibVarFlags Value { get { return _val; } }
}

/// <summary>Specifies whether the type should be marshaled using the Automation marshaler or a custom proxy and stub.</summary>
[AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Class | AttributeTargets.Interface, Inherited = false)]
[ComVisible(true)]
public sealed class AutomationProxyAttribute : Attribute
{
    internal bool _val;
    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.AutomationProxyAttribute" /> class.</summary>
    /// <param name="val"><see langword="true" /> if the class should be marshaled using the Automation Marshaler; <see langword="false" /> if a proxy stub marshaler should be used.</param>
    public AutomationProxyAttribute(bool val)
    {
        _val = val;
    }
    /// <summary>Gets a value indicating the type of marshaler to use.</summary>
    /// <returns><see langword="true" /> if the class should be marshaled using the Automation Marshaler; <see langword="false" /> if a proxy stub marshaler should be used.</returns>
    public bool Value { get { return _val; } }
}

/// <summary>Specifies the version number of an exported type library.</summary>
[AttributeUsage(AttributeTargets.Assembly, Inherited = false)]
[ComVisible(true)]
public sealed class TypeLibVersionAttribute : Attribute
{
    internal int _major;
    internal int _minor;

    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.TypeLibVersionAttribute" /> class with the major and minor version numbers of the type library.</summary>
    /// <param name="major">The major version number of the type library.</param>
    /// <param name="minor">The minor version number of the type library.</param>
    public TypeLibVersionAttribute(int major, int minor)
    {
        _major = major;
        _minor = minor;
    }

    /// <summary>Gets the major version number of the type library.</summary>
    /// <returns>The major version number of the type library.</returns>
    public int MajorVersion { get { return _major; } }
    /// <summary>Gets the minor version number of the type library.</summary>
    /// <returns>The minor version number of the type library.</returns>
    public int MinorVersion { get { return _minor; } }
}

/// <summary>Provides support for user customization of interop stubs in managed-to-COM interop scenarios.</summary>
[AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
[ComVisible(false)]
public sealed class ManagedToNativeComInteropStubAttribute : Attribute
{
    internal Type _classType;
    internal string _methodName;

    /// <summary>Initializes a new instance of the <see cref="T:System.Runtime.InteropServices.ManagedToNativeComInteropStubAttribute" /> class with the specified class type and method name.</summary>
    /// <param name="classType">The class that contains the required stub method.</param>
    /// <param name="methodName">The name of the stub method.</param>
    /// <exception cref="T:System.ArgumentException">
    ///   <paramref name="methodName" /> cannot be found.  
    ///-or-  
    ///The method is not static or non-generic.  
    ///-or-  
    ///The method's parameter list does not match the expected parameter list for the stub.</exception>
    /// <exception cref="T:System.MethodAccessException">The interface that contains the managed interop method has no access to the stub method, because the stub method has private or protected accessibility, or because of a security issue.</exception>
    public ManagedToNativeComInteropStubAttribute(Type classType, string methodName)
    {
        _classType = classType;
        _methodName = methodName;
    }

    /// <summary>Gets the class that contains the required stub method.</summary>
    /// <returns>The class that contains the customized interop stub.</returns>
    public Type ClassType { get { return _classType; } }
    /// <summary>Gets the name of the stub method.</summary>
    /// <returns>The name of a customized interop stub.</returns>
    public string MethodName { get { return _methodName; } }
}

#endif
