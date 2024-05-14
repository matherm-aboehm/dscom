using System.Globalization;
using System.Reflection;
using System.Security;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROAssemblyExtended : Assembly
{
    internal readonly Assembly _roAssembly;
    public ROAssemblyExtended(Assembly roAssembly)
    {
        if (!roAssembly.ReflectionOnly)
        {
            throw new ArgumentOutOfRangeException(nameof(roAssembly), $"{nameof(roAssembly)} must be from reflection-only load context.");
        }
        _roAssembly = roAssembly;
    }

    #region Properties
    [Obsolete("Assembly.CodeBase and Assembly.EscapedCodeBase are only included for .NET Framework compatibility. Use Assembly.Location.")]
    public override string? CodeBase => _roAssembly.CodeBase;
    public override IEnumerable<CustomAttributeData> CustomAttributes => _roAssembly.CustomAttributes;
    public override IEnumerable<TypeInfo> DefinedTypes => _roAssembly.DefinedTypes;
    public override MethodInfo? EntryPoint => _roAssembly.EntryPoint;
    [Obsolete("Assembly.CodeBase and Assembly.EscapedCodeBase are only included for .NET Framework compatibility. Use Assembly.Location.")]
    public override string EscapedCodeBase => _roAssembly.EscapedCodeBase;
    public override string? FullName => _roAssembly.FullName;
    public override Module ManifestModule => new ROModuleExtended(this, _roAssembly.ManifestModule);
    [Obsolete("The Global Assembly Cache is not supported.")]
    public override bool GlobalAssemblyCache => _roAssembly.GlobalAssemblyCache;
    public override long HostContext => _roAssembly.HostContext;
    public override string ImageRuntimeVersion => _roAssembly.ImageRuntimeVersion;
    public override bool IsCollectible => _roAssembly.IsCollectible;
    public override bool IsDynamic => _roAssembly.IsDynamic;
    public override string Location => _roAssembly.Location;
    public override IEnumerable<Module> Modules => _roAssembly.Modules.Select(m => new ROModuleExtended(this, m));
    public override SecurityRuleSet SecurityRuleSet => _roAssembly.SecurityRuleSet;
    public override bool ReflectionOnly => _roAssembly.ReflectionOnly;
    #endregion
    #region Events
    public override event ModuleResolveEventHandler? ModuleResolve
    {
        add => _roAssembly.ModuleResolve += value;
        remove => _roAssembly.ModuleResolve -= value;
    }
    #endregion
    #region Methods
    public override object? CreateInstance(string typeName, bool ignoreCase, BindingFlags bindingAttr, Binder? binder, object[]? args, CultureInfo? culture, object[]? activationAttributes)
        => _roAssembly.CreateInstance(typeName, ignoreCase, bindingAttr, binder, args, culture, activationAttributes);
    public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        => GetCustomAttributesData().GetCustomAttributes(attributeType);
    public override object[] GetCustomAttributes(bool inherit)
        => GetCustomAttributesData().GetCustomAttributes();
    public override bool IsDefined(Type attributeType, bool inherit)
        => GetCustomAttributesData().IsDefined(attributeType);
    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roAssembly.GetCustomAttributesData();
    public override Type[] GetTypes()
        => _roAssembly.GetTypes().Select(t => new ROTypeExtended(this, t)).ToArray();
    public override Type[] GetExportedTypes()
        => _roAssembly.GetExportedTypes().Select(t => new ROTypeExtended(this, t)).ToArray();
    public override Type[] GetForwardedTypes()
        => _roAssembly.GetForwardedTypes().Select(t => new ROTypeExtended(this, t)).ToArray();
    public override FileStream? GetFile(string name)
        => _roAssembly.GetFile(name);
    public override FileStream[] GetFiles()
        => _roAssembly.GetFiles();
    public override FileStream[] GetFiles(bool getResourceModules)
        => _roAssembly.GetFiles(getResourceModules);
    public override Module[] GetLoadedModules(bool getResourceModules)
        => _roAssembly.GetLoadedModules(getResourceModules).Select(m => new ROModuleExtended(this, m)).ToArray();
    public override ManifestResourceInfo? GetManifestResourceInfo(string resourceName)
        => _roAssembly.GetManifestResourceInfo(resourceName);
    public override string[] GetManifestResourceNames()
        => _roAssembly.GetManifestResourceNames();
    public override Stream? GetManifestResourceStream(Type type, string name)
        => _roAssembly.GetManifestResourceStream(type, name);
    public override Stream? GetManifestResourceStream(string name)
        => _roAssembly.GetManifestResourceStream(name);
    public override Module? GetModule(string name)
    {
        var module = _roAssembly.GetModule(name);
        return module is null ? null : new ROModuleExtended(this, module);
    }
    public override Module[] GetModules(bool getResourceModules)
        => _roAssembly.GetModules(getResourceModules).Select(m => new ROModuleExtended(this, m)).ToArray();
    public override AssemblyName GetName() => _roAssembly.GetName();
    public override AssemblyName GetName(bool copiedName) => _roAssembly.GetName(copiedName);
    public override AssemblyName[] GetReferencedAssemblies() => _roAssembly.GetReferencedAssemblies();
    public override Assembly GetSatelliteAssembly(CultureInfo culture)
        => new ROAssemblyExtended(_roAssembly.GetSatelliteAssembly(culture));
    public override Assembly GetSatelliteAssembly(CultureInfo culture, Version? version)
        => new ROAssemblyExtended(base.GetSatelliteAssembly(culture, version));
    public override Type? GetType(string name)
    {
        var type = _roAssembly.GetType(name);
        return type is null ? null : new ROTypeExtended(this, type);
    }
    public override Type? GetType(string name, bool throwOnError)
    {
        var type = _roAssembly.GetType(name, throwOnError);
        return type is null ? null : new ROTypeExtended(this, type);
    }
    public override Type? GetType(string name, bool throwOnError, bool ignoreCase)
    {
        var type = _roAssembly.GetType(name, throwOnError, ignoreCase);
        return type is null ? null : new ROTypeExtended(this, type);
    }
    public override Module LoadModule(string moduleName, byte[]? rawModule, byte[]? rawSymbolStore)
        => new ROModuleExtended(this, _roAssembly.LoadModule(moduleName, rawModule, rawSymbolStore));
    #endregion

    #region Object overrides
    public override bool Equals(object? o)
        => (o is ROAssemblyExtended extended) && _roAssembly.Equals(extended._roAssembly);
    public override int GetHashCode() => _roAssembly.GetHashCode();
    public override string ToString() => _roAssembly.ToString();
    #endregion
}
