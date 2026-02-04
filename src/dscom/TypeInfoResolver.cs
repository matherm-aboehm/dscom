// Copyright 2022 dSPACE GmbH, Mark Lechtermann, Matthias Nissen and Contributors
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using dSPACE.Runtime.InteropServices.Exporter;

namespace dSPACE.Runtime.InteropServices;

internal sealed class TypeInfoResolver : ITypeLibCache
{
    private readonly IDictionary<TypeLibIdentifier, ITypeLib> _typeLibs = new Dictionary<TypeLibIdentifier, ITypeLib>();

    private readonly List<string> _additionalLibs = new();

    private readonly Dictionary<Guid, ITypeInfo> _types = new();

    private readonly Dictionary<TypeLibIdentifier, IDictionary<string, ITypeInfo>> _nameOnlyTypes = new();

    private readonly Dictionary<Type, ITypeInfo?> _resolvedTypeInfos = new();

    public WriterContext WriterContext { get; }

    public TypeInfoResolver(WriterContext writerContext)
    {
        WriterContext = writerContext;

        if (WriterContext.NotifySink is ITypeLibCacheProvider typeLibCacheProvider)
        {
            typeLibCacheProvider.TypeLibCache = this;
        }

        AddTypeLib("stdole", Constants.LCID_NEUTRAL, new Guid(Guids.TLBID_Ole), 2, 0);

        _additionalLibs.AddRange(WriterContext.Options.TLBReference);

        foreach (var directoriesWithTypeLibraries in WriterContext.Options.TLBRefpath)
        {
            if (Directory.Exists(directoriesWithTypeLibraries))
            {
                _additionalLibs.AddRange(Directory.GetFiles(directoriesWithTypeLibraries, "*.tlb"));
            }
        }

        var fullPath = !string.IsNullOrEmpty(WriterContext.Options.Out) ? Path.GetFullPath(WriterContext.Options.Out) : string.Empty;

        var pathToRemove = _additionalLibs.FirstOrDefault(p => string.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase));
        if (pathToRemove != null)
        {
            _additionalLibs.Remove(pathToRemove);
        }

        _additionalLibs = _additionalLibs.Distinct().ToList();
    }

    public ITypeInfo? ResolveTypeInfo(Guid guid)
    {
        _types.TryGetValue(guid, out var typeInfo);
        return typeInfo;
    }

    public ITypeInfo? ResolveTypeInfo(in TypeLibIdentifier identifier, string name)
    {
        if (_nameOnlyTypes.TryGetValue(identifier, out var nameOnlyTypesFromLib))
        {
            nameOnlyTypesFromLib.TryGetValue(name, out var typeInfo);
            return typeInfo;
        }
        return null;
    }

    /// <summary>
    /// Resolve the <see cref="ITypeInfo"/> by a <see cref="Guid"/>. If the <see cref="ITypeInfo"/> is not
    /// present yet, it will try to add the type library for the corresponding assembly.
    /// </summary>
    private ITypeInfo? ResolveTypeInfo(Type type, Guid guid)
    {
        // If the given type is not present in _types
        // this means, that the typelib is not loaded yet
        if (_types.TryGetValue(guid, out var typeInfo))
        {
            return typeInfo;
        }

        var assembly = type.Assembly;

        if ((assembly.ReflectionOnly && type.GUID == Guid.Empty) ||
            (!assembly.ReflectionOnly && type.IsDefined(typeof(GuidAttribute), false)))
        {
            var identifierOle = new TypeLibIdentifier()
            {
                Name = type.Name,
                LanguageIdentifier = Constants.LCID_NEUTRAL,
                LibID = new Guid(Guids.TLBID_Ole),
                MajorVersion = 2,
                MinorVersion = 0
            };

            typeInfo = ResolveTypeInfo(identifierOle, type.Name);
            if (typeInfo is not null)
            {
                var oleType = new TypeInfo((ITypeInfo2)typeInfo, null, nameof(typeInfo));
                //TODO: check similarity of COM type and managed type before returning
                return typeInfo;
            }
        }

        // check if the type library is already present
        var identifier = assembly.GetLibIdentifier(WriterContext.Options.OverrideTlbId);
        var typeLib = GetTypeLibFromIdentifier(identifier);
        if (typeLib is null)
        {
            var name = assembly.GetName().Name ?? string.Empty;

            var additionalLibsWithMatchingName = _additionalLibs
                .Where(additionalLib => Path.GetFileNameWithoutExtension(additionalLib).Equals(name, StringComparison.OrdinalIgnoreCase));

            // At first we try to find a type library that matches the assembly name.
            // We do this to limit the number of type libraries to load.
            // See https://github.com/dspace-group/dscom/issues/310
            foreach (var additionalLib in additionalLibsWithMatchingName)
            {
                if (AddTypeLib(additionalLib) &&
                    (typeLib = GetTypeLibFromIdentifier(identifier)) != null)
                {
                    MarshalExtension.UpdateCacheFromTypeLib(typeLib, assembly);
                    break;
                }
            }

            // If no type was found in a type library with matching name we search in the remaining type libraries.
            if (typeLib is null)
            {
                var additionalLibsWithoutMatchingName = _additionalLibs.Except(additionalLibsWithMatchingName);
                foreach (var additionalLib in additionalLibsWithoutMatchingName)
                {
                    if (AddTypeLib(additionalLib) &&
                        (typeLib = GetTypeLibFromIdentifier(identifier)) != null)
                    {
                        MarshalExtension.UpdateCacheFromTypeLib(typeLib, assembly);
                        break;
                    }
                }
            }

            var notifySink = WriterContext.NotifySink;
            if (typeLib is null && notifySink != null)
            {
                if (notifySink.ResolveRef(assembly) is ITypeLib refTypeLib)
                {
                    if (AddTypeLib(typeLib = refTypeLib))
                    {
                        // In this case, type library was not created by notifySink using local TypeLibConverter,
                        // but instead was loaded from existing file, so cache in MarshalExtension needs to be updated
                        // for GetClassInterfaceGuidForType based on type infos in this type lib.
                        MarshalExtension.UpdateCacheFromTypeLib(typeLib, assembly);
                    }
                }
            }

            if (typeLib != null && type.IsClass)
            {
                // try to fetch again from cache after ResolveRef callback has
                // created new TLB containing this type.
                var classItfGuid = MarshalExtension.GetClassInterfaceGuidForType(type);
                if (classItfGuid != Guid.Empty)
                {
                    guid = classItfGuid;
                }
            }
        }

        // The dictionary will be updated in 'AddTypeLib'.
        // Therefore it should be contain the typeinfo now.
        _types.TryGetValue(guid, out typeInfo);
        return typeInfo;
    }

    public ITypeInfo? ResolveTypeInfo(Type type)
    {
        // special handling for non com visible types and
        // build-in-types. Otherwise it will try to load a typelib for string or object
        if (type.IsSpecialHandledClass())
        {
            Debug.Assert(false, "ResolveTypeInfo should not be called for string or object.");
            return null;
        }

#pragma warning disable IDE0045 // Convert to conditional expression
        if (_resolvedTypeInfos.TryGetValue(type, out var typeInfo))
        {
            return typeInfo;
        }

        ITypeInfo? retval;
        if (type.FullName == "System.Collections.IEnumerator")
        {
            retval = ResolveTypeInfo(new Guid(Guids.IID_IEnumVARIANT));
        }
        else if (type.FullName == "System.Drawing.Color")
        {
            retval = ResolveTypeInfo(new Guid(Guids.TDID_OLECOLOR));
        }
        else if (type.FullName == "System.Guid")
        {
            var identifierOle = new TypeLibIdentifier()
            {
                Name = "System.Guid",
                LanguageIdentifier = Constants.LCID_NEUTRAL,
                LibID = new Guid(Guids.TLBID_Ole),
                MajorVersion = 2,
                MinorVersion = 0
            };
            retval = ResolveTypeInfo(identifierOle, "GUID") ?? throw new COMException("System.Guid not found in any type library");
        }
        else if (type.GUID == new Guid(Guids.IID_IDispatch))
        {
            retval = ResolveTypeInfo(new Guid(Guids.IID_IDispatch));
        }
        else
        {
            var typeGuid = type.IsClass ? MarshalExtension.GetClassInterfaceGuidForType(type)
                : MarshalExtension.GenerateGuidForType(type);
            //HINT: When there was a ClassInterfaceWriter used for the type, the Guid should
            // be cached in MarshalExtension, if not, then no ClassInterfaceWriter was used
            // and Guid of the type itself should be used.
            if (type.IsClass && typeGuid == Guid.Empty)
            {
                typeGuid = MarshalExtension.GenerateGuidForType(type);
            }
            retval = ResolveTypeInfo(type, typeGuid);
        }
#pragma warning restore IDE0045 // Convert to conditional expression

        _resolvedTypeInfos[type] = retval;
        return retval;
    }

    private static ITypeInfo? GetDefaultInterface(ITypeInfo? typeInfo)
    {
        if (typeInfo != null)
        {
            var ppTypeAttr = new IntPtr();
            try
            {
                typeInfo.GetTypeAttr(out ppTypeAttr);
                var typeAttr = Marshal.PtrToStructure<TYPEATTR>(ppTypeAttr);

                var typeInfo64Bit = (ITypeInfo64Bit)typeInfo;
                for (var i = 0; i < typeAttr.cImplTypes; i++)
                {
                    typeInfo64Bit.GetRefTypeOfImplType(i, out var href);
                    typeInfo64Bit.GetRefTypeInfo(href, out var refTypeInfo);
                    typeInfo.GetImplTypeFlags(i, out var pImplTypeFlags);

                    refTypeInfo.GetDocumentation(-1, out var m, out _, out _, out _);

                    if (pImplTypeFlags.HasFlag(IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT))
                    {
                        return (ITypeInfo)refTypeInfo;
                    }
                }
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(ppTypeAttr);
            }
        }

        return null;
    }

    internal ITypeInfo? ResolveDefaultCoClassInterface(Type type)
    {
        if (type.IsClass)
        {
            var typeGuid = MarshalExtension.GenerateGuidForType(type);
            var typeInfo = ResolveTypeInfo(typeGuid);
            return GetDefaultInterface(typeInfo);
        }

        return null;
    }

    public ITypeLib? GetTypeLibFromIdentifier(in TypeLibIdentifier identifier)
    {
        _typeLibs.TryGetValue(identifier, out var value);
        return value;
    }

    public static TypeLibIdentifier GetIdentifierFromTypeLib(ITypeLib typeLib)
    {
        typeLib.GetLibAttr(out var libattrPtr);
        try
        {
            var libattr = Marshal.PtrToStructure<TYPELIBATTR>(libattrPtr);
            typeLib.GetDocumentation(-1, out var name, out var strDocString, out var dwHelpContext, out var strHelpFile);
            var identifier = new TypeLibIdentifier
            {
                Name = name,
                LanguageIdentifier = libattr.lcid,
                LibID = libattr.guid,
                MajorVersion = (ushort)libattr.wMajorVerNum,
                MinorVersion = (ushort)libattr.wMinorVerNum
            };
            return identifier;
        }
        finally
        {
            typeLib.ReleaseTLibAttr(libattrPtr);
        }
    }

    public ITypeLib? LoadTypeLibFromIdentifier(in TypeLibIdentifier identifier, bool throwOnError = true)
    {
        HRESULT hr;
        string typeLibPath = null!;
        try
        {
            hr = OleAut32.QueryPathOfRegTypeLib(identifier.LibID, identifier.MajorVersion, identifier.MinorVersion,
                Constants.LCID_NEUTRAL, out typeLibPath);
        }
        catch (Exception ex) when (!throwOnError)
        {
            hr = (HRESULT)ex.HResult;
        }

        if (hr.Succeeded)
        {
            try
            {
                return LoadTypeLibFromPath(typeLibPath);
            }
            catch (Exception ex) when (!throwOnError || WriterContext.Options.Create64BitTlb.HasValue)
            {
                hr = (HRESULT)ex.HResult;
            }

            // Fallback for downlevel platform, if target platform was specified
            if (WriterContext.Options.Create64BitTlb.HasValue)
            {
                hr = OleAut32.LoadRegTypeLib(identifier.LibID, identifier.MajorVersion, identifier.MinorVersion,
                    Constants.LCID_NEUTRAL, out var typeLib);
                if (hr.Succeeded)
                {
                    return typeLib;
                }
            }
        }

        if (throwOnError)
        {
            hr.ThrowIfFailed($"Failed to load type library {identifier.LibID}.");
        }

        return null;
    }

    public ITypeLib? LoadTypeLibFromPath(string typeLibPath, bool throwOnError = true)
    {
        var regkind = REGKIND.NONE;
        regkind |= WriterContext.Options.Create64BitTlb switch
        {
            true => REGKIND.LOAD_TLB_AS_64BIT,
            false => REGKIND.LOAD_TLB_AS_32BIT,
            _ => 0
        };
        try
        {
            var hr = OleAut32.LoadTypeLibEx(typeLibPath, regkind, out var typeLib);
            if (throwOnError)
            {
                hr.ThrowIfFailed($"Failed to load type library {typeLibPath}.");
            }
            else if (hr.Failed)
            {
                return null;
            }
            return typeLib;
        }
        catch when (!throwOnError)
        {
            return null;
        }
    }

    public void AddTypeLib(string name, int lcid, Guid registeredTypeLib, ushort majorVersion, ushort minorVersion)
    {
        var identifier = new TypeLibIdentifier
        {
            Name = name,
            LanguageIdentifier = lcid,
            LibID = registeredTypeLib,
            MajorVersion = majorVersion,
            MinorVersion = minorVersion
        };

        var typeLib = LoadTypeLibFromIdentifier(identifier)!;

        _typeLibs.Add(identifier, typeLib);
        UpdateTypeLibCache(identifier, typeLib);
    }

    public bool AddTypeLib(ITypeLib typeLib)
    {
        var identifier = GetIdentifierFromTypeLib(typeLib);
        return AddTypeLib(typeLib, identifier);
    }

    public bool AddTypeLib(ITypeLib typeLib, in TypeLibIdentifier identifier)
    {
        if (!_typeLibs.ContainsKey(identifier))
        {
            _typeLibs.Add(identifier, typeLib);
            UpdateTypeLibCache(identifier, typeLib);
            return true;
        }
        return false;
    }

    private void UpdateTypeLibCache(in TypeLibIdentifier identifier, ITypeLib typelib)
    {
        var count = typelib.GetTypeInfoCount();
        for (var i = 0; i < count; i++)
        {
            typelib.GetTypeInfo(i, out var typeInfo);
            AddTypeToCache(typeInfo, identifier);
        }
    }

    public void AddTypeToCache(ITypeInfo? typeInfo, in TypeLibIdentifier? identifier = null)
    {
        if (typeInfo != null)
        {
            typeInfo.GetTypeAttr(out var ppTypAttr);
            try
            {
                var attr = Marshal.PtrToStructure<TYPEATTR>(ppTypAttr);
                if (attr.guid == Guid.Empty)
                {
                    typeInfo.GetDocumentation(-1, out var name, out _, out _, out _);
                    if (identifier is null)
                    {
                        WriterContext.LogWarning($"Warning: No GUID was defined for the type {name}. The cache to resolve type infos, can't be updated.");
                        return;
                    }
                    if (!_nameOnlyTypes.TryGetValue(identifier.Value, out var nameOnlyTypesFromLib))
                    {
                        _nameOnlyTypes.Add(identifier.Value, nameOnlyTypesFromLib = new Dictionary<string, ITypeInfo>());
                    }
                    if (!nameOnlyTypesFromLib.ContainsKey(name))
                    {
                        nameOnlyTypesFromLib.Add(name, typeInfo);
                    }
                }
                else if (!_types.ContainsKey(attr.guid))
                {
                    _types[attr.guid] = typeInfo;
                }
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(ppTypAttr);
            }
        }
    }

    [ExcludeFromCodeCoverage] // UnitTest with dependent type libraries is not supported
    public bool AddTypeLib(string typeLibPath)
    {
        var typeLib = LoadTypeLibFromPath(typeLibPath)!;
        return AddTypeLib(typeLib);
    }
}
