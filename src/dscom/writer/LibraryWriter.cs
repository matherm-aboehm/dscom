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

#pragma warning disable 612, 618

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace dSPACE.Runtime.InteropServices.Writer;

internal sealed class LibraryWriter : BaseWriter
{
    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_IInternetSecurityManager = new("79eac9ee-baf9-11ce-8c82-00aa004ba90b");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_IInternetSecurityMgrSite = new("79eac9ed-baf9-11ce-8c82-00aa004ba90b");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_IObjectArray = new("92ca9dcd-5622-4bba-a805-5e9f541bd8c9");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_IObjectCollection = new("5632b1a4-e38a-400a-928a-d4cd63230295");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_ITaskbarList3 = new("ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_ITaskbarList4 = new("c43dc798-95d1-4bea-9030-bb99e2983a1a");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid IID_IPersistStream = new("00000109-0000-0000-C000-000000000046");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Dictionary<Guid, (string, Guid?)> GuidOverrideFixupTable = new()
    {
        {IID_IInternetSecurityManager, ("IInternetSecurityMgrSite", IID_IInternetSecurityMgrSite)},
        {IID_IObjectArray, ("IObjectCollection", IID_IObjectCollection)},
        {IID_ITaskbarList3, ("ITaskbarList4", IID_ITaskbarList4)},
        {IID_IPersistStream, ("IPersistStreamWithArrays", null)}, // null for overrideGuid means skip instead of override
    };

    public LibraryWriter(Assembly assembly, WriterContext context) : base(context)
    {
        Assembly = assembly;
    }

    private Assembly Assembly { get; }

    private List<TypeWriter> TypeWriters { get; } = new();

    private Dictionary<Guid, (Type type, TypeWriter writer)> TypeWritersForComImport { get; } = new();

    private Dictionary<string, object> UniqueNames { get; } = new();

    public TypeLibIdentifier Identifier { get; private set; }

    public override void Create()
    {
        var name = string.IsNullOrEmpty(Context.Options.OverrideName)
            ? Assembly.GetName().Name!.Replace('.', '_')
            : Context.Options.OverrideName;

        var versionFromAssembly = Assembly.GetTLBVersionForAssembly();
        var guid = Context.Options.OverrideTlbId == Guid.Empty ? Assembly.GetTLBGuidForAssembly() : Context.Options.OverrideTlbId;

        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(LibraryWriter));
        }

        var typeLib = Context.TargetTypeLib;
        if (typeLib != null)
        {
            typeLib.SetGuid(guid);
            typeLib.SetVersion(versionFromAssembly.Major == 0 ? (ushort)1 : (ushort)versionFromAssembly.Major, (ushort)versionFromAssembly.Minor);
            typeLib.SetLcid(Constants.LCID_NEUTRAL);
            typeLib.SetCustData(new Guid(Guids.GUID_ExportedFromComPlus), Assembly.FullName);
            typeLib.SetName(name);
            var description = Assembly.GetCustomAttributes<AssemblyDescriptionAttribute>().FirstOrDefault()?.Description;
            if (description != null)
            {
                typeLib.SetDocString(description);
            }

            Identifier = TypeInfoResolver.GetIdentifierFromTypeLib((ITypeLib)typeLib);
            Context.TypeInfoResolver.AddTypeLib((ITypeLib)typeLib, Identifier);

            if (!string.IsNullOrEmpty(Context.Options.Out))
            {
                typeLib.SetLibFlags((uint)LIBFLAGS.LIBFLAG_FHASDISKIMAGE).ThrowIfFailed("Failed to set LIBFLAGS LIBFLAG_FHASDISKIMAGE.");
            }
        }

        CollectAllTypes();

        // Create and prepare the CreateTypeInfo instance.
        TypeWriters.ForEach(e => e.CreateTypeInfo());

        // Create the type inheritance
        TypeWriters.ForEach(e => e.CreateTypeInheritance());

        // Fill the content of methods and enums.
        TypeWriters.ForEach(e => e.Create());
    }

    private void CollectAllTypes()
    {
        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(LibraryWriter));
        }

        var comVisibleAttributeAssembly = Assembly.GetCustomAttribute<ComVisibleAttribute>();
        var typesAreVisibleForComByDefault = comVisibleAttributeAssembly == null || comVisibleAttributeAssembly.Value;
        var classIntfAttributeAssembly = Assembly.GetCustomAttribute<ClassInterfaceAttribute>();
        var defaultClassInterfaceType = classIntfAttributeAssembly?.Value ?? ClassInterfaceType.AutoDispatch;

        var types = Assembly.GetLoadableTypesAndLog(Context);

        Dictionary<Type, ClassInterfaceWriter> classInterfaceWriters = new();
        Dictionary<Type, ClassWriter> classWriters = new();

        foreach (var type in types)
        {
            var isImport = false;
            // If it is ComImport or WindowsRuntimeImport, skip other checks.
            if (type.IsImport || type.IsProjectedFromWinRT())
            {
                if (type.GUID != Guid.Empty)
                {
                    //TODO: Implement an IDL Parser and use IDL files from Windows SDK
                    // to write correct type infos to the TLB
                    if (Context.TypeInfoResolver.ResolveTypeInfo(type.GUID) is ITypeInfo2 typeInfo)
                    {
                        continue;
                    }
                    isImport = true;
                    var skip = false;
                    while (TypeWritersForComImport.TryGetValue(type.GUID, out var otherTypeWriter))
                    {
                        //Filter out class interfaces from TLB import like in PresentationFramework:
                        //System.Windows.Documents.MsSpellCheckLib.RCW+SpellCheckerFactoryClass
                        //System.Windows.Documents.MsSpellCheckLib.RCW+ISpellCheckerFactory
                        if (type.IsDefined(typeof(CoClassAttribute)) && !otherTypeWriter.type.IsDefined(typeof(CoClassAttribute)))
                        {
                            break;
                        }

                        if (SkipImportTypeOrOverrideGuid(type, otherTypeWriter.type, otherTypeWriter.writer))
                        {
                            skip = true;
                            break;
                        }
                        // If GUID was overridden on current type then
                        // try again to find duplicates with new GUID.
                        // If GUID was overridden on other type then it
                        // will break out of loop when next lookup fails.
                    }
                    if (skip)
                    {
                        continue;
                    }
                }
            }
            else
            {
                // Perf: Don't need to check for array in this special case, because
                // type declarations from an assembly should never be an array,
                // only referenced types can be array types.
                /*if (type.IsArray)
                {
                    continue;
                }*/

                var comVisibleAttribute = type.GetCustomAttribute<ComVisibleAttribute>();

                if (typesAreVisibleForComByDefault && comVisibleAttribute != null && !comVisibleAttribute.Value)
                {
                    continue;
                }

                if (!typesAreVisibleForComByDefault && (comVisibleAttribute == null || !comVisibleAttribute.Value))
                {
                    continue;
                }

                if (!type.IsPublic && !type.IsNestedPublicRecursive())
                {
                    continue;
                }

                if ((type.IsGenericType || type.IsGenericParameter) &&
                    !type.SupportsGenericInterop(WinRTExtensions.InteropKind.NativeToManaged))
                {
                    continue;
                }
            }
            // Add this type to the unique names collection.
            UpdateUniqueNames(type);

            TypeWriter? typeWriter = null;
            if (type.IsInterface)
            {
                var interfaceTypeAttribute = type.GetCustomAttribute<InterfaceTypeAttribute>();

                if (interfaceTypeAttribute == null && type.IsImport && type.IsDefined(typeof(CoClassAttribute)))
                {
                    // Type is a dummy interface type for a COM imported co-class.
                    // It probably isn't the default interface for the class, so find
                    // that first. Use this type as the class interface exported, if
                    // the class wants a class interface and the default interface isn't
                    // found.
                    var coClass = type.GetCustomAttribute<CoClassAttribute>()!.CoClass;
                    var classInterfaceType = coClass.GetCustomAttribute<ClassInterfaceAttribute>()?.Value ?? defaultClassInterfaceType;
                    var defaultInterfaceType = coClass.GetInterfaces().SingleOrDefault(t => !t.Equals(type) && t.GUID == type.GUID);
                    if (classInterfaceType != ClassInterfaceType.None)
                    {
                        if (defaultInterfaceType is not null)
                        {
                            throw new InvalidOperationException($"Can't create class interface for co-class {coClass.FullName} with same GUID as the default interface {defaultInterfaceType.FullName}");
                        }
                        if (classInterfaceWriters.TryGetValue(coClass, out var classInterfaceWriter))
                        {
                            // Class was visited first and class interface writer was already created for it.
                            // So just inject the type into it.
                            classInterfaceWriter.ComImportType = type;
                        }
                        else
                        {
                            classInterfaceWriter = WriterFactory.CreateInstance(new ClassInterfaceWriter.FactoryArgs(classInterfaceType, coClass, type, this, Context));
                            classInterfaceWriters.Add(coClass, classInterfaceWriter);
                            if (classWriters.TryGetValue(coClass, out var classWriter))
                            {
                                classWriter.ClassInterfaceWriter = classInterfaceWriter;
                            }
                        }
                    }
                    else if (defaultInterfaceType is not null)
                    {
                        // If class interface writer was already created because class
                        // was visited first and there are other conditions to enable
                        // the creation than what is checked here, we would need to
                        // dispose and remove that writer.
                        /*if (classInterfaceWriters.TryGetValue(coClass, out var classInterfaceWriter))
                        {
                            classInterfaceWriter.Dispose();
                            classInterfaceWriters.Remove(coClass);
                        }*/
                        // But for now, that the conditions are the same for enablement,
                        // just check, when already visited, if TypeWriter found by GUID
                        // was created for default interface type instead of other type.
                        if (TypeWritersForComImport.TryGetValue(defaultInterfaceType.GUID, out var otherTypeWriter)
                            && !defaultInterfaceType.Equals(otherTypeWriter.type))
                        {
                            // should never reach here, it would mean there are multiple interface types which defined the same GUID
                            // and where none of them are the dummy interface imported by TLB import
                            throw new InvalidOperationException($"The default interface for the co-class {coClass.FullName} can't be determined, at least 2 types share the same GUID {defaultInterfaceType.GUID}, where {defaultInterfaceType.FullName} is in the interface list from the co-class and {otherTypeWriter.type.FullName} is not.");
                        }
                        // If interface type wasn't visited yet or it's the same type as
                        // found by GUID from current type, then there is currently no
                        // other dupplicate GUID, so just skip the current type to not change that situation.
                    }
                    else
                    {
                        throw new InvalidOperationException("The co-class is missing a default interface.");
                    }
                }
                else
                {
                    typeWriter = interfaceTypeAttribute != null
                        ? interfaceTypeAttribute.Value switch
                        {
                            ComInterfaceType.InterfaceIsDual => WriterFactory.CreateInstance(new DualInterfaceWriter.FactoryArgs(type, this, Context)),
                            ComInterfaceType.InterfaceIsIDispatch => WriterFactory.CreateInstance(new DispInterfaceWriter.FactoryArgs(type, this, Context)),
                            ComInterfaceType.InterfaceIsIUnknown => WriterFactory.CreateInstance(new IUnknownInterfaceWriter.FactoryArgs(type, this, Context)),
                            _ => throw new NotSupportedException($"{interfaceTypeAttribute.Value} not supported"),
                        }
                        : WriterFactory.CreateInstance(new DualInterfaceWriter.FactoryArgs(type, this, Context));
                }
            }
            else if (type.IsEnum)
            {
                typeWriter = new EnumWriter(type, this, Context);
            }
            else if ((type.IsValueType && !type.IsPrimitive) || type.IsLayoutSequential
                || type.IsExplicitLayout)
            {
                typeWriter = new StructWriter(type, this, Context);
            }
            else if (type.IsClass)
            {
                typeWriter = new ClassWriter(type, this, Context);
                classWriters.Add(type, (ClassWriter)typeWriter);
            }
            if (typeWriter != null)
            {
                if (isImport)
                {
                    TypeWritersForComImport.Add(type.GUID, (type, typeWriter));
                }

                TypeWriters.Add(typeWriter);
            }

            if (type.IsClass)
            {
                //check for already created class interface writers
                if (type.IsImport && classInterfaceWriters.TryGetValue(type, out var classInterfaceWriter))
                {
                    if (typeWriter is ClassWriter classWriter)
                    {
                        Debug.Assert(classWriter.ClassInterfaceWriter is null);
                        classWriter.ClassInterfaceWriter = classInterfaceWriter;
                    }
                    continue;
                }

                var createClassInterface = true;
                //check for class interfaces to generate:
                var classInterfaceType = type.GetCustomAttribute<ClassInterfaceAttribute>()?.Value ?? defaultClassInterfaceType;
                switch (classInterfaceType)
                {
                    case ClassInterfaceType.AutoDispatch:
                        createClassInterface = true;
                        break;
                    case ClassInterfaceType.AutoDual:
                        if (!Context.Options.AllowAutoDual)
                        {
                            //CA1408: Do not use AutoDual ClassInterfaceType
                            //https://docs.microsoft.com/en-us/visualstudio/code-quality/ca1408?view=vs-2022
                            throw new NotSupportedException("Dual class interfaces not supported!");
                        }
                        createClassInterface = true;
                        break;
                    case ClassInterfaceType.None:
                        createClassInterface = false;
                        break;
                }

                //check for generic base types
                var baseType = type.BaseType;
                while (baseType != null)
                {
                    if (baseType.IsGenericType)
                    {
                        createClassInterface = false;
                        break;
                    }
                    baseType = baseType.BaseType;
                }

                if (createClassInterface)
                {
                    if (typeWriter is ClassWriter classWriter)
                    {
                        classInterfaceWriter = WriterFactory.CreateInstance(new ClassInterfaceWriter.FactoryArgs(classInterfaceType, type, null, this, Context));
                        classInterfaceWriters.Add(type, classInterfaceWriter);
                        classWriter.ClassInterfaceWriter = classInterfaceWriter;
                    }
                }
            }
        }
        TypeWriters.AddRange(classInterfaceWriters.Values);
    }

    private bool SkipImportTypeOrOverrideGuid(Type currentType, Type otherType, TypeWriter otherTypeWriter)
    {
        var skip = false;
        string? skipreason = null;
        // IInternetSecurityMgrSite and IInternetSecurityManager from WindowsBase
        // have defined the same GUID, which can't be exported to the same TLB,
        // so override the GUID or filter the wrong type out and continue on.
        // The correct GUID for IInternetSecurityMgrSite would be
        // 79eac9ed-baf9-11ce-8c82-00aa004ba90b, see UrlMon.idl in Windows SDK.

        // There are more of these GUID conflicts in .NET BCL, just find the correct
        // definitions in Windows SDK and add it to the fixup table.
        if (GuidOverrideFixupTable.TryGetValue(currentType.GUID, out var overrideGuidTuple))
        {
            var (typeNameForGuidOverrideOrSkip, overrideGuid) = overrideGuidTuple;
            if (otherType.Name == typeNameForGuidOverrideOrSkip)
            {
                TypeWritersForComImport.Remove(currentType.GUID);
                if (overrideGuid is not null && otherType is ROTypeExtended roType)
                {
                    roType.OverrideGUID(overrideGuid.Value);
                    TypeWritersForComImport.Add(overrideGuid.Value, (otherType, otherTypeWriter));
                }
                else
                {
                    otherTypeWriter.Dispose();
                    TypeWriters.Remove(otherTypeWriter);
                }
            }
            else if (currentType.Name == typeNameForGuidOverrideOrSkip)
            {
                if (overrideGuid is not null && currentType is ROTypeExtended roType)
                {
                    roType.OverrideGUID(overrideGuid.Value);
                }
                else
                {
                    skip = true;
                }
            }
            else if (currentType.Name == otherType.Name)
            {
                // no warning, as other types may be handled already 
                skip = true;
            }
        }
        else
        {
            skipreason = currentType.Name == otherType.Name
                ? $"There are ambiguous types with name {currentType.Name} that differs only in namespace ({currentType.Namespace}|{otherType.Namespace}) which both have defined the same GUID {currentType.GUID}."
                : $"Different types {currentType.FullName} and {otherType.FullName} have defined the same GUID {currentType.GUID}.";
            skip = true;
        }


        if (skip && skipreason is not null)
        {
            Context.LogWarning($"Warning: {skipreason} So both types can't be exported to the same TLB. {currentType.FullName} is skipped.", HRESULT.TYPE_E_DUPLICATEID);
        }

        return skip;
    }

    private static string GetFullName(Type type)
    {
        if (type.Namespace is null)
        {
            return type.Name;
        }
        StringBuilder namesp = new(type.Namespace);
        namesp = namesp.Replace(".", "_");
        List<string> nestedTypeNames = new();
        for (var parentType = type.IsNested ? type.DeclaringType : null;
            parentType != null;
            parentType = parentType.IsNested ? parentType.DeclaringType : null)
        {
            nestedTypeNames.Add(parentType.Name);
        }
        nestedTypeNames.Reverse();
        foreach (var name in nestedTypeNames)
        {
            namesp.Append('_');
            namesp.Append(name);
        }
        return $"{namesp}_{type.Name}";
    }

    private void UpdateUniqueNames(Type type)
    {
        if (UniqueNames.TryGetValue(type.Name, out var unique))
        {
            if (unique is not Dictionary<Type, string> ambiguousTypes)
            {
                var (searchExistingType, _) = ((Type, string))unique;
                var fullNameExistingType = GetFullName(searchExistingType);
                ambiguousTypes = new()
                {
                    { searchExistingType, fullNameExistingType }
                };
                UniqueNames[type.Name] = ambiguousTypes;
                if (!UniqueNames.ContainsKey(fullNameExistingType))
                {
                    UniqueNames.Add(fullNameExistingType, (searchExistingType, fullNameExistingType));
                }
            }

            var fullName = GetFullName(type);
            if (!UniqueNames.ContainsKey(fullName))
            {
                ambiguousTypes.Add(type, fullName);
                UniqueNames.Add(fullName, (type, fullName));
            }
            else
            {
                for (var i = 1; i < int.MaxValue; i++)
                {
                    var uniqueName = $"{fullName}{i}";
                    if (!UniqueNames.ContainsKey(uniqueName))
                    {
                        ambiguousTypes.Add(type, uniqueName);
                        UniqueNames.Add(uniqueName, (type, uniqueName));
                        break;
                    }
                }
            }
        }
        else
        {
            UniqueNames.Add(type.Name, (type, type.Name));
        }
    }

    //HINT: Need to be nullable string because of using struct types with default values, see:
    //https://learn.microsoft.com/en-us/dotnet/csharp/nullable-references#known-pitfalls
    internal string? GetUniqueTypeName(Type type)
    {
        return UniqueNames.TryGetValue(type.Name, out var unique) ?
            (unique is (Type, string name) ? name :
                (unique is Dictionary<Type, string> ambiguousTypes &&
                 ambiguousTypes.TryGetValue(type, out var searchExistingType) ?
                    searchExistingType : null))
            : null;
    }

    protected override void Dispose(bool disposing)
    {
        if (TypeWriters != null)
        {
            TypeWriters.ForEach(t => t.Dispose());
            TypeWriters.Clear();
            TypeWritersForComImport.Clear();
        }

        base.Dispose(disposing);
    }
}
