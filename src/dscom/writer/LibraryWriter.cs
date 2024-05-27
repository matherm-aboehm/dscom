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

using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace dSPACE.Runtime.InteropServices.Writer;

internal sealed class LibraryWriter : BaseWriter
{
    public LibraryWriter(Assembly assembly, WriterContext context) : base(context)
    {
        Assembly = assembly;
    }

    private Assembly Assembly { get; }

    private List<TypeWriter> TypeWriters { get; } = new();

    private Dictionary<string, object> UniqueNames { get; } = new();

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

            Context.TypeInfoResolver.AddTypeLib((ITypeLib)typeLib);

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

        List<TypeWriter> classInterfaceWriters = new();

        foreach (var type in types)
        {
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

            if (type.Namespace is null)
            {
                continue;
            }

            if (type.IsGenericType)
            {
                continue;
            }

            if (type.IsImport)
            {
                continue;
            }
            // Add this type to the unique names collection.
            UpdateUniqueNames(type);

            TypeWriter? typeWriter = null;
            if (type.IsInterface)
            {
                var interfaceTypeAttribute = type.GetCustomAttribute<InterfaceTypeAttribute>();

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
            else if (type.IsEnum)
            {
                typeWriter = new EnumWriter(type, this, Context);
            }
            else if (type.IsClass)
            {
                typeWriter = new ClassWriter(type, this, Context);
            }
            else if (type.IsValueType && !type.IsPrimitive)
            {
                typeWriter = new StructWriter(type, this, Context);
            }
            if (typeWriter != null)
            {

                TypeWriters.Add(typeWriter);
            }

            if (type.IsClass)
            {
                var createClassInterface = true;
                //check for class interfaces to generate:
                var classInterfaceType = type.GetCustomAttribute<ClassInterfaceAttribute>()?.Value ?? defaultClassInterfaceType;
                switch (classInterfaceType)
                {
                    case ClassInterfaceType.AutoDispatch:
                        createClassInterface = true;
                        break;
                    case ClassInterfaceType.AutoDual:
                        //CA1408: Do not use AutoDual ClassInterfaceType
                        //https://docs.microsoft.com/en-us/visualstudio/code-quality/ca1408?view=vs-2022
                        throw new NotSupportedException("Dual class interfaces not supported!");
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
                        var classInterfaceWriter = WriterFactory.CreateInstance(new ClassInterfaceWriter.FactoryArgs(classInterfaceType, type, this, Context));
                        classInterfaceWriters.Add(classInterfaceWriter);
                        classWriter.ClassInterfaceWriter = classInterfaceWriter;
                    }
                }
            }
        }
        TypeWriters.AddRange(classInterfaceWriters);
    }

    private static string GetFullNamespace(Type type)
    {
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
        return namesp.ToString();
    }

    private void UpdateUniqueNames(Type type)
    {
        if (UniqueNames.TryGetValue(type.Name, out var unique))
        {
            if (unique is not Dictionary<Type, string> ambiguousTypes)
            {
                var (searchExistingType, _) = ((Type, string))unique;
                ambiguousTypes = new()
                {
                    { searchExistingType, $"{GetFullNamespace(searchExistingType)}_{searchExistingType.Name}" }
                };
                UniqueNames[type.Name] = ambiguousTypes;
            }

            ambiguousTypes.Add(type, $"{GetFullNamespace(type)}_{type.Name}");
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
        }

        base.Dispose(disposing);
    }
}
