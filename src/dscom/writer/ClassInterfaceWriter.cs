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

using System.Reflection;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices.Writer;

internal sealed class ClassInterfaceWriter : DualInterfaceWriter
{
    internal new readonly record struct FactoryArgs(ClassInterfaceType ClassInterfaceType, Type SourceType, Type? ComImportType, LibraryWriter LibraryWriter, WriterContext Context)
        : WriterFactory.IWriterArgsFor<ClassInterfaceWriter>
    {
        ClassInterfaceWriter WriterFactory.IWriterArgsFor<ClassInterfaceWriter>.CreateInstance()
        {
            return new ClassInterfaceWriter(ClassInterfaceType, SourceType, ComImportType, LibraryWriter, Context);
        }
    }
    private static readonly Guid _ireflectGuid = new(Guids.GUID_IReflect);
    private static readonly Guid _iexpandoGuid = new(Guids.GUID_IExpando);

    private ClassInterfaceWriter(ClassInterfaceType classInterfaceType, Type sourceType, Type? comImportType, LibraryWriter libraryWriter, WriterContext context) : base(sourceType, libraryWriter, context)
    {
        TypeFlags = TYPEFLAGS.TYPEFLAG_FDUAL | TYPEFLAGS.TYPEFLAG_FDISPATCHABLE | TYPEFLAGS.TYPEFLAG_FOLEAUTOMATION | TYPEFLAGS.TYPEFLAG_FHIDDEN;
        var dispexInterfaces = sourceType.GetInterfaces().Where(
            static i => i.Equals(typeof(IReflect)) ||
            (i.FullName == typeof(IReflect).FullName && i.GUID == _ireflectGuid) ||
            (i.FullName == "System.Runtime.InteropServices.Expando.IExpando" &&
            i.GUID == _iexpandoGuid)).ToArray();
        // TODO: if type implements at least IReflect, it will be marshaled as IDispatchEx by runtime
        // with help of CCW, so do the same here for TLB?
        if (dispexInterfaces.Length >= 1)
        {
        }
        // if it implements IExpando then it also must implement IReflect, so count would be 2,
        // if not then it isn't extensible, but mark it only if it is also AutoDual.
        if (dispexInterfaces.Length != 2 &&
            classInterfaceType == ClassInterfaceType.AutoDual)
        {
            TypeFlags |= TYPEFLAGS.TYPEFLAG_FNONEXTENSIBLE;
        }
        ClassInterfaceType = classInterfaceType;
        ComDefaultInterface = sourceType.GetCustomAttribute<ComDefaultInterfaceAttribute>()?.Value;
        ComImportType = comImportType;
    }

    private Type? _comImportType;

    protected override string Name => ComImportType?.Name ?? $"_{base.Name!}";

    public ClassInterfaceType ClassInterfaceType { get; }

    public Type? ComDefaultInterface { get; }

    public Type? ComImportType
    {
        get => _comImportType;
        internal set
        {
            if (_comImportType is null && value is not null
                && value.GUID != Guid.Empty)
            {
                MarshalExtension.AddClassInterfaceTypeToCache(SourceType, value);
            }

            _comImportType = value;
        }
    }

    public override void Create()
    {
        Context.LogTypeExported($"Class interface '{Name}' exported.");

        if (ClassInterfaceType != ClassInterfaceType.AutoDual)
        {
            // skip method export when it's not AutoDual
            return;
        }

        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(ClassInterfaceWriter));
        }

        // Create all writer.
        var index = 0;
        var functionIndex = 0;
        foreach (var methodWriter in MethodWriters)
        {
            if (methodWriter.IsVisibleMethod)
            {
                methodWriter.FunctionIndex = functionIndex;
                methodWriter.VTableOffset = VTableOffsetUserMethodStart + (index * PtrSize);
                methodWriter.Create();
                functionIndex += methodWriter.IsValid ? 1 : 0;
            }
            index++;
        }

        TypeInfo.LayOut().ThrowIfFailed($"Failed to layout type {SourceType}.");
    }

    protected override void CreateMethodWriters()
    {
        //TODO: re-implement this to include fake method writers for class fields.
        //see: ComMTMemberInfoMap::SetupPropsForIClassX from CoreCLR\src\vm\commtmemberinfomap.cpp
        base.CreateMethodWriters();
    }

    protected override Guid GetTypeGuid()
    {
        return MarshalExtension.GetClassInterfaceGuidForType(SourceType, this);
    }

    /// <summary>
    /// Return 0 as major version. Even if this is questionable, tlbexp behaves like this.
    /// </summary>
    protected override ushort MajorVersion => 0;
}
