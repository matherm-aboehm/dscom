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
using System.Globalization;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices.Writer;

internal abstract class InterfaceWriter : TypeWriter, WriterFactory.IProvidesFinishCreateInstance
{
    protected InterfaceWriter(Type sourceType, LibraryWriter libraryWriter, WriterContext context) : base(sourceType, libraryWriter, context)
    {
        VTableOffsetUserMethodStart = 7 * PtrSize;
        DispatchIdCreator = new DispatchIdCreator(this);
    }

    protected virtual void FinishCreateInstance()
    {
        // Base types from interfaces should be pre-defined in stdole2 and not custom types.
        // Custom types can only be resolved after TypeWriter.CreateTypeInfo() was called,
        // but CreateMethodWriters() needs type info here for duplicate name checking.
        BaseTypeInfo = Context.TypeInfoResolver.ResolveTypeInfo(BaseInterfaceGuid);
        Debug.Assert(BaseTypeInfo is not null);

        CreateMethodWriters();

        // Check all DispIDs
        // Handle special IDs like 0 or -4, and try to fix duplicate DispIds if possible.
        DispatchIdCreator.NormalizeIds();
    }

    void WriterFactory.IProvidesFinishCreateInstance.FinishCreateInstance()
    {
        FinishCreateInstance();
    }

    public DispatchIdCreator DispatchIdCreator { get; protected set; }

    public int VTableOffsetUserMethodStart { get; set; }

    public ComInterfaceType ComInterfaceType { get; set; }

    public FUNCKIND FuncKind { get; protected set; }

    public abstract Guid BaseInterfaceGuid { get; }

    protected internal List<MethodWriter> MethodWriters { get; } = new();

    private ITypeInfo? BaseTypeInfo { get; set; }

    /// <summary>
    /// Gets or Sets a value indicating whether all method should return HRESULT as return value.
    /// </summary>
    internal bool UseHResultAsReturnValue { get; set; }

    public override void CreateTypeInheritance()
    {
        if (BaseTypeInfo != null)
        {
            TypeInfo.AddRefTypeInfo(BaseTypeInfo, out var phRefType)
                .ThrowIfFailed($"Failed to add IDispatch reference to {SourceType}.");
            TypeInfo.AddImplType(0, phRefType)
                .ThrowIfFailed($"Failed to add IDispatch implementation type to {SourceType}.");
        }
    }

    public override void Create()
    {
        Context.LogTypeExported($"Interface '{Name}' exported.");

        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(InterfaceWriter));
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

            // Increment the index for the VTableOffset
            index++;
        }

        TypeInfo.LayOut().ThrowIfFailed($"Failed to layout type {SourceType}.");
    }

    protected virtual void CreateMethodWriters()
    {
        var methods = SourceType.GetMethods().ToList();
        methods.Sort((a, b) => a.MetadataToken - b.MetadataToken);

        foreach (var method in methods)
        {
            var methodName = method.Name;
            if (methodName == ".ctor")
            {
                //HINT: Should never reach here, GetMethods() excludes ConstructorInfos,
                // which is the same the CLR does even for CCWs.
                // But for the case someone defined an ordinary method with this name,
                // this conversion is required in order to be compatible with CLR implementation.
                methodName = "Init"; // COM name for constructor
            }
            //HINT: don't replace method.Name with methodName in next code line,
            // this is intentionally, because first compare is on original name.
            var numIdenticalNames = MethodWriters.Count(z => z.IsVisibleMethod && (z.MethodInfo.Name == method.Name || z.MethodName.StartsWith(methodName + "_", StringComparison.Ordinal)));

            numIdenticalNames += GetMethodNamesOfBaseTypeInfo(BaseTypeInfo).Count(z => z == methodName || z.StartsWith(methodName + "_", StringComparison.Ordinal));

            var alternateName = numIdenticalNames == 0 ? methodName : methodName + "_" + (numIdenticalNames + 1).ToString(CultureInfo.InvariantCulture);
            MethodWriter methodWriter;
            if ((methodName.StartsWith("get_", StringComparison.Ordinal) || methodName.StartsWith("set_", StringComparison.Ordinal)) && method.IsSpecialName)
            {
                // Use ReflectedType instead of DeclaringType, because when method was only defined on base type,
                // ReflectedType of all property accessors would be different to current method when retrieved from DeclaringType
                // and so Equals would always return false which results in a Exception for the First() call.
                var propertyInfo = method.ReflectedType!.GetProperties().FirstOrDefault(p => p.GetGetMethod() == method || p.GetSetMethod() == method);
                if (propertyInfo is null)
                {
                    // If method definition is from base type
                    if (!method.ReflectedType!.Equals(method.DeclaringType))
                    {
                        var othermethod = methods.Single(m => m.Name == method.Name && m != method);
                        // If there is another method with the same name that belongs to a property,
                        // it isn't an overload, but a property defined with "new" keyword, to hide
                        // the base property, in this case, its base definition is pointing to itself
                        // see: https://stackoverflow.com/questions/288357/how-does-reflection-tell-me-when-a-property-is-hiding-an-inherited-member-with-t
                        if (othermethod != null && othermethod.IsSpecialName && !othermethod.DeclaringType!.Equals(method.DeclaringType)
                            // GetBaseDefinition can't be used in reflection-only context, so check only
                            // DeclaringType of both methods
                            /*othermethod.GetBaseDefinition() != method.GetBaseDefinition()*/)
                        {
                            // Skip this method, because reflection doesn't list a duplicate property,
                            // which belongs to this method, only the method is listed twice.
                            continue;
                        }
                    }
                    throw new InvalidOperationException($"Can't find property that belongs to the method {method.Name}");
                }
                alternateName = alternateName.Substring(4);
                if (methodName.StartsWith("get_", StringComparison.Ordinal))
                {
                    methodWriter = WriterFactory.CreateInstance(new PropertyGetMethodWriter.FactoryArgs(this, propertyInfo, method, Context, alternateName));
                }
                else
                {
                    Debug.Assert(methodName.StartsWith("set_", StringComparison.Ordinal));
                    methodWriter = WriterFactory.CreateInstance(new PropertySetMethodWriter.FactoryArgs(this, propertyInfo, method, Context, alternateName));
                }
            }
            else if (methodName == nameof(ToString))
            {
                // use same behavior as in .NET Framework, so convert this method to property getter
                methodWriter = WriterFactory.CreateInstance(new PropertyGetMethodWriter.FactoryArgs(this, null, method, Context, alternateName));
            }
            else
            {
                methodWriter = WriterFactory.CreateInstance(new MethodWriter.FactoryArgs(this, method, Context, alternateName));
            }

            MethodWriters.Add(methodWriter);
            if (methodWriter.IsVisibleMethod)
            {
                DispatchIdCreator!.RegisterMember(methodWriter);
            }
            else
            {
                // In case of ComVisible false, we need to increment the the next free DispId
                // This offset is needed to keep the DispId in sync with the VTable and the behavior of tlbexp.
                DispatchIdCreator!.GetNextFreeDispId();
            }
        }
    }

    private static IEnumerable<string> GetMethodNamesOfBaseTypeInfo(ITypeInfo? typeInfo)
    {
        var names = new List<string>();
        if (typeInfo != null)
        {
            var ppTypeAttr = new IntPtr();
            try
            {
                typeInfo.GetTypeAttr(out ppTypeAttr);
                var typeAttr = Marshal.PtrToStructure<TYPEATTR>(ppTypeAttr);

                for (var i = 0; i < typeAttr.cFuncs; i++)
                {
                    typeInfo.GetFuncDesc(i, out var ppFuncDesc);
                    var funcDesc = Marshal.PtrToStructure<FUNCDESC>(ppFuncDesc);
                    typeInfo.GetDocumentation(funcDesc.memid, out var name, out _, out _, out _);
                    names.Add(name);
                }

                for (var i = 0; i < typeAttr.cImplTypes; i++)
                {
                    typeInfo.GetRefTypeOfImplType(i, out var href);
                    typeInfo.GetRefTypeInfo(href, out var refTypeInfo);
                    names.AddRange(GetMethodNamesOfBaseTypeInfo(refTypeInfo));
                }
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(ppTypeAttr);
            }
        }

        return names;
    }


    protected override void Dispose(bool disposing)
    {
        if (MethodWriters != null)
        {
            MethodWriters.ForEach(t => t.Dispose());
            MethodWriters.Clear();
        }

        base.Dispose(disposing);
    }
}
