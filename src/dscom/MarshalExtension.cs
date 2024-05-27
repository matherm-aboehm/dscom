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

using System.Collections.Concurrent;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Provides extension methods for marshaling data between managed and unmanaged code.
/// </summary>
internal static class MarshalExtension
{
    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid COMPLUS_RUNTIME_GUID_FULLFRAMEWORK = new("c9cbf969-05da-d111-9408-0000f8083460");

    [SuppressMessage("Microsoft.Style", "IDE1006", Justification = "")]
    private static readonly Guid COMPLUS_RUNTIME_GUID_CORE = new("69f9cbc9-da05-11d1-9408-0000f8083460");

    private static readonly ConcurrentDictionary<Type, Guid> _cacheTypeGuids = new();
    private static readonly ConcurrentDictionary<Type, Guid> _cacheClassIntfGuids = new();

    internal static void ClearCaches()
    {
        _cacheTypeGuids.Clear();
        _cacheClassIntfGuids.Clear();
    }

    internal static Guid GetClassInterfaceGuidForType(Type type, Writer.ClassInterfaceWriter? writer = null)
    {
        if (writer is null)
        {
            return _cacheClassIntfGuids.TryGetValue(type, out var guidFromCache) ? guidFromCache : Guid.Empty;
        }
        var guid = _cacheClassIntfGuids.GetOrAdd(type, (t) =>
        {
            var nsGuid = GetGuidNamespaceFromAssembly(t.Assembly);
            var rv = GetStringizedClassItfDef(t, writer);
            return GuidFromName(nsGuid, rv);
        });
        return guid;
    }

    internal static Guid GenerateGuidForType(Type type)
    {
        if (type.GUID != Guid.Empty)
        {
            return type.GUID;
        }
#if !NETSTANDARD2_0
        if (!type.Assembly.ReflectionOnly)
        {
            // works only with runtime types
            return Marshal.GenerateGuidForType(type);
        }
#endif
        static Guid createNewGuidForType(Type type)
        {
            var nsGuid = GetGuidNamespaceFromAssembly(type.Assembly);
            byte[] rv;
            if (type.IsInterface)
            {
                rv = GetStringizedItfDef(type);
            }
            else
            {
                var name = GetFullyQualifiedNameForClassNestedAware(type);
                var nameBytes = Encoding.Unicode.GetBytes(name);
                var assemblyNameBytes = GetStringizedTypeLibGuidForAssembly(type.Assembly);

                Debug.Assert((nameBytes.Length + assemblyNameBytes.Length) % UnicodeEncoding.CharSize == 0);
                rv = new byte[nameBytes.Length + assemblyNameBytes.Length];
                Buffer.BlockCopy(nameBytes, 0, rv, 0, nameBytes.Length);
                Buffer.BlockCopy(assemblyNameBytes, 0, rv, nameBytes.Length, assemblyNameBytes.Length);
            }
            return GuidFromName(nsGuid, rv);
        }

        var guid = _cacheTypeGuids.GetOrAdd(type, createNewGuidForType);
        return guid;
    }

    private static string GetFullyQualifiedNameForClassNestedAware(Type type)
    {
        StringBuilder builder = new();
        if (type.IsArray)
        {
            return builder.ToString();
        }
        List<string> nameParts = new();
        while (type.IsNested)
        {
            nameParts.Add(type.Name);
            nameParts.Add("+");
            type = type.DeclaringType!;
        }
        if (type.Namespace is null)
        {
            return builder.ToString();
        }
        nameParts.Add(type.Name);
        nameParts.Add(".");
        nameParts.Add(type.Namespace);
        foreach (var part in nameParts.Reverse<string>())
        {
            builder.Append(part);
        }
        return builder.ToString();
    }

    /// <summary>
    /// Implementation of name-based algorithm from RFC 4122.
    /// </summary>
    /// <param name="nameBytes">Name as canonical sequence of octets.</param>
    /// <param name="nsGuid">Namespace</param>
    /// <returns>GUID based on name and namespace.</returns>
    private static Guid GuidFromName(Guid nsGuid, byte[] nameBytes)
    {
        static void SwapBytes(byte[] bytes, int ileft, int iright)
        {
            (bytes[iright], bytes[ileft]) = (bytes[ileft], bytes[iright]);
        }

        static void SwapByteOrder(byte[] guid)
        {
            if (BitConverter.IsLittleEndian)
            {
                //1st component
                SwapBytes(guid, 0, 3);
                SwapBytes(guid, 1, 2);
                //2nd component
                SwapBytes(guid, 4, 5);
                //3rd component
                SwapBytes(guid, 6, 7);
                //other components already in correct order
            }
        }
        var nsBytes = nsGuid.ToByteArray();
        // convert guid byte array to network order (big endian)
        SwapByteOrder(nsBytes);

        var rv = new byte[nsBytes.Length + nameBytes.Length];
        Buffer.BlockCopy(nsBytes, 0, rv, 0, nsBytes.Length);
        Buffer.BlockCopy(nameBytes, 0, rv, nsBytes.Length, nameBytes.Length);

        using var md5 = MD5.Create();
        var hash = md5.ComputeHash(rv);
        hash[6] = (byte)((hash[6] & 0x0f) | 0x30); //version = 3
        hash[8] = (byte)((hash[8] & 0x3f) | 0x80); //variant = 2
        // convert from network order back to guid byte array
        SwapByteOrder(hash);

        return new Guid(hash);
    }

    internal static Guid GetTypeLibGuidForAssembly(Assembly assembly)
    {
        Guid guid;
        GuidAttribute? guidAttribute = null;
        try
        {
            guidAttribute = assembly.GetCustomAttribute<GuidAttribute>();
        }
        catch (FileNotFoundException)
        {
        }

        if (guidAttribute != null)
        {
            guid = new Guid(guidAttribute.Value);
        }
        else
        {
            var stringGuid = GetStringizedTypeLibGuidForAssembly(assembly);
            var nsGuid = GetGuidNamespaceFromAssembly(assembly);
            guid = GuidFromName(nsGuid, stringGuid);
        }
        return guid;
    }

    internal static bool IsTypeVisibleFromCom(Type type)
    {
#if !NETSTANDARD2_0
        if (!type.Assembly.ReflectionOnly)
        {
            // works only with runtime types
            return Marshal.IsTypeVisibleFromCom(type);
        }
#endif
        if ((type.IsGenericType || type.IsGenericParameter) &&
            !type.SupportsGenericInterop(WinRTExtensions.InteropKind.NativeToManaged))
        {
            return false;
        }
        if ((type.IsInterface && type.IsImport) || type.IsProjectedFromWinRT())
        {
            return true;
        }
        if (type.IsArray)
        {
            return false;
        }
        if (!type.IsPublic && !type.IsNestedPublicRecursive())
        {
            return false;
        }
        var comVisibleAttribute = type.GetCustomAttribute<ComVisibleAttribute>(false) ??
            type.Assembly.GetCustomAttribute<ComVisibleAttribute>();

        return comVisibleAttribute?.Value ?? true;
    }

    internal static bool IsMemberVisibleFromCom(Type declaringType, MemberInfo info, MethodInfo? miAssociate = null)
    {
        bool? fromWinRT = null;
        switch (info)
        {
            case FieldInfo fi:
                Debug.Assert(miAssociate == null);
                if (!fi.IsPublic)
                {
                    return false;
                }
                break;
            case MethodInfo mi:
                Debug.Assert(miAssociate == null);
                if (!mi.IsPublic || mi.IsGenericMethod)
                {
                    return false;
                }
                break;
            case PropertyInfo:
                Debug.Assert(miAssociate != null);
                if (!miAssociate.IsPublic)
                {
                    return false;
                }
                fromWinRT = declaringType.IsProjectedFromWinRT() || declaringType.IsExportedToWinRT() || declaringType.IsWinRTObjectType();
                if (!fromWinRT.Value)
                {
                    var comVisibleAttribute = miAssociate.GetCustomAttribute<ComVisibleAttribute>(false);
                    if (comVisibleAttribute != null)
                    {
                        return comVisibleAttribute.Value;
                    }
                }
                break;
        }
        if (!(fromWinRT ?? (declaringType.IsProjectedFromWinRT() || declaringType.IsExportedToWinRT() || declaringType.IsWinRTObjectType())))
        {
            var comVisibleAttribute = info.GetCustomAttribute<ComVisibleAttribute>(false);
            if (comVisibleAttribute != null)
            {
                return comVisibleAttribute.Value;
            }
        }
        return true;
    }

    internal static string GenerateProgIdForType(Type type)
    {
        if (type is null)
        {
            throw new ArgumentNullException(nameof(type));
        }
        if (type.IsImport)
        {
            throw new ArgumentException("The type must not be ComImport.", nameof(type));
        }
        if (type.IsGenericType)
        {
            throw new ArgumentException("The type needs to be non-generic.", nameof(type));
        }
        if (!RegistrationServices.TypeRequiresRegistrationHelper(type))
        {
            throw new ArgumentException("The type must be com creatable.", nameof(type));
        }
        IList<CustomAttributeData> customAttributes = CustomAttributeData.GetCustomAttributes(type);
        for (int i = 0; i < customAttributes.Count; i++)
        {
            if (customAttributes[i].Constructor.DeclaringType == typeof(ProgIdAttribute))
            {
                IList<CustomAttributeTypedArgument> constructorArguments = customAttributes[i].ConstructorArguments;
                string progId = (string?)constructorArguments[0].Value ?? string.Empty;
                return progId;
            }
        }
        return type.FullName!;
    }

    private static byte[] StructureToByteArray(object obj)
    {
        var len = Marshal.SizeOf(obj);
        var arr = new byte[len];
        var ptr = Marshal.AllocHGlobal(len);
        Marshal.StructureToPtr(obj, ptr, true);
        Marshal.Copy(ptr, arr, 0, len);
        Marshal.FreeHGlobal(ptr);
        return arr;
    }

    private static Guid GetGuidNamespaceFromAssembly(Assembly assembly)
    {
        var nsGuid = COMPLUS_RUNTIME_GUID_CORE;
        var targetFrameworkAttribute = assembly.GetCustomAttribute<TargetFrameworkAttribute>();
        if (targetFrameworkAttribute != null && targetFrameworkAttribute.FrameworkName.StartsWith(".NETFramework", StringComparison.Ordinal))
        {
            nsGuid = COMPLUS_RUNTIME_GUID_FULLFRAMEWORK;
        }
        return nsGuid;
    }

    private static byte[] GetStringizedTypeLibGuidForAssembly(Assembly assembly)
    {
        const string typelibKeyName = "TypeLib";
        var assemblyName = assembly.GetName();
        var name = assemblyName.Name;
        var publicKeyBytes = assemblyName.GetPublicKey();
        ushort majorVersion = 0;
        ushort minorVersion = 0;
        ushort buildNumber = 0;
        ushort revisionNumber = 0;

        name ??= "";
        name = name.Replace('.', '_');
        name = name.Replace(' ', '_');
        name = name.ToLower(CultureInfo.CurrentCulture);

        ComCompatibleVersionAttribute? versionAttr = null;
        try
        {
            versionAttr = assembly.GetCustomAttribute<ComCompatibleVersionAttribute>();
        }
        catch (FileNotFoundException)
        {

        }

        static ushort GetVersionPartFromInt(int part)
        {
            return (part is < 0 or > ushort.MaxValue) ? (ushort)0 : (ushort)part;
        }

        if (versionAttr != null)
        {
            majorVersion = GetVersionPartFromInt(versionAttr.MajorVersion);
            minorVersion = GetVersionPartFromInt(versionAttr.MinorVersion);
            buildNumber = GetVersionPartFromInt(versionAttr.BuildNumber);
            revisionNumber = GetVersionPartFromInt(versionAttr.RevisionNumber);
        }
        else
        {
            var version = assemblyName.Version;

            if (version != null)
            {
                majorVersion = (ushort)version.Major;
                minorVersion = (ushort)version.Minor;
                buildNumber = (ushort)version.Build;
                revisionNumber = (ushort)version.Revision;
            }
        }

        //HINT: minor = major is intentionally, don't fix it
        var versionInfo = new Versioninfo() { MajorVersion = majorVersion, MinorVersion = majorVersion, BuildNumber = buildNumber, RevisionNumber = revisionNumber };

        var nameBytes = Encoding.Unicode.GetBytes(name);
        var typelibBytes = Encoding.ASCII.GetBytes(typelibKeyName);
        var versionBytes = StructureToByteArray(versionInfo);
        var returnLength = nameBytes.Length + typelibBytes.Length + versionBytes.Length;
        byte[]? minorBytes = null;

        if (minorVersion != 0)
        {
            minorBytes = BitConverter.GetBytes(minorVersion);
            returnLength += minorBytes.Length;
        }

        if (publicKeyBytes != null && publicKeyBytes.Length != 0)
        {
            returnLength += publicKeyBytes.Length;
        }

        // pad to a whole WCHAR
        if (returnLength % UnicodeEncoding.CharSize != 0)
        {
            returnLength += UnicodeEncoding.CharSize - (returnLength % UnicodeEncoding.CharSize);
        }

        var rv = new byte[returnLength];

        var currentStart = 0;
        Buffer.BlockCopy(nameBytes, 0, rv, currentStart, nameBytes.Length);
        currentStart += nameBytes.Length;
        Buffer.BlockCopy(typelibBytes, 0, rv, currentStart, typelibBytes.Length);
        currentStart += typelibBytes.Length;
        Buffer.BlockCopy(versionBytes, 0, rv, currentStart, versionBytes.Length);
        currentStart += versionBytes.Length;
        if (minorBytes != null)
        {
            Buffer.BlockCopy(minorBytes, 0, rv, currentStart, minorBytes.Length);
            currentStart += minorBytes.Length;
        }
        if (publicKeyBytes != null && publicKeyBytes.Length != 0)
        {
            Buffer.BlockCopy(publicKeyBytes, 0, rv, currentStart, publicKeyBytes.Length);
        }

        return rv;
    }

    enum DefaultInterfaceType
    {
        Explicit = 0,
        IUnknown = 1,
        AutoDual = 2,
        AutoDispatch = 3,
        BaseComClass = 4
    };
    private static byte[] GetStringizedClassItfDef(Type classType, Writer.ClassInterfaceWriter writer)
    {
        Debug.Assert(!classType.IsInterface);
        using MemoryStream rv = new();
        using BinaryWriter rvwriter = new(rv, Encoding.Unicode, leaveOpen: true);

        var bGenerateMethods = false;
        var defItfFlags = GetDefaultInterfaceForClassWrapper(classType, writer, out var defItfType);

        if (defItfType == classType && defItfFlags == DefaultInterfaceType.AutoDual)
        {
            //should not reach here, because AutoDual is filtered out from LibraryWriter
            bGenerateMethods = true;
        }

        var name = GetFullyQualifiedNameForClassNestedAware(classType);
        rvwriter.Write(name.ToCharArray());

        if (bGenerateMethods)
        {
            //TODO: use same semantics as in Framework CLR/CoreCLR
            // but for now just use implementation from InterfaceWriter
            foreach (var mwriter in writer.MethodWriters)
            {
                if (mwriter.IsVisibleMethod)
                {
                    if (mwriter.MemberInfo is FieldInfo fi)
                    {
                        GetStringizedFieldDef(classType, fi, rv);
                    }
                    else
                    {
                        GetStringizedMethodDef(classType, mwriter.MethodInfo, rv);
                    }
                }
            }
        }

        // pad to a whole WCHAR
        if (rv.Length % UnicodeEncoding.CharSize != 0)
        {
            rv.SetLength(rv.Length + UnicodeEncoding.CharSize - (rv.Length % UnicodeEncoding.CharSize));
        }
        return rv.ToArray();
    }

    private static DefaultInterfaceType GetDefaultInterfaceForClassWrapper(Type classType, Writer.ClassInterfaceWriter? writer, out Type? defItfType)
    {
        defItfType = null;
        if (classType.IsWinRTObjectType() || classType.IsExportedToWinRT())
        {
            return DefaultInterfaceType.IUnknown;
        }

        var classInterfaceType = classType.IsImport ? ClassInterfaceType.None
            : writer?.ClassInterfaceType ?? classType.GetComClassInterfaceType();
        // no need to check for com visibility when writer instance is specified,
        // it was already checked by LibraryWriter, if there is no writer instance,
        // it means it has to check again.
        var bComVisible = writer != null || classType.IsImport || IsTypeVisibleFromCom(classType);
        if (!bComVisible)
        {
            return DefaultInterfaceType.IUnknown;
        }

        Type? defaultInterfaceType = writer != null ?
            writer.ComDefaultInterface :
            classType.GetCustomAttribute<ComDefaultInterfaceAttribute>()?.Value;
        if (defaultInterfaceType != null)
        {
            Debug.Assert(defaultInterfaceType.IsInterface);
            Debug.Assert(defaultInterfaceType.Assembly.ReflectionOnly == classType.Assembly.ReflectionOnly);
            if (!defaultInterfaceType.IsAssignableFrom(classType))
            {
                throw new TypeLoadException($"The class {classType} doesn't implement {defaultInterfaceType}");
            }
            defItfType = defaultInterfaceType;
            return DefaultInterfaceType.Explicit;
        }

        if (classInterfaceType != ClassInterfaceType.None)
        {
            defItfType = classType;
            return classInterfaceType == ClassInterfaceType.AutoDispatch ?
                DefaultInterfaceType.AutoDispatch : DefaultInterfaceType.AutoDual;
        }

        for (var baseType = classType.BaseType; baseType != null; baseType = baseType.BaseType)
        {
            foreach (var itf in classType.GetInterfaces())
            {
                if (!itf.IsGenericType)
                {
                    if (IsTypeVisibleFromCom(itf) && !itf.IsAssignableFrom(baseType))
                    {
                        defItfType = itf;
                        return DefaultInterfaceType.Explicit;
                    }
                }
            }
        }

        if (classType.IsImport)
        {
            return DefaultInterfaceType.IUnknown;
        }

        var parentType = classType.GetComPlusParentType();
        if (parentType != null)
        {
            return GetDefaultInterfaceForClassWrapper(parentType, null, out defItfType);
        }

        if (classType.IsCOMObject)
        {
            return DefaultInterfaceType.BaseComClass;
        }
        return DefaultInterfaceType.IUnknown;
    }

    private static byte[] GetStringizedItfDef(Type interfaceType)
    {
        using MemoryStream rv = new();
        using BinaryWriter writer = new(rv, Encoding.Unicode, leaveOpen: true);
        var name = GetFullyQualifiedNameForClassNestedAware(interfaceType);
        writer.Write(name.ToCharArray());

        foreach (var mi in interfaceType.GetMethods())
        {
            GetStringizedMethodDef(interfaceType, mi, rv);
        }
        foreach (var fi in interfaceType.GetFields())
        {
            GetStringizedFieldDef(interfaceType, fi, rv);
        }
        // pad to a whole WCHAR
        if (rv.Length % UnicodeEncoding.CharSize != 0)
        {
            rv.SetLength(rv.Length + UnicodeEncoding.CharSize - (rv.Length % UnicodeEncoding.CharSize));
        }
        return rv.ToArray();
    }

    private static void GetStringizedMethodDef(Type declaringType, MethodInfo mi, Stream outStream)
    {
        if (!IsMemberVisibleFromCom(declaringType, mi))
        {
            return;
        }
        var sigBytes = declaringType.Module.ResolveSignature(mi.MetadataToken);
        var sigBlob = Extensions.GetBlobReaderFromByteArray(sigBytes, out var sigBlobContext);
        string sigString;
        using (sigBlobContext)
        {
            sigString = PrettyPrintSig(declaringType, sigBlob);
        }
        var sigStringBytes = Encoding.Default.GetBytes(sigString);
        outStream.Write(sigStringBytes, 0, sigStringBytes.Length);
        foreach (var pi in mi.GetParameters())
        {
            outStream.WriteByte((byte)pi.Attributes);
        }
    }

    private static void GetStringizedFieldDef(Type declaringType, FieldInfo fi, Stream outStream)
    {
        if (!IsMemberVisibleFromCom(declaringType, fi))
        {
            return;
        }
        var sigBytes = declaringType.Module.ResolveSignature(fi.MetadataToken);
        var sigBlob = Extensions.GetBlobReaderFromByteArray(sigBytes, out var sigBlobContext);
        string sigString;
        using (sigBlobContext)
        {
            sigString = PrettyPrintSig(declaringType, sigBlob);
        }
        var sigStringBytes = Encoding.Default.GetBytes(sigString);
        outStream.Write(sigStringBytes, 0, sigStringBytes.Length);
    }

    private class SignatureTypeProviderForComInterop : ISignatureTypeProvider<string, Type>
    {
        public static SignatureTypeProviderForComInterop Instance { get; }
            = new SignatureTypeProviderForComInterop();

        public string GetArrayType(string elementType, ArrayShape shape)
        {
            StringBuilder result = new();
            result.Append(elementType);
            if (shape.Rank == 0)
            {
                result.Append("[??]");
            }
            else
            {
                result.Append('[');
                Debug.Assert(shape.Sizes.Length == shape.Rank &&
                    shape.LowerBounds.Length == shape.Rank);
                for (int i = 0; i < shape.Rank; i++)
                {
                    //maybe should use "||" not "&&", but original PrettyPrintTypeA
                    //has this bug, so do the same here
                    if (shape.Sizes[i] != 0 && shape.LowerBounds[i] != 0)
                    {
                        if (shape.LowerBounds[i] == 0)
                        {
                            result.Append(shape.Sizes[i]);
                        }
                        else
                        {
                            result.Append(shape.LowerBounds[i]);
                            result.Append("...");
                            if (shape.Sizes[i] != 0)
                            {
                                result.Append(shape.LowerBounds[i] + shape.Sizes[i] + 1);
                            }
                        }
                    }
                    if (i < (shape.Rank - 1))
                    {
                        result.Append(',');
                    }
                }
                result.Append(']');
            }
            return result.ToString();
        }

        public string GetByReferenceType(string elementType)
        {
            return $"{elementType}&";
        }

        public string GetFunctionPointerType(MethodSignature<string> signature)
        {
            StringBuilder result = new();
            var callConv = signature.Header.RawValue;
            if (signature.Header.IsInstance)
            {
                result.Append("instance ");
            }
            if (signature.Header.IsGeneric)
            {
                result.Append("generic ");
            }
            if ((callConv & SignatureHeader.CallingConventionOrKindMask) < IMAGE_CEE_CS_CALLCONV_MAX)
            {
                result.Append(_callConvNames[callConv & SignatureHeader.CallingConventionOrKindMask]);
            }
            result.Append(signature.ReturnType);
            var requiredParameterCount = signature.RequiredParameterCount;
            result.Append('(');
            foreach (var (typeString, i) in signature.ParameterTypes.Select((t, i) => (t, i)))
            {
                if (i > 0)
                {
                    result.Append(',');
                }
                if (i == requiredParameterCount)
                {
                    result.Append("...");
                    result.Append(',');
                }
                result.Append(typeString);
            }
            result.Append(')');
            return result.ToString();
        }

        public string GetGenericInstantiation(string genericType, ImmutableArray<string> typeArguments)
        {
            StringBuilder result = new();
            result.Append(genericType);
            result.Append('<');
            foreach (var (typeString, i) in typeArguments.Select((t, i) => (t, i)))
            {
                if (i > 0)
                {
                    result.Append(',');
                }
                result.Append(typeString);
            }
            result.Append('>');
            return result.ToString();
        }

        public string GetGenericMethodParameter(Type genericContext, int index)
        {
            return $"!!{index}";
        }

        public string GetGenericTypeParameter(Type genericContext, int index)
        {
            return $"!{index}";
        }

        public string GetModifiedType(string modifier, string unmodifiedType, bool isRequired)
        {
            StringBuilder result = new();
            result.Append(isRequired ? "required_modifier " : "optional_modifier ");
            result.Append(modifier);
            if (!string.IsNullOrEmpty(unmodifiedType))
            {
                //SignatureDecoder continues decoding with DecodeType into "unmodifiedType"
                //after decoding with DecodeTypeHandle into "modifier",
                //but original PrettyPrintTypeA doesn't do this and just returns,
                //so append as it would be another parameter.
                result.Append(',');
                result.Append(unmodifiedType);
            }
            return result.ToString();
        }

        public string GetPinnedType(string elementType)
        {
            return $"{elementType} pinned";
        }

        public string GetPointerType(string elementType)
        {
            return $"{elementType}*";
        }

        public string GetPrimitiveType(PrimitiveTypeCode typeCode)
        {
            return typeCode switch
            {
                PrimitiveTypeCode.Void => "void",
                PrimitiveTypeCode.Boolean => "bool",
                PrimitiveTypeCode.Char => "wchar",
                PrimitiveTypeCode.SByte => "int8",
                PrimitiveTypeCode.Byte => "unsigned int8",
                PrimitiveTypeCode.Int16 => "int16",
                PrimitiveTypeCode.UInt16 => "unsigned int16",
                PrimitiveTypeCode.Int32 => "int32",
                PrimitiveTypeCode.UInt32 => "unsigned int32",
                PrimitiveTypeCode.Int64 => "int64",
                PrimitiveTypeCode.UInt64 => "unsigned int64",
                PrimitiveTypeCode.Single => "float32",
                PrimitiveTypeCode.Double => "float64",
                PrimitiveTypeCode.UIntPtr => "unsigned int",
                PrimitiveTypeCode.IntPtr => "int",
                PrimitiveTypeCode.Object => "class System.Object",
                PrimitiveTypeCode.String => "class System.String",
                PrimitiveTypeCode.TypedReference => "refany",
                _ => throw new InvalidOperationException($"invalid typecode {typeCode}")
            };
        }

        public string GetSZArrayType(string elementType)
        {
            return $"{elementType}[]";
        }

        private static void AppendSigTypeKind(StringBuilder builder, SignatureTypeKind sigTypeKind)
        {
            switch (sigTypeKind)
            {
                case SignatureTypeKind.ValueType: builder.Append("value class "); break;
                case SignatureTypeKind.Class: builder.Append("class "); break;
                case SignatureTypeKind.Unknown: break;
                default: throw new InvalidOperationException($"invalid typekind {sigTypeKind}");
            }
        }

        public string GetTypeFromDefinition(MetadataReader reader, TypeDefinitionHandle handle, byte rawTypeKind)
        {
            StringBuilder result = new();
            AppendSigTypeKind(result, (SignatureTypeKind)rawTypeKind);
            var typeDef = reader.GetTypeDefinition(handle);
            const string errorName = "Invalid TypeDef record";
            var ns = typeDef.Namespace.IsNil ? errorName : reader.GetString(typeDef.Namespace);
            var name = typeDef.Name.IsNil ? errorName : reader.GetString(typeDef.Name);
            if (!string.IsNullOrEmpty(ns))
            {
                result.Append(ns);
                result.Append('.');
            }
            result.Append(name);
            return result.ToString();
        }

        public string GetTypeFromReference(MetadataReader reader, TypeReferenceHandle handle, byte rawTypeKind)
        {
            StringBuilder result = new();
            AppendSigTypeKind(result, (SignatureTypeKind)rawTypeKind);
            var typeRef = reader.GetTypeReference(handle);
            const string errorName = "Invalid TypeRef record";
            var ns = typeRef.Namespace.IsNil ? errorName : reader.GetString(typeRef.Namespace);
            var name = typeRef.Name.IsNil ? errorName : reader.GetString(typeRef.Name);
            if (!string.IsNullOrEmpty(ns))
            {
                result.Append(ns);
                result.Append('.');
            }
            result.Append(name);
            return result.ToString();
        }

        public string GetTypeFromSpecification(MetadataReader reader, Type genericContext, TypeSpecificationHandle handle, byte rawTypeKind)
        {
            StringBuilder result = new();
            AppendSigTypeKind(result, (SignatureTypeKind)rawTypeKind);
            var typeSpec = reader.GetTypeSpecification(handle);
            result.Append(typeSpec.DecodeSignature(this, genericContext));
            return result.ToString();
        }
    }

    const byte IMAGE_CEE_CS_CALLCONV_MAX = 12;
    static readonly string[] _callConvNames =
            new string[IMAGE_CEE_CS_CALLCONV_MAX]
            {
            "",
            "unmanaged cdecl ",
            "unmanaged stdcall ",
            "unmanaged thiscall ",
            "unmanaged fastcall ",
            "vararg ",
            "<error> ",
            "<error> ",
            "",
            "",
            "",
            "native vararg "
            };

    private static string PrettyPrintSig(Type declaringType, BlobReader sigBlob, string? name = "")
    {
        StringBuilder sigString = new();
        ImmutableArray<string> types;
        int requiredParameterCount;
        var numTyArgs = 0;
        var reader = declaringType.Assembly.GetMetadataReader();
        SignatureDecoder<string, Type> sigDecoder = new(SignatureTypeProviderForComInterop.Instance, reader, declaringType);
        if (name != null) //null means a local var sig 
        {
            var offset = sigBlob.Offset;
            var callConv = sigBlob.ReadCompressedInteger();
            sigBlob.Offset = offset;
            SignatureHeader sigHeader = new((byte)callConv);
            if (sigHeader.Kind == SignatureKind.Field)
            {
                sigString.Append(sigDecoder.DecodeFieldSignature(ref sigBlob));
                if (!string.IsNullOrEmpty(name))
                {
                    sigString.Append(' ');
                    sigString.Append(name);
                }
                return sigString.ToString();
            }
            var methodSig = sigDecoder.DecodeMethodSignature(ref sigBlob);
            if (sigHeader.IsInstance)
            {
                sigString.Append("instance ");
            }
            if (sigHeader.IsGeneric)
            {
                sigString.Append("generic ");
                numTyArgs = methodSig.GenericParameterCount;
            }
            if ((callConv & SignatureHeader.CallingConventionOrKindMask) < IMAGE_CEE_CS_CALLCONV_MAX)
            {
                sigString.Append(_callConvNames[callConv & SignatureHeader.CallingConventionOrKindMask]);
            }
            sigString.Append(methodSig.ReturnType);
            types = methodSig.ParameterTypes;
            requiredParameterCount = methodSig.RequiredParameterCount;
        }
        else
        {
            types = sigDecoder.DecodeLocalSignature(ref sigBlob);
            requiredParameterCount = types.Length;
        }

        if (!string.IsNullOrEmpty(name))
        {
            sigString.Append(' ');
            sigString.Append(name);
        }
        sigString.Append('(');
        foreach (var (typeString, i) in types.Select((t, i) => (t, i)))
        {
            if (i > 0)
            {
                sigString.Append(',');
            }
            if (i == requiredParameterCount)
            {
                sigString.Append("...");
                sigString.Append(',');
            }
            sigString.Append(typeString);
        }
        sigString.Append(')');
        return sigString.ToString();
    }

}
