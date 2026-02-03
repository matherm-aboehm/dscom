#if NET5_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
#endif
using System.Diagnostics;
#if NETCOREAPP
using System.Linq.Expressions;
#endif
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROModuleExtended : Module, IReflectionOnlyModuleExtension
{
    private readonly ROAssemblyExtended _roAssemblyExtended;
    private readonly Module _roModule;
#if NETCOREAPP
    private static Func<Module, MetadataReader>? _getInternalReaderFunc;
#else
    private readonly Lazy<MetadataReader> _metadataReader;
#endif

    public ROModuleExtended(ROAssemblyExtended roAssemblyExtended, Module roModule)
    {
        _roAssemblyExtended = roAssemblyExtended._roAssembly == roModule.Assembly ?
            roAssemblyExtended : new ROAssemblyExtended(roModule.Assembly);
        _roModule = roModule;
#if NETCOREAPP
        if (_getInternalReaderFunc == null)
        {
            var roModuleType = _roModule.GetType();
            var propReader = roModuleType.GetProperty(nameof(Reader),
                BindingFlags.Instance | BindingFlags.NonPublic);
            var funcType = typeof(Func<,>).MakeGenericType(roModuleType, typeof(MetadataReader));
            var typedFunc = propReader?.GetGetMethod(true)?
                .CreateDelegate(funcType, null)
                ?? throw new MissingMemberException(roModuleType.FullName, nameof(Reader));
            var typedFuncExp = Expression.Constant(typedFunc);
            var modParamExp = Expression.Parameter(typeof(Module), "mod");
            var bodyExp = Expression.Invoke(typedFuncExp,
                Expression.Convert(modParamExp, roModuleType));
            _getInternalReaderFunc = Expression.Lambda<Func<Module, MetadataReader>>(
                bodyExp, modParamExp).Compile();
        }
#else
        _metadataReader = new Lazy<MetadataReader>(() =>
        {
            unsafe
            {
                var runtimeAssembly = _roAssemblyExtended._roAssembly;
                if (runtimeAssembly.TryGetRawMetadata(out var matadataBlob, out var len))
                {
                    return new MetadataReader(matadataBlob, len);
                }
                else
                {
                    throw new InvalidOperationException($"Assembly {runtimeAssembly} with Type {runtimeAssembly.GetType()} doesn't provide metadata.");
                }
            }
        });
#endif
    }

    public override Assembly Assembly => _roAssemblyExtended;
    internal const string UnknownStringMessageInRAF = "Returns <Unknown> for modules with no file path";

#if NET5_0_OR_GREATER
    [RequiresAssemblyFiles(UnknownStringMessageInRAF)]
#endif
    public override string FullyQualifiedName => _roModule.FullyQualifiedName;
    public override IEnumerable<CustomAttributeData> CustomAttributes => base.CustomAttributes;
#if NETCOREAPP
    public MetadataReader Reader => _getInternalReaderFunc!(_roModule);
#else
    public MetadataReader Reader => _metadataReader.Value;
#endif

    public override IList<CustomAttributeData> GetCustomAttributesData()
        => _roModule.GetCustomAttributesData();
    public override object[] GetCustomAttributes(Type attributeType, bool inherit)
        => GetCustomAttributesData().GetCustomAttributes(attributeType);
    public override object[] GetCustomAttributes(bool inherit)
        => GetCustomAttributesData().GetCustomAttributes();
    public override bool IsDefined(Type attributeType, bool inherit)
        => GetCustomAttributesData().IsDefined(attributeType);

    public override byte[] ResolveSignature(int metadataToken)
    {
#if NETCOREAPP
        var reader = Reader;
        var handle = MetadataTokens.Handle(metadataToken);
        var sigHandle = handle.Kind switch
        {
            HandleKind.MethodDefinition => reader.GetMethodDefinition((MethodDefinitionHandle)handle).Signature,
            HandleKind.FieldDefinition => reader.GetFieldDefinition((FieldDefinitionHandle)handle).Signature,
            HandleKind.PropertyDefinition => reader.GetPropertyDefinition((PropertyDefinitionHandle)handle).Signature,
            HandleKind.StandaloneSignature => reader.GetStandaloneSignature((StandaloneSignatureHandle)handle).Signature,
            HandleKind.TypeSpecification => reader.GetTypeSpecification((TypeSpecificationHandle)handle).Signature,
            HandleKind.MemberReference => reader.GetMemberReference((MemberReferenceHandle)handle).Signature,
            _ => throw new InvalidOperationException($"can not resolve signature of token with kind {handle.Kind}")
        };
        return reader.GetBlobBytes(sigHandle);
#else
        return _roModule.ResolveSignature(metadataToken);
#endif
    }

    public MarshalAsAttribute ComputeMarshalAsAttribute(int token)
    {
        var reader = Reader;
        var handle = MetadataTokens.Handle(token);
        BlobReader marshalReader;
        if (handle.Kind == HandleKind.Parameter)
        {
            var param = reader.GetParameter((ParameterHandle)handle);
            var marshalBlob = param.GetMarshallingDescriptor();
            marshalReader = reader.GetBlobReader(marshalBlob);
        }
        else if (handle.Kind == HandleKind.FieldDefinition)
        {
            var field = reader.GetFieldDefinition((FieldDefinitionHandle)handle);
            var marshalBlob = field.GetMarshallingDescriptor();
            marshalReader = reader.GetBlobReader(marshalBlob);
        }
        else
        {
            throw new InvalidOperationException($"can not use token with kind {handle.Kind} to get the marshalling descriptors");
        }

        MarshalAsAttribute marshalAs;
        var unmanagedType = (UnmanagedType)marshalReader.ReadCompressedInteger();
        switch (unmanagedType)
        {
            case UnmanagedType.LPArray: //NATIVE_TYPE_ARRAY
                {
                    var paramIndex = (short)-1; // use -1 for the attribute,
                                                // as there is no other way to specify
                                                // missing value for non-nullable field
                    var numElements = 0;
                    var elemMult = 0;
                    if (!marshalReader.TryReadCompressedInteger(out var arrayElemType))
                    {
                        arrayElemType = 0x50; //NATIVE_TYPE_MAX
                    }
                    else if (marshalReader.TryReadCompressedInteger(out var paramNum))
                    {
                        //spec says it's 1-based index, but compiler actually writes zero-based
                        paramIndex = (short)paramNum; //paramNum-1
                        elemMult = 1; // default multiplier
                        if (!marshalReader.TryReadCompressedInteger(out numElements))
                        {
                            numElements = 0;
                        }
                        else if (marshalReader.TryReadCompressedInteger(out var flags) &&
                            (flags & 0x1) == 0) //ntaSizeParamIndexSpecified bit not set
                        {
                            paramIndex = -1;
                            elemMult = 0;
                        }
                    }
                    Debug.Assert(elemMult is 0 or 1);
                    Debug.Assert((elemMult == 0 && paramIndex == -1) || (elemMult == 1 && paramIndex >= 0));
                    marshalAs = new(unmanagedType)
                    {
                        ArraySubType = arrayElemType == 0x50 ? default : (UnmanagedType)arrayElemType,
                        SizeParamIndex = paramIndex,
                        SizeConst = numElements
                    };
                }
                break;
            case UnmanagedType.CustomMarshaler: //NATIVE_TYPE_CUSTOMMARSHALLER
                var guidString = marshalReader.ReadSerializedString();
                Guid? tlbguid = string.IsNullOrEmpty(guidString) ? null : Guid.Parse(guidString);
                var customUnmanagedType = marshalReader.ReadSerializedString();
                var customManagedType = marshalReader.ReadSerializedString();
                var cookie = marshalReader.ReadSerializedString();
                Type? customMarshallerType = null;
                if (!string.IsNullOrEmpty(customManagedType))
                {
                    customMarshallerType = _roAssemblyExtended.GetType(customManagedType);
                    if (customMarshallerType is null ||
                        !typeof(ICustomMarshaler).IsAssignableFromReflectionOnly(customMarshallerType))
                    {
                        throw new InvalidOperationException($"Can't find the custom marshaller type with name {customManagedType} in assembly {_roAssemblyExtended.FullName}");
                    }
                }
                else if (tlbguid is not null && !string.IsNullOrEmpty(customUnmanagedType))
                {
                    throw new NotSupportedException($"Native custom marshaller ({tlbguid?.ToString() ?? customUnmanagedType}) is not supported");
                }
                marshalAs = new(unmanagedType)
                {
                    MarshalType = customManagedType,
                    MarshalTypeRef = customMarshallerType,
                    MarshalCookie = cookie
                };
                break;
            case UnmanagedType.ByValArray: //NATIVE_TYPE_FIXEDARRAY
                {
                    var numElements = marshalReader.ReadCompressedInteger();
                    if (!marshalReader.TryReadCompressedInteger(out var arrayElemType))
                    {
                        arrayElemType = 0x50; //NATIVE_TYPE_MAX
                    }
                    marshalAs = new(unmanagedType)
                    {
                        ArraySubType = arrayElemType == 0x50 ? default : (UnmanagedType)arrayElemType,
                        SizeConst = numElements
                    };
                }
                break;
            case UnmanagedType.SafeArray: //NATIVE_TYPE_SAFEARRAY
                //spec says there must be a VT_* constant following the NATIVE_TYPE_SAFEARRAY,
                //but it seems compiler also allows to be missing
                var safeArrayElemType = marshalReader.TryReadCompressedInteger(out var vt) ? (VarEnum)vt : default;
                Type? safeArrayElemUserType = null;
                //also not in spec, but coded in CLR VM
                if (marshalReader.RemainingBytes > 0)
                {
                    var typeName = marshalReader.ReadSerializedString();
                    if (!string.IsNullOrEmpty(typeName))
                    {
                        safeArrayElemUserType = _roAssemblyExtended.GetType(typeName);
                    }
                }
                marshalAs = new(unmanagedType)
                {
                    SafeArraySubType = safeArrayElemType,
                    SafeArrayUserDefinedSubType = safeArrayElemUserType
                };
                break;
            default:
                marshalAs = new(unmanagedType);
                break;
        }

        if (marshalReader.RemainingBytes != 0)
        {
            throw new InvalidOperationException($"the marshalling descriptor blob was not read completely it has {marshalReader.RemainingBytes} bytes left");
        }

        return marshalAs;
    }

    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROModuleExtended extended ?
            _roModule.Equals(extended._roModule) : _roModule.Equals(obj));
    public override int GetHashCode() => _roModule.GetHashCode();
    public override string ToString() => _roModule.ToString();

    #endregion
}
