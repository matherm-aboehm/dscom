#if NET5_0_OR_GREATER
using System.Diagnostics.CodeAnalysis;
#endif
#if NETCOREAPP
using System.Linq.Expressions;
#endif
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;

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
    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROModuleExtended extended ?
            _roModule.Equals(extended._roModule) : _roModule.Equals(obj));
    public override int GetHashCode() => _roModule.GetHashCode();
    public override string ToString() => _roModule.ToString();

    #endregion
}
