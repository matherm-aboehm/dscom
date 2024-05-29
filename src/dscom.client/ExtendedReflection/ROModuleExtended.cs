using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;

namespace dSPACE.Runtime.InteropServices;

internal sealed class ROModuleExtended : Module, IReflectionOnlyModuleExtension
{
    private readonly ROAssemblyExtended _roAssemblyExtended;
    private readonly Module _roModule;
    private static Func<Module, MetadataReader>? _getInternalReaderFunc;

    public ROModuleExtended(ROAssemblyExtended roAssemblyExtended, Module roModule)
    {
        _roAssemblyExtended = roAssemblyExtended._roAssembly == roModule.Assembly ?
            roAssemblyExtended : new ROAssemblyExtended(roModule.Assembly);
        _roModule = roModule;
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
    }

    public override Assembly Assembly => _roAssemblyExtended;
    internal const string UnknownStringMessageInRAF = "Returns <Unknown> for modules with no file path";

#if NETCOREAPP
    [RequiresAssemblyFiles(UnknownStringMessageInRAF)]
#endif
    public override string FullyQualifiedName => _roModule.FullyQualifiedName;
    public override IEnumerable<CustomAttributeData> CustomAttributes => base.CustomAttributes;
    public MetadataReader Reader => _getInternalReaderFunc!(_roModule);

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
    }
    #region Object overrides
    public override bool Equals(object? obj)
        => obj is not null && (obj is ROModuleExtended extended ?
            _roModule.Equals(extended._roModule) : _roModule.Equals(obj));
    public override int GetHashCode() => _roModule.GetHashCode();
    public override string ToString() => _roModule.ToString();

    #endregion
}
