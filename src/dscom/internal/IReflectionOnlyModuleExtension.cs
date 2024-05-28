using System.Reflection.Metadata;

namespace dSPACE.Runtime.InteropServices;

internal interface IReflectionOnlyModuleExtension
{
    MetadataReader Reader { get; }
}
