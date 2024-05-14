using System.Reflection.Metadata;

namespace dSPACE.Runtime.InteropServices;

public interface IReflectionOnlyModuleExtension
{
    MetadataReader Reader { get; }
}
