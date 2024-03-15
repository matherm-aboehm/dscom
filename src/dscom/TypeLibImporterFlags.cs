using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

/// <summary>Indicates how an assembly should be produced.</summary>
[Serializable]
[Flags]
[ComVisible(true)]
public enum TypeLibImporterFlags
{
    /// <summary>No special settings. This is the default.</summary>
    None = 0x0,
    /// <summary>Generates a primary interop assembly. For more information, see the <see cref="T:System.Runtime.InteropServices.PrimaryInteropAssemblyAttribute" /> attribute. A keyfile must be specified.</summary>
    PrimaryInteropAssembly = 0x1,
    /// <summary>Imports all interfaces as interfaces that suppress the common language runtime's stack crawl for <see cref="F:System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode" /> permission. Be sure you understand the responsibilities associated with suppressing this security check.</summary>
    UnsafeInterfaces = 0x2,
    /// <summary>Imports all <see langword="SAFEARRAY" /> instances as <see cref="T:System.Array" /> instead of typed, single-dimensional, zero-based managed arrays. This option is useful when dealing with multi-dimensional, non-zero-based <see langword="SAFEARRAY" /> instances, which otherwise cannot be accessed unless you edit the resulting assembly by using the MSIL Disassembler (Ildasm.exe) and MSIL Assembler (Ilasm.exe) tools.</summary>
    SafeArrayAsSystemArray = 0x4,
    /// <summary>Transforms <see langword="[out, retval]" /> parameters of methods on dispatch-only interfaces (dispinterface) into return values.</summary>
    TransformDispRetVals = 0x8,
    /// <summary>Not used.</summary>
    PreventClassMembers = 0x10,
    /// <summary>Uses serializable classes.</summary>
    SerializableValueClasses = 0x20,
    /// <summary>Imports a type library for the x86 platform.</summary>
    ImportAsX86 = 0x100,
    /// <summary>Imports a type library for the x86 64-bit platform.</summary>
    ImportAsX64 = 0x200,
    /// <summary>Imports a type library for the Itanium platform.</summary>
    ImportAsItanium = 0x400,
    /// <summary>Imports a type library for any platform.</summary>
    ImportAsAgnostic = 0x800,
    /// <summary>Uses reflection-only loading.</summary>
    ReflectionOnlyLoading = 0x1000,
    /// <summary>Prevents inclusion of a version resource in the interop assembly. For more information, see the <see cref="M:System.Reflection.Emit.AssemblyBuilder.DefineVersionInfoResource" /> method.</summary>
    NoDefineVersionResource = 0x2000,
    /// <summary>Imports a library for the ARM platform.</summary>
    ImportAsArm = 0x4000
}
