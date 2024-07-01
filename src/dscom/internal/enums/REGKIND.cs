#pragma warning disable 1591

using System.ComponentModel;

namespace dSPACE.Runtime.InteropServices.ComTypes.Internal;

/// <summary>
/// Controls how a type library is registered.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public enum REGKIND
{
    /// <summary>
    /// Use default register behavior.
    /// </summary>
    DEFAULT,

    /// <summary>
    /// Register this type library.
    /// </summary>
    REGISTER,

    /// 
    /// <summary>Do not register this type library.
    /// </summary>
    NONE,

    /// <summary>
    /// Load this type library for 32-bit targets.
    /// </summary>
    REGKIND_LOAD_TLB_AS_32BIT = 0x20,
    /// <summary>
    /// Load this type library for 64-bit targets.
    /// </summary>
    REGKIND_LOAD_TLB_AS_64BIT = 0x40
}
