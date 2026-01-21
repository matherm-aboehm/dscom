using System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices.Test;

[ComVisible(true), Guid("E27C3BDB-ACEF-44F9-8568-481D44681C04"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IExportableType
{
    void DoIt();
}
