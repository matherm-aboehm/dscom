using System;
using System.Globalization;

[assembly: AssemblyFixture(typeof(dSPACE.Runtime.InteropServices.Tests.AssemblyFixture))]

namespace dSPACE.Runtime.InteropServices.Tests;

public sealed class AssemblyFixture : IDisposable
{
    private readonly CultureInfo? _oldDefaultThreadUICulture;
    public AssemblyFixture()
    {
        _oldDefaultThreadUICulture = CultureInfo.DefaultThreadCurrentUICulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.CurrentCulture;
    }

    public void Dispose()
    {
        CultureInfo.DefaultThreadCurrentUICulture = _oldDefaultThreadUICulture;
    }
}
