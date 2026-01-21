using System.Windows.Controls;

namespace dSPACE.Runtime.InteropServices.Test;

public class ExplodeyBaseClass : UserControl
{
    private sealed class PrivateImplementationDetails
    {
        // required to reproduce #259
    }
}
