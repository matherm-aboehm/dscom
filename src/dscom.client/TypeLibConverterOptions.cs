// Copyright 2022 dSPACE GmbH, Mark Lechtermann, Matthias Nissen and Contributors
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
extern alias dscom;
using dscom::dSPACE.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
///  Represents the options used by dscom tlbexport.
/// </summary>
public class TypeLibConverterOptions : TypeLibConverterSettings
{
    /// <summary>
    /// Gets or sets a value indicating whether the output is silent.
    /// </summary>
    public bool Silent { get; set; }

    /// <summary>
    /// Gets or sets an array of warnings that should ignored.
    /// </summary>
    /// <returns></returns>
    public string[] Silence { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Gets or sets a value indicating whether the output should be verbose.
    /// </summary>
    public bool Verbose { get; set; }

    /// <summary>
    /// Gets or sets whether to generate missing TLB files for dependent assemblies.  
    /// </summary>
    public bool? CreateMissingDependentTLBs { get; set; } = true;

    /// <summary>
    /// Gets or sets wether the target platform bitness is 32-bit.
    /// </summary>
    public bool? Win32 { get => !Create64BitTlb; set => Create64BitTlb = !value; }

    /// <summary>
    /// Gets or sets wether the target platform bitness is 64-bit.
    /// </summary>
    public bool? Win64 { get => Create64BitTlb; set => Create64BitTlb = value; }
}
