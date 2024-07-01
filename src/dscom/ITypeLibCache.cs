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

using System.ComponentModel;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Provides a cache that could be used the returns an ITypeLib.
/// </summary>
[Browsable(false)]
public interface ITypeLibCache
{
    /// <summary>
    /// Returns a ITypeLib.
    /// </summary>
    /// <param name="identifier">An assembly identifier.</param>
    ITypeLib? GetTypeLibFromIdentifier(in TypeLibIdentifier identifier);

    /// <summary>
    /// Loads a type library with information from registry.
    /// </summary>
    /// <param name="identifier">Identification data for a registered type libraray.</param>
    /// <param name="throwOnError">Specifies wether it should throw on any error
    /// or just return <c>null</c>.</param>
    /// <returns>ITypeLib</returns>
    ITypeLib? LoadTypeLibFromIdentifier(in TypeLibIdentifier identifier, bool throwOnError = true);

    /// <summary>
    /// Loads a type library from path.
    /// </summary>
    /// <param name="typeLibPath">Path to the type library.</param>
    /// <param name="throwOnError">Specifies wether it should throw on any error
    /// or just return <c>null</c>.</param>
    /// <returns>ITypeLib</returns>
    ITypeLib? LoadTypeLibFromPath(string typeLibPath, bool throwOnError = true);
}
