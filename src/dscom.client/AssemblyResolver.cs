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

using System.Reflection;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Uses the "ASMPath" option to handle the AppDomain.CurrentDomain.AssemblyResolve event and try to load the specified assemblies.
/// </summary>
internal sealed class AssemblyResolver : MetadataAssemblyResolver, IDisposable
{
    private readonly MetadataLoadContext _context;
    private bool _disposedValue;

    internal AssemblyResolver(TypeLibConverterOptions options)
    {
        Options = options;
        _context = new MetadataLoadContext(this);
    }

    public override Assembly? Resolve(MetadataLoadContext context, AssemblyName assemblyName)
    {
        var dir = Path.GetDirectoryName(Options.Assembly);

        var asmPaths = Options.ASMPath;
        if (Directory.Exists(dir))
        {
            asmPaths = asmPaths.Prepend(dir).ToArray();
        }

        foreach (var path in asmPaths)
        {
            var dllToLoad = Path.Combine(path, $"{assemblyName.Name}.dll");
            if (File.Exists(dllToLoad))
            {
                return context.LoadFromAssemblyPath(dllToLoad);
            }

            var exeToLoad = Path.Combine(path, $"{assemblyName.Name}.exe");
            if (File.Exists(exeToLoad))
            {
                return context.LoadFromAssemblyPath(exeToLoad);
            }
        }

        // Last resort, search for simple name in currently loaded runtime assemblies.
        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
        {
            var an = assembly.GetName();
            if (an.Name == assemblyName.Name &&
                !string.IsNullOrEmpty(assembly.Location))
            {
                return context.LoadFromAssemblyPath(assembly.Location);
            }
        }

        return null;
    }

    public ROAssemblyExtended LoadAssembly(string path)
    {
        return new(_context.LoadFromAssemblyPath(path));
    }

    public TypeLibConverterOptions Options { get; }

    private void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                _context.Dispose();
            }

            _disposedValue = true;
        }
    }

    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
