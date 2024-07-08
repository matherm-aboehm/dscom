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
    private readonly bool _enableLogging;

    internal AssemblyResolver(TypeLibConverterOptions options)
    {
        Options = options;
        _context = new MetadataLoadContext(this);
        // Do not enable logging until MetadataLoadContext is created, to filter out
        // some missing core assemblies, but the constructor should still throw if all
        // core assemblies are missing.
        _enableLogging = true;
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
            if (an.Name == assemblyName.Name)
            {
#pragma warning disable IL3000 // single file case is handled below
                if (!string.IsNullOrEmpty(assembly.Location))
                {
                    return context.LoadFromAssemblyPath(assembly.Location);
#pragma warning restore IL3000
                }
                else
                {
                    return context.LoadFromStream(new MetadataOnlyPEImageStream(assembly));
                }
            }
        }

        LogWarning(0, $"Failed to resolve {assemblyName.Name} from the following directories: {string.Join(", ", asmPaths)}");

        return null;
    }

    private void LogWarning(int eventCode, string eventMsg)
    {
        if (_enableLogging)
        {
            Console.Error.WriteLine($"dscom : warning TX{eventCode:X8} : {eventMsg}");
        }
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
