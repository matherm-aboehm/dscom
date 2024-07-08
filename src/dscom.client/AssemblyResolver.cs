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
using System.Runtime.Loader;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Uses the "ASMPath" option to handle the AppDomain.CurrentDomain.AssemblyResolve event and try to load the specified assemblies.
/// </summary>
internal sealed class AssemblyResolver : MetadataAssemblyResolver, IDisposable
{
    private readonly string[] _paths;
    private readonly MetadataLoadContext _reflectionOnlyContext;
    private readonly AssemblyLoadContext _runtimeContext;

    private bool _disposedValue;
    private readonly bool _enableLogging;

    internal AssemblyResolver(string[] paths)
    {
        _paths = paths;
        _reflectionOnlyContext = new MetadataLoadContext(this);
        _runtimeContext = new AssemblyLoadContext("dscom", true);
        _runtimeContext.Resolving += Context_Resolving;
        // Do not enable logging until MetadataLoadContext is created, to filter out
        // some missing core assemblies, but the constructor should still throw if all
        // core assemblies are missing.
        _enableLogging = true;
    }

    private Assembly? Context_Resolving(AssemblyLoadContext context, AssemblyName name)
    {
        foreach (var path in _paths)
        {
            var dllToLoad = Path.Combine(path, $"{name.Name}.dll");
            if (File.Exists(dllToLoad))
            {
                return context.LoadFromAssemblyPath(dllToLoad);
            }

            var exeToLoad = Path.Combine(path, $"{name.Name}.exe");
            if (File.Exists(exeToLoad))
            {
                return context.LoadFromAssemblyPath(exeToLoad);
            }
        }

        LogWarning(0, $"Failed to resolve {name.Name} from the following directories: {string.Join(", ", _paths)}");

        return null;
    }

    public override Assembly? Resolve(MetadataLoadContext context, AssemblyName assemblyName)
    {
        foreach (var path in _paths)
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
        var rtAssembly = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(
            a => a.GetName().Name == assemblyName.Name);
        if (_enableLogging && rtAssembly == null)
        {
            // Or try to load into runtime context using AssemblyName 
            // (use default runtime resolution)
            using var scope = _runtimeContext.EnterContextualReflection();
            rtAssembly = _runtimeContext.LoadFromAssemblyName(assemblyName);
        }

        if (rtAssembly != null)
        {
#pragma warning disable IL3000 // single file case is handled below
            if (!string.IsNullOrEmpty(rtAssembly.Location))
            {
                return context.LoadFromAssemblyPath(rtAssembly.Location);
#pragma warning restore IL3000
            }
            else
            {
                return context.LoadFromStream(new MetadataOnlyPEImageStream(rtAssembly));
            }
        }

        LogWarning(0, $"Failed to resolve {assemblyName.Name} from the following directories: {string.Join(", ", _paths)}");

        return null;
    }

    private void LogWarning(int eventCode, string eventMsg)
    {
        if (_enableLogging)
        {
            Console.Error.WriteLine($"dscom : warning TX{eventCode:X8} : {eventMsg}");
        }
    }

    public Assembly LoadAssembly(string path)
    {
        return _runtimeContext.LoadFromAssemblyPath(path);
    }

    public ROAssemblyExtended LoadROAssembly(string path)
    {
        return new(_reflectionOnlyContext.LoadFromAssemblyPath(path));
    }

    private void Dispose(bool disposing)
    {
        if (!_disposedValue)
        {
            if (disposing)
            {
                _reflectionOnlyContext.Dispose();
                _runtimeContext.Resolving -= Context_Resolving;
                _runtimeContext.Unload();
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
