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
using System.Reflection;
using System.Reflection.Metadata;
using System.Runtime.Versioning;

#if NETFRAMEWORK || NETSTANDARD
using System.Buffers;
using System.Collections.Immutable;
using System.Reflection.PortableExecutable;
#endif

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Extension methods for <see cref="Assembly"/>.
/// </summary>
[Browsable(false)]
public static class AssemblyExtensions
{
    /// <summary>
    /// Returns a assembly identifier.
    /// </summary>
    /// <param name="assembly">An assembly that is used to create an identifer.</param>
    /// <param name="overrideGuid">A guid that should be used</param>
    public static TypeLibIdentifier GetLibIdentifier(this Assembly assembly, Guid overrideGuid)
    {
        var version = assembly.GetTLBVersionForAssembly();

        // From Major, Minor, Language, Guid
        return new TypeLibIdentifier()
        {
            Name = assembly.GetName().Name ?? string.Empty,
            MajorVersion = (ushort)version.Major,
            MinorVersion = (ushort)version.Minor,
            LibID = overrideGuid == Guid.Empty ? assembly.GetTLBGuidForAssembly() : overrideGuid,
            LanguageIdentifier = assembly.GetTLBLanguageIdentifierForAssembly()
        };
    }

    /// <summary>
    /// Returns a assembly identifier.
    /// </summary>
    /// <param name="assembly">An assembly that is used to create an identifer.</param>
    public static TypeLibIdentifier GetLibIdentifier(this Assembly assembly)
    {
        return assembly.GetLibIdentifier(Guid.Empty);
    }

    /// <summary>
    /// Returns the <see cref="FrameworkName"/> from a assembly if it was specified.
    /// </summary>
    /// <param name="assembly">An assembly for that a <see cref="FrameworkName"/> is retrieved.</param>
    /// <returns><see cref="FrameworkName"/> if it was specified for the assembly.</returns>
    public static FrameworkName? GetFrameworkName(this Assembly assembly)
    {
        var targetFrameworkAttribute = assembly.GetCustomAttribute<TargetFrameworkAttribute>();
        if (targetFrameworkAttribute != null)
        {
            return new FrameworkName(targetFrameworkAttribute.FrameworkName);
        }
        return null;
    }

    internal static IEnumerable<Type> GetLoadableTypes(this Assembly assembly)
    {
        return GetLoadableTypesAndLog(assembly, null);
    }

    internal static IEnumerable<Type> GetLoadableTypesAndLog(this Assembly assembly, WriterContext? context)
    {
        try
        {
            return assembly.GetTypes();
        }
        // https://stackoverflow.com/questions/7889228/how-to-prevent-reflectiontypeloadexception-when-calling-assembly-gettypes
        catch (ReflectionTypeLoadException e)
        {
            context?.LogWarning($"Type library exporter encountered an error while processing '{assembly.GetName().Name}'. Error: {e.LoaderExceptions.First()!.Message}");

            return e.Types.Where(t => t is not null)!;
        }
    }

    internal static unsafe MetadataReader GetMetadataReader(this Assembly assembly)
    {
        if (assembly.ReflectionOnly &&
            assembly.ManifestModule is IReflectionOnlyModuleExtension extendedROModule)
        {
            return extendedROModule.Reader;
        }
        else if (assembly.TryGetRawMetadata(out var matadataBlob, out var len))
        {
            return new MetadataReader(matadataBlob, len);
        }
        throw new InvalidOperationException($"Assembly {assembly} with Type {assembly.GetType()} doesn't provide metadata.");
    }

    // new API is only available since .NET Core 2.2, see GitHub issue:
    //https://github.com/dotnet/runtime/issues/15017
    // provide a similar implementation for everything older by using internal GetRawBytes from RuntimeAssembly
#if NETFRAMEWORK || NETSTANDARD
    private static readonly MethodInfo? _miGetRawBytes = typeof(object).Assembly.GetType().GetMethod("GetRawBytes", BindingFlags.NonPublic | BindingFlags.Instance);
    private static readonly Delegate? _getRawBytes = _miGetRawBytes?.CreateDelegate(typeof(Func<,>).MakeGenericType(_miGetRawBytes.DeclaringType!, typeof(byte[])), null);
    private static readonly Dictionary<Assembly, (ReadOnlyMemory<byte> mem, MemoryHandle handle)> _rawMetadataHandlesForAssembly = new();
    private static int _cleanUpOnShutdownRegistered;
    private static void RegisterCleanupMemoryHandlesOnAppDomainUnload()
    {
        if (Interlocked.Exchange(ref _cleanUpOnShutdownRegistered, 1) == 0)
        {
            AppDomain.CurrentDomain.DomainUnload += static (sender, e) =>
            {
                foreach (var (_, (_, handle)) in _rawMetadataHandlesForAssembly)
                {
                    handle.Dispose();
                }
            };
        }
    }
    internal static unsafe bool TryGetRawMetadata(this Assembly assembly, out byte* blob, out int length)
    {
        if (assembly is null)
        {
            throw new ArgumentNullException(nameof(assembly));
        }
        if (_getRawBytes is null)
        {
            throw new PlatformNotSupportedException($"GetRawBytes is missing in Type {typeof(object).Assembly.GetType()}");
        }
        if (assembly.GetType() != typeof(object).Assembly.GetType())
        {
            throw new ArgumentOutOfRangeException(nameof(assembly), $"Assembly must be of Type {typeof(object).Assembly.GetType()} but is {assembly.GetType()}");
        }
        RegisterCleanupMemoryHandlesOnAppDomainUnload();
        if (!_rawMetadataHandlesForAssembly.TryGetValue(assembly, out var metadataHandle))
        {
            var assemblyBytes = (byte[])_getRawBytes.DynamicInvoke(assembly)!;
            using PEReader reader = new(assemblyBytes.ToImmutableArray());
            var metadataBlock = reader.GetMetadata();
            var metadataBytes = metadataBlock.GetContent();
            var memory = metadataBytes.AsMemory();
            _rawMetadataHandlesForAssembly.Add(assembly, metadataHandle = (memory, memory.Pin()));
        }
        blob = (byte*)metadataHandle.handle.Pointer;
        length = metadataHandle.mem.Length;
        return true;
    }
#endif
}
