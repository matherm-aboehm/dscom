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

using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using dSPACE.Runtime.InteropServices.Exporter;
using dSPACE.Runtime.InteropServices.Writer;

namespace dSPACE.Runtime.InteropServices;

/// <summary>
/// Provides a set of services that convert a managed assembly to a COM type library ot to convert a type library to a text file.
/// </summary>
[SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses", Justification = "Compatibility to the mscorelib TypeLibConverter class")]
[SuppressMessage("Microsoft.Performance", "CA1822:MarkMembersAsStatic", Justification = "Compatibility to the mscorelib TypeLibConverter class")]
[Guid("F1C3BF79-C3E4-11d3-88E7-00902754C43A")]
[ClassInterface(ClassInterfaceType.None)]
[ComVisible(true)]
public class TypeLibConverter : ITypeLibConverter
{
    /// <summary>Converts an assembly to a COM type library.</summary>
    /// <param name="assembly">The assembly to convert.</param>
    /// <param name="tlbFilePath">The file pathe of the resulting type library.</param>
    /// <param name="notifySink">The <see cref="T:dSPACE.Runtime.InteropServices.ITypeLibExporterNotifySink" /> interface implemented by the caller.</param>
    /// <returns>An object that implements the <see langword="ITypeLib" /> interface.</returns>
    public object? ConvertAssemblyToTypeLib(Assembly assembly, string tlbFilePath, ITypeLibExporterNotifySink? notifySink)
    {
        var options = new TypeLibConverterSettings
        {
            Assembly = assembly.IsDynamic ? string.Empty : assembly.Location,
            Out = tlbFilePath,
        };

        if (!string.IsNullOrEmpty(tlbFilePath))
        {
            options.TLBRefpath = new[] { Path.GetDirectoryName(tlbFilePath)! };
        }

        return ConvertAssemblyToTypeLib(assembly, options, notifySink);
    }

    /// <summary>Converts an assembly to a COM type library.</summary>
    /// <returns>An object that implements the <see langword="ITypeLib" /> interface.</returns>
    /// <param name="assembly">The assembly to convert.</param>
    /// <param name="typeLibName">The file name of the resulting type library.</param>
    /// <param name="flags">A <see cref="T:dSPACE.Runtime.InteropServices.TypeLibExporterFlags" /> value indicating any special settings.</param>
    /// <param name="notifySink">The <see cref="T:dSPACE.Runtime.InteropServices.ITypeLibExporterNotifySink" /> interface implemented by the caller.</param>
    public object? ConvertAssemblyToTypeLib(Assembly assembly, string typeLibName, TypeLibExporterFlags flags, ITypeLibExporterNotifySink? notifySink)
    {
        var options = new TypeLibConverterSettings
        {
            Assembly = assembly.IsDynamic ? string.Empty : assembly.Location,
            Out = typeLibName,
        };

        if ((flags & TypeLibExporterFlags.ExportAs32Bit) != 0)
        {
            options.Create64BitTlb = false;
        }
        else if ((flags & TypeLibExporterFlags.ExportAs64Bit) != 0)
        {
            options.Create64BitTlb = true;
        }

        if (!string.IsNullOrEmpty(typeLibName))
        {
            var typeLibDir = Path.GetDirectoryName(typeLibName);
            if (string.IsNullOrEmpty(typeLibDir))
            {
                typeLibDir = Environment.CurrentDirectory;
            }

            options.TLBRefpath = new[] { typeLibDir };
        }

        return ConvertAssemblyToTypeLib(assembly, options, notifySink);
    }

    /// <summary>Converts an assembly to a COM type library.</summary>
    /// <param name="assembly">The assembly to convert.</param>
    /// <param name="settings">The <see cref="T:dSPACE.Runtime.InteropServices.TypeLibConverterSettings" /> to configure the converter.</param>
    /// <param name="notifySink">The <see cref="T:dSPACE.Runtime.InteropServices.ITypeLibExporterNotifySink" /> interface implemented by the caller.</param>
    /// <returns>An object that implements the <see langword="ITypeLib" /> interface.</returns>
    public object? ConvertAssemblyToTypeLib(Assembly assembly, TypeLibConverterSettings settings, ITypeLibExporterNotifySink? notifySink)
    {
        CheckPlatform();

        SYSKIND syskind = (settings.Create64BitTlb ?? Environment.Is64BitProcess) ?
            SYSKIND.SYS_WIN64 : SYSKIND.SYS_WIN32;

        OleAut32.CreateTypeLib2(syskind, settings.Out!, out var typelib).ThrowIfFailed("Failed to create type library.");
        using var writer = new LibraryWriter(assembly, new WriterContext(settings, typelib, notifySink));
        writer.Create();

        return typelib;
    }

#if NET48
    private sealed class WrappingTypeLibImporterSink : System.Runtime.InteropServices.ITypeLibImporterNotifySink
    {
        private ITypeLibImporterNotifySink _sink;
        public WrappingTypeLibImporterSink(ITypeLibImporterNotifySink sink)
        {
            _sink = sink;
        }
        public void ReportEvent(System.Runtime.InteropServices.ImporterEventKind eventKind, int eventCode, string eventMsg)
            => _sink.ReportEvent((ImporterEventKind)eventKind, eventCode, eventMsg);

        public Assembly ResolveRef(object typeLib) => _sink.ResolveRef(typeLib);
    }
#endif

    /// <summary>Converts a COM type library to an assembly.</summary>
    /// <returns>An <see cref="T:System.Reflection.Emit.AssemblyBuilder" /> object containing the converted type library.</returns>
    /// <param name="typeLib">The object that implements the <see langword="ITypeLib" /> interface.</param>
    /// <param name="asmFileName">The file name of the resulting assembly.</param>
    /// <param name="flags">A <see cref="T:dSPACE.Runtime.InteropServices.TypeLibImporterFlags" /> value indicating any special settings.</param>
    /// <param name="notifySink"><see cref="T:dSPACE.Runtime.InteropServices.ITypeLibImporterNotifySink" /> interface implemented by the caller.</param>
    /// <param name="publicKey">A <see langword="byte" /> array containing the public key.</param>
    /// <param name="keyPair">A <see cref="T:System.Reflection.StrongNameKeyPair" /> object containing the public and private cryptographic key pair.</param>
    /// <param name="asmNamespace">The namespace for the resulting assembly.</param>
    /// <param name="asmVersion">The version of the resulting assembly. If <see langword="null" />, the version of the type library is used.</param>
#if !NET48
    [Obsolete(nameof(ConvertTypeLibToAssembly) + "is not supported yet and throws PlatformNotSupportedException.")]
#endif
    public AssemblyBuilder? ConvertTypeLibToAssembly([MarshalAs(UnmanagedType.Interface)] object typeLib, string asmFileName, TypeLibImporterFlags flags, ITypeLibImporterNotifySink? notifySink, byte[]? publicKey, StrongNameKeyPair? keyPair, string? asmNamespace, Version? asmVersion)
    {
#if NET48
        var sinkwrapper = notifySink != null ? new WrappingTypeLibImporterSink(notifySink) : null;
        return new System.Runtime.InteropServices.TypeLibConverter().ConvertTypeLibToAssembly(
            typeLib, asmFileName, (System.Runtime.InteropServices.TypeLibImporterFlags)flags, sinkwrapper, publicKey, keyPair, asmNamespace, asmVersion);
#else
        throw new PlatformNotSupportedException();
#endif
    }

    /// <summary>Gets the name and code base of a primary interop assembly for a specified type library.</summary>
    /// <returns><see langword="true" /> if the primary interop assembly was found in the registry; otherwise <see langword="false" />.</returns>
    /// <param name="g">The GUID of the type library.</param>
    /// <param name="major">The major version number of the type library.</param>
    /// <param name="minor">The minor version number of the type library.</param>
    /// <param name="lcid">The LCID of the type library.</param>
    /// <param name="asmName">On successful return, the name of the primary interop assembly associated with <paramref name="g" />.</param>
    /// <param name="asmCodeBase">On successful return, the code base of the primary interop assembly associated with <paramref name="g" />.</param>
#if !NET48
    [Obsolete(nameof(GetPrimaryInteropAssembly) + "is not supported yet and throws PlatformNotSupportedException.")]
#endif
    public bool GetPrimaryInteropAssembly(Guid g, int major, int minor, int lcid, out string asmName, out string asmCodeBase)
    {
#if NET48
        return new System.Runtime.InteropServices.TypeLibConverter().GetPrimaryInteropAssembly(
            g, major, minor, lcid, out asmName, out asmCodeBase);
#else
        throw new PlatformNotSupportedException();
#endif
    }

    /// <summary>Converts a COM type library to an assembly.</summary>
    /// <returns>An <see cref="T:System.Reflection.Emit.AssemblyBuilder" /> object containing the converted type library.</returns>
    /// <param name="typeLib">The object that implements the <see langword="ITypeLib" /> interface.</param>
    /// <param name="asmFileName">The file name of the resulting assembly.</param>
    /// <param name="flags">A <see cref="T:dSPACE.Runtime.InteropServices.TypeLibImporterFlags" /> value indicating any special settings.</param>
    /// <param name="notifySink"><see cref="T:dSPACE.Runtime.InteropServices.ITypeLibImporterNotifySink" /> interface implemented by the caller.</param>
    /// <param name="publicKey">A <see langword="byte" /> array containing the public key.</param>
    /// <param name="keyPair">A <see cref="T:System.Reflection.StrongNameKeyPair" /> object containing the public and private cryptographic key pair.</param>
    /// <param name="unsafeInterfaces">If <see langword="true" />, the interfaces require link time checks for <see cref="F:System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode" /> permission. If <see langword="false" />, the interfaces require run time checks that require a stack walk and are more expensive, but help provide greater protection.</param>
#if !NET48
    [Obsolete(nameof(ConvertTypeLibToAssembly) + "is not supported yet and throws PlatformNotSupportedException.")]
#endif
    public AssemblyBuilder? ConvertTypeLibToAssembly([MarshalAs(UnmanagedType.Interface)] object typeLib, string asmFileName, int flags, ITypeLibImporterNotifySink? notifySink, byte[]? publicKey, StrongNameKeyPair? keyPair, bool unsafeInterfaces = false)
    {
#if NET48
        var sinkwrapper = notifySink != null ? new WrappingTypeLibImporterSink(notifySink) : null;
        return new System.Runtime.InteropServices.TypeLibConverter().ConvertTypeLibToAssembly(
            typeLib, asmFileName, flags, sinkwrapper, publicKey, keyPair, unsafeInterfaces);
#else
        throw new PlatformNotSupportedException();
#endif
    }

    /// <summary>
    /// Export a type library to a text file.
    /// Creates a new file, writes the specified string to the file, and then closes the file. If the target file already exists, it is overwritten.
    /// </summary>
    /// <param name="settings">The <see cref="TypeLibTextConverterSettings"/> object.</param>
    public void ConvertTypeLibToText(TypeLibTextConverterSettings settings)
    {
        CheckPlatform();

        var refTypeLibs = LoadTypeLibrariesFromOptions(settings).ToList();

        File.WriteAllText(settings.Out, GetYamlTextFromTlb(settings.TypeLibrary, settings.FilterRegex));
        // Need to keep the referenced type lib com objects alive until conversion has finished
        // see: https://stackoverflow.com/questions/40749975/accessing-an-itypeinfo-that-references-an-itypeinfo-from-an-importlib-ed-unregis
        GC.KeepAlive(refTypeLibs);
    }

    [ExcludeFromCodeCoverage] // UnitTest with dependent type libraries is not supported
    private static IEnumerable<ITypeLib> LoadTypeLibrariesFromOptions(TypeLibTextConverterSettings options)
    {
        foreach (var tlbFile in options.TLBReference)
        {
            if (!File.Exists(tlbFile))
            {
                throw new ArgumentException($"File {tlbFile} not exist.");
            }

            OleAut32.LoadTypeLibEx(tlbFile, REGKIND.NONE, out var typeLib).ThrowIfFailed();
            yield return typeLib;
        }

        foreach (var tlbDirectory in options.TLBRefpath)
        {
            if (!Directory.Exists(tlbDirectory))
            {
                throw new ArgumentException($"Directory {tlbDirectory} not exist.");
            }

            foreach (var tlbFile in Directory.GetFiles(tlbDirectory, "*.tlb"))
            {
                // Try to load any tlb found in the folder and ignore any fails.
                var hr = OleAut32.LoadTypeLibEx(tlbFile, REGKIND.NONE, out var typeLib);
                if (hr.Succeeded)
                {
                    yield return typeLib;
                }
            }
        }
    }

    // UnitTest on different platforms is not supported
    [ExcludeFromCodeCoverage]
    private static void CheckPlatform()
    {
        // Only Windows is supported, because we need COM
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            throw new PlatformNotSupportedException();
        }
    }

    /// <summary>
    /// Load a type library and convert the library to YAML text.
    /// </summary>
    /// <param name="inputTlb">The path to a type library</param>
    /// <param name="filters">An array of regular expressions that can be used to filter the output.</param>
    /// <returns>The YAML output.</returns>
    private static string GetYamlTextFromTlb(string inputTlb, string[]? filters)
    {
        var typeLibInfo = new TypelLibInfo(inputTlb);

        StringBuilder builder = new();
        var regExs = filters != null ? filters.ToList().Select(f => new Regex(f)) : Array.Empty<Regex>();
        CreateYaml(builder, typeLibInfo, regExs);
        return builder.ToString();
    }

    /// <summary>
    /// Create a YAML string from <see cref="BaseInfo"/>,
    /// </summary>
    /// <param name="stringBuilder">The <see cref="StringBuilder"/> to use.</param>
    /// <param name="data">The <see cref="BaseInfo"/> to use.</param>
    /// <param name="filters">An array of regex filter.</param>
    /// <param name="indentLevel">The indentation level</param>
    /// <param name="collectionIndex">The collection index, if the item is collection item; otherwise -1</param>
    private static void CreateYaml(StringBuilder stringBuilder, BaseInfo data, IEnumerable<Regex>? filters, int indentLevel = 0, int collectionIndex = -1)
    {
        var type = data.GetType();

        var propertyIndex = 0;
        foreach (var property in type.GetProperties())
        {
            var name = property.Name;
            name = char.ToLowerInvariant(name[0]) + name.Substring(1);
            var value = property.GetValue(data);

            var isFirstPropertyInFirstCollectionItem = propertyIndex == 0 && collectionIndex != -1;
            if (property.GetCustomAttributes(true).Any(c => c is IgnoreAttribute) || value == null)
            {
                continue;
            }

            if (value is IEnumerable elements and not string)
            {
                if (!elements.OfType<object>().Any())
                {
                    continue;
                }
            }

            // Use all Regex strings to filter the output..
            if (filters != null)
            {
                var isMatch = false;
                foreach (var filter in filters)
                {
                    var path = $"{data.GetPath()}.{name}={value}";

                    if (filter.IsMatch(path))
                    {
                        isMatch = true;
                        continue;
                    }
                }

                if (isMatch)
                {
                    continue;
                }
            }

            for (var i = 0; i < (indentLevel - (isFirstPropertyInFirstCollectionItem ? 1 : 0)); i++)
            {
                stringBuilder.Append("  ");
            }

            if (isFirstPropertyInFirstCollectionItem)
            {
                stringBuilder.Append("- ");
            }

            if (value is BaseInfo baseInfo)
            {

                stringBuilder.AppendLine($"{name}:");
                CreateYaml(stringBuilder, baseInfo, filters, indentLevel + 1);
            }
            else if (value is IEnumerable items and not string)
            {
                stringBuilder.AppendLine($"{name}:");
                var index = 0;
                foreach (var item in items)
                {
                    if (item is BaseInfo childBaseItem)
                    {
                        CreateYaml(stringBuilder, childBaseItem, filters, indentLevel + 1, index);
                    }

                    index++;
                }
            }
            else if (value != null)
            {
                stringBuilder.AppendLine($"{name}: {value}");
            }

            propertyIndex++;
        }
    }
}
