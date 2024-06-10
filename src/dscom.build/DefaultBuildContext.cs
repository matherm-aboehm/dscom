// Copyright 2022 dSPACE GmbH, Carsten Igel and Contributors
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
using System.Security;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

using COMException = System.Runtime.InteropServices.COMException;
using SystemInteropServices = System.Runtime.InteropServices;

namespace dSPACE.Runtime.InteropServices.BuildTasks;

/// <summary>
/// Default implementation of the <see cref="IBuildContext" /> interface
/// using <see cref="TypeLibConverter" /> as implementation for conversion
/// and <see cref="LoggingTypeLibExporterSink" /> as implementation for
/// event handling.
/// </summary>
internal sealed class DefaultBuildContext :
#if !NET5_0_OR_GREATER
    MarshalByRefObject,
#endif
    IBuildContext
{
    /// <inheritdoc cref="IBuildContext.IsRunningOnWindows" />
    public bool IsRunningOnWindows => SystemInteropServices.RuntimeInformation.IsOSPlatform(SystemInteropServices.OSPlatform.Windows);

#if NET5_0_OR_GREATER
    /// <inheritdoc cref="IBuildContext.RuntimeDescription" />
    public string RuntimeDescription => $"{SystemInteropServices.RuntimeInformation.OSDescription} {SystemInteropServices.RuntimeInformation.OSArchitecture} ({SystemInteropServices.RuntimeInformation.ProcessArchitecture} [{SystemInteropServices.RuntimeInformation.RuntimeIdentifier} - {SystemInteropServices.RuntimeInformation.FrameworkDescription}])";
#else
    /// <inheritdoc cref="IBuildContext.RuntimeDescription" />
    public string RuntimeDescription => $"{SystemInteropServices.RuntimeInformation.OSDescription} {SystemInteropServices.RuntimeInformation.OSArchitecture} ({SystemInteropServices.RuntimeInformation.ProcessArchitecture} [{SystemInteropServices.RuntimeInformation.FrameworkDescription}])";

    private bool IsLoadContext { get; }

    // Need a proxy type here, so that IDisposable implementation can do AppDomain.Unload
    // from the original AppDomain where the instance of this type was created.
    // It wouldn't work when implemented on the remote object itself, because the Dispose()
    // call would then be redirected to the AppDomain, which is about to be unloaded, and then
    // block the unloading (or throws exception).
    // see: https://learn.microsoft.com/de-de/dotnet/api/system.appdomainunloadedexception?view=net-8.0
    private sealed class BuildContextProxy : IBuildContext, IDisposable
    {
        private readonly IBuildContext _remoteObj;
        private readonly AppDomain _domain;

        public BuildContextProxy(IBuildContext remoteObj, AppDomain domain)
        {
            _remoteObj = remoteObj;
            _domain = domain;
        }

        public void Dispose()
        {
            AppDomain.Unload(_domain);
            GC.SuppressFinalize(this);
        }

        public bool IsRunningOnWindows => _remoteObj.IsRunningOnWindows;
        public string RuntimeDescription => _remoteObj.RuntimeDescription;
        public bool ConvertAssemblyToTypeLib(TypeLibConverterSettings settings, TaskLoggingHelper log)
            => _remoteObj.ConvertAssemblyToTypeLib(settings, log);
        public bool EmbedTypeLib(TypeLibEmbedderSettings settings, TaskLoggingHelper log)
            => _remoteObj.EmbedTypeLib(settings, log);
        public bool EnsureDirectoryExists(string? directoryPath)
            => _remoteObj.EnsureDirectoryExists(directoryPath);
        public bool EnsureFileExists(string? fileNameAndPath)
            => _remoteObj.EnsureFileExists(fileNameAndPath);
    }

    public DefaultBuildContext(bool isLoadContext = false)
    {
        IsLoadContext = isLoadContext;
    }
#endif

    /// <inheritdoc cref="IBuildContext.EnsureFileExists(string?)" />
    public bool EnsureFileExists(string? fileNameAndPath)
    {
        return File.Exists(fileNameAndPath);
    }

    /// <inheritdoc cref="IBuildContext.EnsureDirectoryExists(string?)" />
    public bool EnsureDirectoryExists(string? directoryPath)
    {
        return Directory.Exists(directoryPath);
    }

    /// <inheritdoc cref="IBuildContext.ConvertAssemblyToTypeLib(TypeLibConverterSettings, TaskLoggingHelper)" />
    public bool ConvertAssemblyToTypeLib(TypeLibConverterSettings settings, TaskLoggingHelper log)
    {
        // Load assembly from file.
#if NET5_0_OR_GREATER
        var loadContext = CreateLoadContext(settings, log);
        var assembly = LoadAssembly(settings, loadContext, log);
#else
        if (!IsLoadContext)
        {
            using var loadContext = CreateLoadContext();
            return loadContext.ConvertAssemblyToTypeLib(settings, log);
        }
        var assembly = LoadAssembly(settings, log, out var resolveHandler);
#endif

        if (assembly is null)
        {
            log.LogWarning("Failed to load assembly {0}. Task failed.", settings.Assembly);
            return false;
        }

        try
        {
            // Create type library converter.
            var converter = new TypeLibConverter();

            // Choose appropriate name resolver based on inputs with the Com Alias as the fallback.
            var nameResolver = settings.Names.Length != 0
                ? NameResolver.Create(settings.Names)
                : NameResolver.Create(assembly);

            // Create event handler.
            var sink = new LoggingTypeLibExporterSink(log, nameResolver);

            // create conversion.
            var tlb = converter.ConvertAssemblyToTypeLib(assembly, settings, sink);
            if (tlb == null)
            {
                log.LogError("The following type library could not be created successfully: {0}. Reason: Operation was not successful.", settings.Out);
            }
            else
            {
                log.LogMessage(MessageImportance.High, "Finished generation of the following type library: {0}", settings.Out);
            }

            if (!File.Exists(settings.Out))
            {
                log.LogWarning(
                    null,
                    "DSCOM001",
                    null,
                    null,
                    0,
                    0,
                    0,
                    0,
                    "Could not find the type library at the following location: {0}", settings.Out);
            }

            return tlb != null;
        }
        catch (COMException e)
        {
            log.LogErrorFromException(e, false, true, settings.Assembly);
            return false;
        }
#if NET5_0_OR_GREATER
        finally
        {
            try
            {
                loadContext.Dispose();
            }
            catch (Exception e)
            {
                log.LogWarningFromException(new InvalidOperationException("Failed to unload the assembly load context.", e));
            }
        }
#else
        finally
        {
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve -= resolveHandler;
        }
#endif
    }

    public bool EmbedTypeLib(TypeLibEmbedderSettings settings, TaskLoggingHelper log)
    {
        try
        {
            var result = TypeLibEmbedder.EmbedTypeLib(settings);
            if (!result)
            {
                log.LogError("Could not embed type library {0} into assembly {1}", settings.SourceTypeLibrary, settings.TargetAssembly);
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            log.LogErrorFromException(ex);
            return false;
        }
    }

#if NET5_0_OR_GREATER
    private sealed class BuildContextAssemblyResolver : MetadataAssemblyResolver
    {
        public string LoadContextName { get; }
        public TypeLibConverterSettings Settings { get; }
        public TaskLoggingHelper Log { get; }
        public BuildContextAssemblyResolver(string name, TypeLibConverterSettings settings, TaskLoggingHelper log)
        {
            LoadContextName = name;
            Settings = settings;
            Log = log;
        }

        public override Assembly? Resolve(MetadataLoadContext context, AssemblyName assemblyName)
        {
            if (TryResolveAssemblyFromSettings(assemblyName.Name ?? string.Empty, Settings, Log, out var assemblyPath))
            {
                return context.LoadFromAssemblyPath(assemblyPath);
            }

            Log.LogWarning("Failed to resolve {0} from the following files: {1}", assemblyName.Name, string.Join(", ", Settings.ASMPath));

            return default;
        }
    }

    /// <summary>
    /// Creates an instance of <see cref="MetadataLoadContext"/> that will
    /// take care of the loading and unloading the target assemblies and 
    /// can be unloaded afterwards.
    /// </summary>
    /// <param name="settings">The type library settings.</param>
    /// <returns>A <see cref="MetadataLoadContext"/> that can be unloaded.</returns>
    private static MetadataLoadContext CreateLoadContext(TypeLibConverterSettings settings, TaskLoggingHelper log)
    {
        var loadContext = new MetadataLoadContext(
            new BuildContextAssemblyResolver($"msbuild-load-ctx-{Guid.NewGuid()}",
                settings, log));

        return loadContext;
    }

    /// <summary>
    /// Tries to load the assembly specified in the <paramref name="settings"/> using the specified
    /// <paramref name="loadContext"/>. If the assembly cannot be loaded, the result will be <c>null</c>.
    /// </summary>
    /// <param name="settings">The settings.</param>
    /// <param name="loadContext">The assembly load context.</param>
    /// <param name="log">The log to write messages to.</param>
    /// <returns>The assembly loaded.</returns>
    private static Assembly? LoadAssembly(TypeLibConverterSettings settings, MetadataLoadContext loadContext, TaskLoggingHelper log)
    {
        Assembly assembly;
        try
        {
            assembly = loadContext.LoadFromAssemblyPath(settings.Assembly);
        }
        catch (Exception e) when
            (e is ArgumentNullException
               or FileNotFoundException
               or FileLoadException
               or BadImageFormatException
               or SecurityException
               or ArgumentException
               or PathTooLongException)
        {
            log.LogErrorFromException(e, true, true, settings.Assembly);
            try
            {
                loadContext.Dispose();
            }
            catch (Exception e2)
            {
                log.LogWarningFromException(new InvalidOperationException($"Failed to unload the following assembly: {settings.Assembly}.", e2));
            }

            return default;
        }

        return new ROAssemblyExtended(assembly);
    }
#else
    /// <summary>
    /// Creates an instance of <see cref="BuildContextProxy"/> that will
    /// redirect all calls to a remote object on another <see cref="AppDomain"/>, so that
    /// it can take care of the loading and unloading the target assemblies and 
    /// the other <see cref="AppDomain"/> can be unloaded afterwards.
    /// </summary>
    /// <returns>A <see cref="BuildContextProxy"/> that can be unloaded.</returns>
    /// <exception cref="InvalidOperationException">Creating remote object in other AppDomain failed.</exception>
    private static BuildContextProxy CreateLoadContext()
    {
        var setup = AppDomain.CurrentDomain.SetupInformation;
        var evidence = AppDomain.CurrentDomain.Evidence;
        var newDomain = AppDomain.CreateDomain($"msbuild-load-ctx-{Guid.NewGuid()}", evidence, setup);

        var proxiedType = typeof(DefaultBuildContext);
        var remoteObj = (IBuildContext)newDomain.CreateInstanceAndUnwrap(
            proxiedType.Assembly.FullName, proxiedType.FullName, false,
            BindingFlags.Default, null, new object?[] { true }, null, null)
            ?? throw new InvalidOperationException("Failed to create instance in other AppDomain.");
        var loadContext = new BuildContextProxy(remoteObj, newDomain);
        return loadContext;
    }

    /// <summary>
    /// Tries to load the assembly specified in the <paramref name="settings"/> using the current
    /// <see cref="AppDomain"/>. If the assembly cannot be loaded, the result will be <c>null</c>.
    /// </summary>
    /// <param name="settings">The settings.</param>
    /// <param name="log">The log to write messages to.</param>
    /// <param name="resolveHandler">When the method returns, this parameter contains the
    /// <see cref="ResolveEventHandler" /> that was registered by this method on the current
    /// <see cref="AppDomain"/>. It can be used to unregister it, when it is no longer needed.</param>
    /// <returns>The assembly loaded.</returns>
    private static Assembly? LoadAssembly(TypeLibConverterSettings settings, TaskLoggingHelper log, out ResolveEventHandler resolveHandler)
    {
        var appDomain = AppDomain.CurrentDomain;
        resolveHandler = CreateResolveHandler(settings, log);
        var detachResolveHandler = false;
        try
        {
            var content = File.ReadAllBytes(settings.Assembly);

            appDomain.ReflectionOnlyAssemblyResolve += resolveHandler;
            return new ROAssemblyExtended(Assembly.ReflectionOnlyLoad(content));
        }
        catch (Exception e) when
            (e is ArgumentNullException
               or PathTooLongException
               or DirectoryNotFoundException
               or IOException
               or UnauthorizedAccessException
               or FileNotFoundException
               or NotSupportedException
               or SecurityException)
        {
            detachResolveHandler = true;
            log.LogErrorFromException(e, true, true, settings.Assembly);
            return default;
        }
        catch
        {
            detachResolveHandler = true;
            throw;
        }
        finally
        {
            if (detachResolveHandler)
            {
                appDomain.ReflectionOnlyAssemblyResolve -= resolveHandler;
            }
        }
    }

    /// <summary>
    /// Creates an <see cref="ResolveEventHandler" /> that tries to look up a dependent assembly.
    /// </summary>
    /// <param name="settings">The conversion settings.</param>
    /// <param name="log">The task logging helper.</param>
    /// <returns>A new resolve event handler instance.</returns>
    private static ResolveEventHandler CreateResolveHandler(TypeLibConverterSettings settings, TaskLoggingHelper log)
    {
        Assembly? AssemblyResolveClosure(object? sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name);
            Assembly? assembly = default;
            if (TryResolveAssemblyFromSettings(assemblyName.Name ?? string.Empty, settings, log, out var assemblyPath))
            {
                assembly = Assembly.ReflectionOnlyLoadFrom(assemblyPath);
            }
            else
            {
                // Try load from GAC
                try
                {
                    assembly = Assembly.ReflectionOnlyLoad(assemblyName.FullName);
                }
                catch (FileNotFoundException)
                {
                    assembly = null;
                }
            }
            if (assembly is null)
            {
                log.LogWarning("Failed to resolve {0} from GAC or the following files: {1}", args.Name, string.Join(", ", settings.ASMPath));
            }
            return assembly;
        }

        return AssemblyResolveClosure;
    }
#endif

    /// <summary>
    /// Tries to resolve the managed assembly with the specified <paramref name="assemblyFileName"/> from the <paramref name="settings"/>.
    /// If this method returns <c>true</c>, the resolved file path will be written to <paramref name="assemblyPath"/>.
    /// </summary>
    /// <param name="assemblyFileName">The name of the assembly file to load (without extension).</param>
    /// <param name="settings">The settings for type conversions.</param>
    /// <param name="log">The logging helper.</param>
    /// <param name="assemblyPath">Path to resolved file.</param>
    /// <returns><c>true</c>, if the assembly could be resolved; <c>false</c> otherwise.</returns>
    private static bool TryResolveAssemblyFromSettings(string assemblyFileName, TypeLibConverterSettings settings, TaskLoggingHelper log, out string assemblyPath)
    {
        var validAssemblyExtensions = new string[] { ".dll", ".exe" };
        if (TryResolveAssemblyFromReferencedFiles(assemblyFileName, settings, log, validAssemblyExtensions, out assemblyPath)
            || TryResolveAssemblyFromAdjacentFiles(assemblyFileName, settings, log, validAssemblyExtensions, out assemblyPath))
        {
            return true;
        }

        assemblyPath = string.Empty;
        return false;
    }

    /// <summary>
    /// Tries to resolve the managed assembly with the specified <paramref name="assemblyFileName"/> from the <paramref name="settings"/> 
    /// using the <see cref="TypeLibConverterSettings.ASMPath" /> property to identify the assembly.
    /// If this method returns <c>true</c>, the resolved file path will be written to <paramref name="assemblyPath"/>.
    /// </summary>
    /// <param name="assemblyFileName">The name of the assembly file to load (without extension).</param>
    /// <param name="settings">The settings for type conversions.</param>
    /// <param name="log">The logging helper.</param>
    /// <param name="validAssemblyExtensions">Any extension that might be considered as valid assembly extension.</param>
    /// <param name="assemblyPath">Path to resolved file.</param>
    /// <returns><c>true</c>, if the assembly could be resolved; <c>false</c> otherwise.</returns>
    private static bool TryResolveAssemblyFromReferencedFiles(string assemblyFileName, TypeLibConverterSettings settings, TaskLoggingHelper log, string[] validAssemblyExtensions, out string assemblyPath)
    {
        log.LogMessage(MessageImportance.Low, "Trying to resolve assembly {0} from referenced files.", assemblyFileName);
        foreach (var path in settings.ASMPath)
        {
            var currentAssemblyFileName = Path.GetFileName(path) ?? string.Empty;
            log.LogMessage(MessageImportance.Low, "Current file is {0}. Maybe it matches.", path);
            foreach (var extension in validAssemblyExtensions)
            {
                var possibleFileName = assemblyFileName + extension;
                log.LogMessage(MessageImportance.Low, "Trying to resolve assembly {0} as {1}.", assemblyFileName, possibleFileName);
                if (StringComparer.InvariantCultureIgnoreCase.Equals(possibleFileName, currentAssemblyFileName)
                    && File.Exists(path))
                {
                    log.LogMessage(MessageImportance.Low, "Assembly resolved as {0}.", path);
                    assemblyPath = path;
                    return true;
                }
            }
        }

        assemblyPath = string.Empty;
        return false;
    }

    /// <summary>
    /// Tries to resolve the managed assembly with the specified <paramref name="assemblyFileName"/> from the <paramref name="settings"/> 
    /// using the <see cref="TypeLibConverterSettings.ASMPath" /> property to look up directories that might contain the file.
    /// If this method returns <c>true</c>, the resolved file path will be written to <paramref name="assemblyPath"/>.
    /// </summary>
    /// <param name="assemblyFileName">The name of the assembly file to load (without extension).</param>
    /// <param name="settings">The settings for type conversions.</param>
    /// <param name="log">The logging helper.</param>
    /// <param name="validAssemblyExtensions">Any extension that might be considered as valid assembly extension.</param>
    /// <param name="assemblyPath">Path to resolved file.</param>
    /// <returns><c>true</c>, if the assembly could be resolved; <c>false</c> otherwise.</returns>
    private static bool TryResolveAssemblyFromAdjacentFiles(string assemblyFileName, TypeLibConverterSettings settings, TaskLoggingHelper log, string[] validAssemblyExtensions, out string assemblyPath)
    {
        log.LogMessage(MessageImportance.Low, "Trying to resolve assembly {0} from adjacent files.", assemblyFileName);
        foreach (var path in settings.ASMPath)
        {
            var currentAssemblyFileName = Path.GetFileName(path) ?? string.Empty;
            var assemblyDirectoryName = Path.GetDirectoryName(path) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(assemblyDirectoryName))
            {
                continue;
            }

            log.LogMessage(MessageImportance.Low, "Current directory to look at is {0}. Maybe it matches.", assemblyDirectoryName);
            foreach (var extension in validAssemblyExtensions)
            {
                var possibleFileName = assemblyFileName + extension;
                var possibleAssemblyFilePath = Path.Combine(assemblyDirectoryName, possibleFileName);
                log.LogMessage(MessageImportance.Low, "Trying to resolve assembly {0} as {1}.", assemblyFileName, possibleAssemblyFilePath);
                if (File.Exists(possibleAssemblyFilePath))
                {
                    log.LogMessage(MessageImportance.Low, "Assembly resolved as {0}.", possibleAssemblyFilePath);
                    assemblyPath = possibleAssemblyFilePath;
                    return true;
                }
            }
        }

        assemblyPath = string.Empty;
        return false;
    }
}
