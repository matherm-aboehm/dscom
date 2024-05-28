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

namespace dSPACE.Runtime.InteropServices.Tests;

public class CompileReleaseFixture : IDisposable
{
    private sealed class TestContext : IDisposable
    {
        private readonly CompileReleaseFixture _owner;
        private int _lockTaken;
        public TestContext(CompileReleaseFixture owner)
        {
            _owner = owner;
        }
        public void Dispose()
        {
            // Dispose can be called from the thread which has the lock and from
            // another thread which had created the outer CompileReleaseFixture instance.
            // So use atomic operation here and release the lock only once on race condition.
            if (Interlocked.Exchange(ref _lockTaken, 0) == 1)
            {
                Monitor.Exit(this);
            }
        }
        public void Enter()
        {
            var lockTaken = false;
            Monitor.Enter(this, ref lockTaken);
            _lockTaken = lockTaken ? 1 : 0;
            _owner.PrepareTestDirectory();
        }
    }

    private readonly TestContext _testContext;

    public string Workdir { get; private set; } = string.Empty;

    public string CurrentDir { get; } = Environment.CurrentDirectory;

    public string DSComPath { get; private set; } = string.Empty;

    public string TestAssemblyPath { get; private set; } = string.Empty;

    public string TestAssemblyDependencyPath { get; private set; } = string.Empty;

    public CompileReleaseFixture()
    {
        var workdir = new DirectoryInfo(CurrentDir).Parent?.Parent?.Parent?.Parent?.Parent;
        if (workdir == null || !workdir.Exists)
        {
            throw new DirectoryNotFoundException("Workdir not found.");
        }

        Workdir = workdir.FullName;

#if DEBUG
        var configuration = "Debug";
#else
        var configuration = "Release";
#endif

#if NET8_0_OR_GREATER
        var frameworkVersion = "net8.0";
#elif NET6_0_OR_GREATER
        var frameworkVersion = "net6.0";
#else
        var frameworkVersion = "net48";
#endif

        // Path to descom.exe
        DSComPath = Path.Combine(Workdir, "src", "dscom.client", "bin", configuration, "net8.0", "dscom.exe");

        if (!File.Exists(DSComPath))
        {
            throw new FileNotFoundException("dscom.exe not found. Please build all projects before running the tests.", DSComPath);
        }

        // Path to dscom.demo assemblies
        TestAssemblyPath = GetPathToBuildOutput("dscom.test.assembly", configuration, frameworkVersion, "dSPACE.Runtime.InteropServices.Test.Assembly.dll");
        TestAssemblyDependencyPath = GetPathToBuildOutput("dscom.test.assembly.dependency", configuration, frameworkVersion, "dSPACE.Runtime.InteropServices.Test.Assembly.Dependency.dll");

        if (!File.Exists(TestAssemblyPath))
        {
            throw new FileNotFoundException($"The test assembly {TestAssemblyPath} not found. Please build all projects before running the tests.", TestAssemblyPath);
        }

        if (!File.Exists(TestAssemblyDependencyPath))
        {
            throw new FileNotFoundException($"The test assembly {TestAssemblyDependencyPath} not found. Please build all projects before running the tests.", TestAssemblyDependencyPath);
        }

        _testContext = new TestContext(this);
    }

    private string GetPathToBuildOutput(string projectName, string configuration, string targetFramework, string fileName)
    {
        var outDirWithoutTfm = Path.Combine(Workdir, "src", projectName, "bin", configuration);
#if !NET48
        var outDirSpecificPlatform = Path.Combine(outDirWithoutTfm, targetFramework + "-windows");
        if (Directory.Exists(outDirSpecificPlatform))
        {
            return Path.Combine(outDirSpecificPlatform, fileName);
        }
#endif
        return Path.Combine(outDirWithoutTfm, targetFramework, fileName);
    }

    private static void DeleteFilesSyncBlocking(string dirpath, string filter)
    {
        using FileSystemWatcher watcher = new(dirpath, filter);
        using CountdownEvent cde = new(1);
        watcher.Deleted += (s, e) =>
        {
            cde.Signal();
        };
        watcher.EnableRaisingEvents = true;
        foreach (var file in Directory.EnumerateFiles(dirpath, filter))
        {
            cde.AddCount();
            File.Delete(file);
        }
        cde.Signal();
        cde.Wait();
    }

    public void PrepareTestDirectory()
    {
        RetryHandler.Retry(() =>
        {
            DeleteFilesSyncBlocking(CurrentDir, "*.tlb");
        }, new[] { typeof(UnauthorizedAccessException) });
        RetryHandler.Retry(() =>
        {
            DeleteFilesSyncBlocking(CurrentDir, "*.yaml");
        }, new[] { typeof(UnauthorizedAccessException) });
    }

    public IDisposable GetPreparedTestDirectoryContext()
    {
        _testContext.Enter();
        return _testContext;
    }

    public void Dispose()
    {
        _testContext.Dispose();
        GC.SuppressFinalize(this);
    }
}
