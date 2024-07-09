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
            try
            {
                _owner.PrepareTestDirectory();
            }
            catch
            {
                if (lockTaken)
                {
                    // don't risk a deadlock on exception, just release the lock
                    // and let another thread take the lock.
                    _lockTaken = 0;
                    Monitor.Exit(this);
                }
                throw;
            }
        }
    }

    private readonly TestContext _testContext;

    public string Workdir { get; private set; } = string.Empty;

    public string CurrentDir { get; } = Environment.CurrentDirectory;

    public string DSComPath { get; private set; } = string.Empty;

    public string DemoProjectAssembly1Path { get; private set; } = string.Empty;

    public string DemoProjectAssembly2Path { get; private set; } = string.Empty;

    public string DemoProjectAssembly3Path { get; private set; } = string.Empty;

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

        // Path to descom.exe
        DSComPath = Path.Combine(Workdir, "src", "dscom.client", "bin", configuration, "net6.0", "dscom.exe");

        // Path to dscom.demo assemblies
        DemoProjectAssembly1Path = Path.Combine(Workdir, "src", "dscom.demo", "assembly1", "bin", configuration, "net6.0", "dSPACE.Runtime.InteropServices.DemoAssembly1.dll");
        DemoProjectAssembly2Path = Path.Combine(Workdir, "src", "dscom.demo", "assembly2", "bin", configuration, "net6.0", "dSPACE.Runtime.InteropServices.DemoAssembly2.dll");
        DemoProjectAssembly3Path = Path.Combine(Workdir, "src", "dscom.demo", "assembly3", "bin", configuration, "net6.0", "dSPACE.Runtime.InteropServices.DemoAssembly3.dll");

        _testContext = new TestContext(this);
    }

    private static void DeleteFilesSyncBlocking(string dirpath, string filter)
    {
        using FileSystemWatcher watcher = new(dirpath, filter);
        using CountdownEvent cde = new(1);
        watcher.Renamed += (s, e) =>
        {
            cde.Signal();
        };
        watcher.EnableRaisingEvents = true;
        foreach (var file in Directory.GetFiles(dirpath, filter))
        {
            cde.AddCount();
            //HINT: Some anti-virus can still block deletion of file although its deletion
            // event was raised already, so use rename operation followed by a delete here.
            var tmpfile = Path.Combine(dirpath, $"{Guid.NewGuid}.tmp");
            File.Move(file, tmpfile);
            var retry = true;
            var retrycount = 0;
            do
            {
                try
                {
                    File.Delete(tmpfile);
                    retry = false;
                }
                catch (IOException e) when
                    (retrycount < 100 && e.HResult == unchecked((int)0x80070020) /* HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION) */)
                {
                    Thread.Sleep(100);
                    retrycount++;
                }
            } while (retry);
        }
        cde.Signal();
        var timeout = !cde.Wait(30000);
        if (timeout)
        {
            throw new TimeoutException($"Preparing the test directory failed. Some files with pattern '{filter}' can't be deleted.");
        }
    }

    public void PrepareTestDirectory()
    {
        DeleteFilesSyncBlocking(CurrentDir, "*.tlb");
        DeleteFilesSyncBlocking(CurrentDir, "*.yaml");
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
