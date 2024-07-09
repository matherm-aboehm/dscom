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

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.Loader;
using System.Text;
using Xunit;

namespace dSPACE.Runtime.InteropServices.Tests;

/// <summary>
/// Provides the base implementation for running CLI tests.
/// </summary>
/// <remarks>The CLI tests are not available for .NET Framework</remarks>
public abstract class CLITestBase : IClassFixture<CompileReleaseFixture>, IDisposable
{
    protected const string ErrorNoCommandOrOptions = "Required command was not provided.";

    internal record struct ProcessOutput(string StdOut, string StdErr, int ExitCode);

    internal CompileReleaseFixture CompileFixture { get; }

    internal string DSComPath { get; set; } = string.Empty;

    internal string? TestAssemblyTemporyTlbFilePath { get; private set; }

    internal string? TestAssemblyTemporyYamlFilePath { get; private set; }

    internal string TestAssemblyPath { get; }

    internal string TestAssemblyDependencyPath { get; }

    internal string? TestAssemblyDependencyTemporyTlbFilePath { get; private set; }

    internal string? TemporaryTestDirectoryPath { get; private set; }

    public CLITestBase(CompileReleaseFixture compileFixture)
    {
        CompileFixture = compileFixture;
        DSComPath = compileFixture.DSComPath;
        TestAssemblyPath = compileFixture.TestAssemblyPath;
        TestAssemblyDependencyPath = compileFixture.TestAssemblyDependencyPath;
    }

    [MemberNotNull(nameof(TemporaryTestDirectoryPath))]
    [MemberNotNull(nameof(TestAssemblyTemporyTlbFilePath))]
    [MemberNotNull(nameof(TestAssemblyTemporyYamlFilePath))]
    [MemberNotNull(nameof(TestAssemblyDependencyTemporyTlbFilePath))]
    internal void PrepareTemporaryTestDirectory()
    {
        TemporaryTestDirectoryPath = Path.Combine(CompileFixture.CurrentDir, Guid.NewGuid().ToString());

        Directory.CreateDirectory(TemporaryTestDirectoryPath);

        TestAssemblyTemporyTlbFilePath = Path.Combine(TemporaryTestDirectoryPath, Path.GetFileNameWithoutExtension(TestAssemblyPath) + ".tlb");
        TestAssemblyTemporyYamlFilePath = TestAssemblyTemporyTlbFilePath.Replace(".tlb", ".yaml");
        TestAssemblyDependencyTemporyTlbFilePath = Path.Combine(TemporaryTestDirectoryPath, Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath) + ".tlb");
    }

    internal void KeepTemporaryTestDirectory()
    {
        TemporaryTestDirectoryPath = null;
    }

    internal static ProcessOutput Execute(string filename, params string[] args)
    {
        var processOutput = new ProcessOutput();
        using var process = new Process();
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true;
        // To avoid deadlocks, need async read from both streams to keep the buffers empty, see:
        //https://stackoverflow.com/questions/139593/processstartinfo-hanging-on-waitforexit-why
        var sbErr = new StringBuilder();
        var sbOut = new StringBuilder();
        using var mreErr = new ManualResetEvent(false);
        using var mreOut = new ManualResetEvent(false);
        static void AppendLineFromEventArgs(DataReceivedEventArgs e, StringBuilder sb, ManualResetEvent mre)
        {
            if (e.Data is not null)
            {
                sb.Append(e.Data);
            }
            else // redirected stream is closed
            {
                mre.Set();
            }
        }
        process.ErrorDataReceived += new DataReceivedEventHandler((sender, e) => AppendLineFromEventArgs(e, sbErr, mreErr));
        process.OutputDataReceived += new DataReceivedEventHandler((sender, e) => AppendLineFromEventArgs(e, sbOut, mreOut));
        process.StartInfo.FileName = filename;
        process.StartInfo.Arguments = string.Join(" ", args);
        process.Start();

        process.BeginErrorReadLine();
        process.BeginOutputReadLine();
        var timeout = !process.WaitForExit(60000);
        if (timeout)
        {
            process.Kill();
        }
        mreOut.WaitOne(10000);
        mreErr.WaitOne(10000);
        process.WaitForExit();
        processOutput.StdOut = sbOut.ToString();
        processOutput.StdErr = sbErr.ToString();
        processOutput.ExitCode = timeout ? -1 : process.ExitCode;

        return processOutput;
    }

    /// <summary>
    /// When running tests that involves the <c>tlbembed</c> command or <c>tlbexport</c> 
    /// command with the <c>--embed</c> switch enabled, the test should call the function 
    /// to handle the difference between .NET 4.8 which normally would expect COM objects 
    /// to be defined in the generated assembly versus .NET 5+ where COM objects will be 
    /// present in a <c>*.comhost.dll</c> instead.
    /// </summary>
    /// <param name="sourceAssemblyFile">The path to the source assembly from where the COM objects are defined.</param>
    internal static string GetEmbeddedPath(string sourceAssemblyFile)
    {
        var embedFile = Path.Combine(Path.GetDirectoryName(sourceAssemblyFile) ?? string.Empty, Path.GetFileNameWithoutExtension(sourceAssemblyFile) + ".comhost" + Path.GetExtension(sourceAssemblyFile));
        Assert.True(File.Exists(embedFile));
        return embedFile;
    }

    internal static void TlbCommandShouldNotRegisterTypeLib(string command, params string[] assemblyPaths)
    {
        var loadContext = new AssemblyLoadContext("load-ctx-clitest", true);
        try
        {
            foreach (var assemblyPath in assemblyPaths)
            {
                //var anDemoProjectAssembly2 = AssemblyName.GetAssemblyName(TestAssemblyDependencyPath);
                var identifier = loadContext.LoadFromAssemblyPath(assemblyPath).GetLibIdentifier();
                //Assert.Equal("6dfe7b4e-9400-3aae-ac08-c120a280ef6b", identifier.LibID.ToString());
                var hr = OleAut32.QueryPathOfRegTypeLib(identifier.LibID, identifier.MajorVersion, identifier.MinorVersion,
                    Constants.LCID_NEUTRAL, out var typeLibPathFromRegistry);
                Assert.True(0 != hr, $"{command} should not register any type lib, but it did with path '{typeLibPathFromRegistry}'");
            }
        }
        finally
        {
            loadContext.Unload();
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (!string.IsNullOrEmpty(TemporaryTestDirectoryPath) &&
                Directory.Exists(TemporaryTestDirectoryPath))
            {
                RetryHandler.Retry(() =>
                {
                    Directory.Delete(TemporaryTestDirectoryPath, true);
                }, new[] { typeof(UnauthorizedAccessException) });
            }
        }
    }
}
