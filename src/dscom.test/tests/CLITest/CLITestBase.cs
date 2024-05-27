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
        var sb = new StringBuilder();
        process.ErrorDataReceived += new DataReceivedEventHandler((sender, e) => { sb.Append(e.Data); });
        process.StartInfo.FileName = filename;
        process.StartInfo.Arguments = string.Join(" ", args);
        process.Start();

        process.BeginErrorReadLine();
        process.WaitForExit();
        processOutput.StdOut = process.StandardOutput.ReadToEnd();
        processOutput.StdErr = sb.ToString();
        processOutput.ExitCode = process.ExitCode;

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
