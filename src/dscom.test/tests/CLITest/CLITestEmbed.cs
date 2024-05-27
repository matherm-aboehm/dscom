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
/// <summary>
/// The tests for embeds are performed as a separate class due to the additional setup
/// required and the need for creating a TLB file as part of export prior to testing
/// the embed functionality itself. There are parallelization issues with running them,
/// even within the same class. Thus, the separate class has additional setup during the
/// constructor to perform the export once and ensure the process relating to export is
/// disposed with before attempting to test the embed functionality.
/// </summary>
[Collection("CLI Tests")]
public class CLITestEmbed : CLITestBase
{
    internal string TlbFilePath { get; }

    internal string DependentTlbPath { get; }

    public CLITestEmbed(CompileReleaseFixture compileFixture) : base(compileFixture)
    {
        PrepareTemporaryTestDirectory();
        TlbFilePath = TestAssemblyTemporyTlbFilePath;

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", TlbFilePath);
        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");

        result = Execute(DSComPath, "tlbexport", TestAssemblyDependencyPath);
        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");

        DependentTlbPath = TestAssemblyDependencyTemporyTlbFilePath;

        Assert.True(File.Exists(TlbFilePath), $"File {TlbFilePath} should be available.");
        Assert.True(File.Exists(DependentTlbPath), $"File {DependentTlbPath} should be available.");

        // This is necessary to ensure the process from previous Execute for the export command has completely disposed before running tests.
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static void AssertOnLoadableAndEqualTypeLibs(string embeddedTypeLibPath, string sourceTypeLibPath)
    {
        OleAut32.LoadTypeLibEx(embeddedTypeLibPath, REGKIND.NONE, out var embeddedTypeLib);
        using (embeddedTypeLib.AsDisposableComObject())
        {
            OleAut32.LoadTypeLibEx(sourceTypeLibPath, REGKIND.NONE, out var sourceTypeLib);
            using (sourceTypeLib.AsDisposableComObject())
            {
                try
                {

                    embeddedTypeLib.GetDocumentation(-1, out var embeddedTypeLibName, out _, out _, out _);
                    sourceTypeLib.GetDocumentation(-1, out var sourceTypeLibName, out _, out _, out _);

                    Assert.Equal(sourceTypeLibName, embeddedTypeLibName);
                }
                finally
                {
                    embeddedTypeLib = null;
                    sourceTypeLib = null;
                }
            }
        }
        // This is necessary to ensure the com objects have completely disposed before trying to delete the temp folder.
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    [Fact]
    public void TlbEmbedAssembly_ExitCodeIs0AndTlbIsEmbeddedAndValid()
    {
        var embedPath = GetEmbeddedPath(TestAssemblyPath);

        var result = Execute(DSComPath, "tlbembed", TlbFilePath, embedPath);
        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");

        AssertOnLoadableAndEqualTypeLibs(embedPath, TlbFilePath);
    }

    [Fact]
    public void TlbEmbedAssemblyWithArbitraryIndex_ExitCodeIs0AndTlbIsEmbeddedAndValid()
    {
        var embedPath = GetEmbeddedPath(TestAssemblyPath);
        var result = Execute(DSComPath, "tlbembed", TlbFilePath, embedPath, "--index 2");
        Assert.Equal(0, result.ExitCode);

        AssertOnLoadableAndEqualTypeLibs(embedPath + "\\2", TlbFilePath);
    }

    [Fact]
    public void TlbEmbedAssemblyWithArbitraryTlbAndArbitraryIndex_ExitCodeIs0AndTlbIsEmbeddedAndValid()
    {
        var embedPath = GetEmbeddedPath(TestAssemblyPath);
        var result = Execute(DSComPath, "tlbembed", TlbFilePath, embedPath, "--index 3");
        Assert.Equal(0, result.ExitCode);

        AssertOnLoadableAndEqualTypeLibs(embedPath + "\\3", TlbFilePath);
    }

    [Fact]
    public void TlbEmbedAssemblyWithMultipleTypeLibraries_ExitCodeAre0AndTlbsAreEmbeddedAndValid()
    {
        var embedPath = GetEmbeddedPath(TestAssemblyPath);
        var result = Execute(DSComPath, "tlbembed", TlbFilePath, embedPath);
        Assert.Equal(0, result.ExitCode);

        result = Execute(DSComPath, "tlbembed", DependentTlbPath, TestAssemblyPath, "--index 2");
        Assert.Equal(0, result.ExitCode);

        AssertOnLoadableAndEqualTypeLibs(embedPath, TlbFilePath);
        AssertOnLoadableAndEqualTypeLibs(TestAssemblyPath + "\\2", DependentTlbPath);
    }
}
