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
/// The basic CLI tests to run.
/// </summary>
[Collection("CLI Tests")]
public class CLITest : CLITestBase
{
    public CLITest(CompileReleaseFixture compileFixture) : base(compileFixture) { }

    [Fact]
    public void CallWithoutCommandOrOption_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath);

        Assert.Equal(1, result.ExitCode);
        Assert.Contains(ErrorNoCommandOrOptions, result.StdErr.Trim());
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void CallWithoutCommandABC_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath, "ABC");

        Assert.Equal(1, result.ExitCode);
        Assert.Contains(ErrorNoCommandOrOptions, result.StdErr.Trim());
        Assert.Contains("Unrecognized command or argument 'ABC'", result.StdErr.Trim());
    }

    [Fact]
    public void CallWithVersionOption_VersionIsAssemblyInformationalVersionAttributeValue()
    {
        var assemblyInformationalVersion = typeof(TypeLibConverter).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        Assert.NotNull(assemblyInformationalVersion);
        var versionFromLib = assemblyInformationalVersion!.InformationalVersion;

        var result = Execute(DSComPath, "--version");
        Assert.Equal(0, result.ExitCode);
        Assert.StartsWith(versionFromLib, result.StdOut.Trim());
    }

    [Fact]
    public void CallWithHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "--help");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void TlbExportAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbexport", "--help");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void TlbDumpAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbdump", "--help");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void TlbRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbregister", "--help");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void TlbUnRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbunregister", "--help");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("Description", result.StdOut.Trim());
    }

    [Fact]
    public void TlbUnRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbunregister", "abc");
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("not found", result.StdErr.Trim());
    }

    [Fact]
    public void TlbRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbregister", "abc");
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("not found", result.StdErr.Trim());
    }

    [Fact]
    public void TlbDumpAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbdump", "abc");
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("not found", result.StdErr.Trim());
    }

    [Fact]
    public void TlbExportAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbexport", "abc");
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("not found", result.StdErr.Trim());
    }

    [Fact]
    public void TlbExportAndDemoAssemblyAndCallWithTlbDump_ExitCodeIs0AndTlbIsAvailableAndValid()
    {
        PrepareTemporaryTestDirectory();
        var dependentTlbPath = TestAssemblyDependencyTemporyTlbFilePath;

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", TestAssemblyTemporyTlbFilePath);
        Assert.Equal(0, result.ExitCode);
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.True(File.Exists(dependentTlbPath), $"File {dependentTlbPath} should be available.");

        var dumpResult = Execute(DSComPath, "tlbdump", TestAssemblyTemporyTlbFilePath, "--out", TestAssemblyTemporyYamlFilePath, "--tlbrefpath", TemporaryTestDirectoryPath);
        // xunit does not support Assert.Equal with custom message, so the following is the only way
        // see: https://github.com/xunit/xunit/issues/350
        Assert.True(0 == dumpResult.ExitCode, $"because it should succeed. ExitCode: {dumpResult.ExitCode} Error: {dumpResult.StdErr}. Output: {dumpResult.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyYamlFilePath), $"File {TestAssemblyTemporyYamlFilePath} should be available.");
    }

    [Fact]
    public void TlbExportAndEmbedAssembly_ExitCodeIs0AndTlbIsEmbeddedAndValid()
    {
        PrepareTemporaryTestDirectory();
        var embedPath = GetEmbeddedPath(TestAssemblyPath);

        var dependentTlbPath = TestAssemblyDependencyTemporyTlbFilePath;

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, $"--embed {embedPath}", "--out", TestAssemblyTemporyTlbFilePath);
        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");

        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.True(File.Exists(dependentTlbPath), $"File {dependentTlbPath} should be available.");

        OleAut32.LoadTypeLibEx(embedPath, REGKIND.NONE, out var embeddedTypeLib);
        using (embeddedTypeLib.AsDisposableComObject())
        {
            OleAut32.LoadTypeLibEx(TestAssemblyTemporyTlbFilePath, REGKIND.NONE, out var sourceTypeLib);
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
    public void TlbExportCreateMissingDependentTLBsFalseAndOverrideTlbId_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var parameters = new[] { "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "false", "--overridetlbid", "12345678-1234-1234-1234-123456789012" };

        var result = Execute(DSComPath, parameters);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        var fileName = Path.GetFileNameWithoutExtension(TestAssemblyPath);

        Assert.DoesNotContain($"{fileName} does not have a type library", result.StdOut);
        Assert.DoesNotContain($"{fileName} does not have a type library", result.StdErr);

        Assert.True(File.Exists(tlbFilePath), $"File {tlbFilePath} should be available.");
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsFalse_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        PrepareTemporaryTestDirectory();
        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "false", "--out", TestAssemblyTemporyTlbFilePath);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.False(File.Exists(TestAssemblyDependencyTemporyTlbFilePath), $"File {TestAssemblyDependencyTemporyTlbFilePath} should not be available.");
        Assert.Contains("auto generation of dependent type libs is disabled", result.StdErr);
        Assert.Contains(Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath), result.StdErr);
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsTrue_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        PrepareTemporaryTestDirectory();
        var dependentTlbPath = TestAssemblyDependencyTemporyTlbFilePath;

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "true", "--out", TestAssemblyTemporyTlbFilePath);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.True(File.Exists(dependentTlbPath), $"File {dependentTlbPath} should be available.");
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsNoValue_ExitCodeIs0()
    {
        PrepareTemporaryTestDirectory();
        var dependentTlbPath = TestAssemblyDependencyTemporyTlbFilePath;

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "--out", TestAssemblyTemporyTlbFilePath);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.True(File.Exists(dependentTlbPath), $"File {dependentTlbPath} should be available.");
    }

    [Fact]
    public void TlbExportAndOptionSilent_StdOutAndStdErrIsEmpty()
    {
        PrepareTemporaryTestDirectory();
        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--silent", "--out", TestAssemblyTemporyTlbFilePath);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.Empty(result.StdOut.Trim());
        Assert.Empty(result.StdErr.Trim());
    }

    [Fact]
    public void TlbExportAndOptionSilenceTX801311A6andTX0013116F_StdOutAndStdErrIsEmpty()
    {
        PrepareTemporaryTestDirectory();
        // HACK: new test assemblies needs also to silence TX8013117D, TX800288C6,
        // as long as they are referencing WPF libs
        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--silence", "TX801311A6", "--silence", "TX0013116F", "--silence", "TX8013117D", "--silence", "TX800288C6", "--out", TestAssemblyTemporyTlbFilePath);

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");
        Assert.Empty(result.StdOut.Trim());
        Assert.Empty(result.StdErr.Trim());
    }

    [Fact]
    public void TlbExportAndOptionOverrideTLBId_TLBIdIsCorrect()
    {
        PrepareTemporaryTestDirectory();
        var guid = Guid.NewGuid();

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", TestAssemblyTemporyTlbFilePath, "--overridetlbid", guid.ToString());

        Assert.True(0 == result.ExitCode, $"because it should succeed. ExitCode: {result.ExitCode} Error: {result.StdErr}. Output: {result.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyTlbFilePath), $"File {TestAssemblyTemporyTlbFilePath} should be available.");

        var dumpResult = Execute(DSComPath, "tlbdump", TestAssemblyTemporyTlbFilePath, "/out", TestAssemblyTemporyYamlFilePath, "/tlbrefpath", TemporaryTestDirectoryPath);
        Assert.True(0 == dumpResult.ExitCode, $"because it should succeed. Error: ${dumpResult.StdErr}. Output: ${dumpResult.StdOut}");
        Assert.True(File.Exists(TestAssemblyTemporyYamlFilePath), $"File {TestAssemblyTemporyYamlFilePath} should be available.");

        var yamlContent = File.ReadAllText(TestAssemblyTemporyYamlFilePath);
        Assert.Contains($"guid: {guid}", yamlContent);
    }
}
