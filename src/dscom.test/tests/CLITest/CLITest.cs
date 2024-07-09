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

    [StaFact]
    public void CallWithoutCommandOrOption_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath);

        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Be(ErrorNoCommandOrOptions);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void CallWithoutCommandABC_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath, "ABC");

        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain(ErrorNoCommandOrOptions);
        result.StdErr.Trim().Should().Contain("Unrecognized command or argument 'ABC'");
    }

    [StaFact]
    public void CallWithVersionOption_VersionIsAssemblyInformationalVersionAttributeValue()
    {
        var assemblyInformationalVersion = typeof(TypeLibConverter).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        assemblyInformationalVersion.Should().NotBeNull("AssemblyInformationalVersionAttribute is not set");
        var versionFromLib = assemblyInformationalVersion!.InformationalVersion;

        var result = Execute(DSComPath, "--version");
        result.ExitCode.Should().Be(0);
        versionFromLib.Should().StartWith(result.StdOut.Trim());
    }

    [StaFact]
    public void CallWithHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void TlbExportAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbexport", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void TlbDumpAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbdump", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void TlbRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbregister", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void TlbUnRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbunregister", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [StaFact]
    public void TlbUnRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbunregister", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [StaFact]
    public void TlbRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbregister", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [StaFact]
    public void TlbDumpAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbdump", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [StaFact]
    public void TlbExportAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbexport", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [StaFact]
    public void TlbExportAndDemoAssemblyAndCallWithTlbDump_ExitCodeIs0AndTlbIsAvailableAndValid()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var yamlFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.yaml";

        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var yamlFilePath = Path.Combine(Environment.CurrentDirectory, yamlFileName);
        var dependentTlbPath = $"{Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath)}.tlb";

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath);
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);

        var dumpResult = Execute(DSComPath, "tlbdump", tlbFilePath);
        dumpResult.ExitCode.Should().Be(0);

        File.Exists(yamlFilePath).Should().BeTrue($"File {yamlFilePath} should be available.");

        if (Environment.OSVersion.Version == new Version(10, 0, 20348, 0)) //Windows Server 2022 Version 21H2
        {
            var unregisterResult = Execute(DSComPath, "tlbunregister", dependentTlbPath);
            unregisterResult.ExitCode.Should().Be(0);
        }
        TlbCommandShouldNotRegisterTypeLib("tlbdump", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportAndEmbedAssembly_ExitCodeIs0AndTlbIsEmbeddedAndValid()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var embedPath = GetEmbeddedPath(TestAssemblyPath);

        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";

        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = $"{Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath)}.tlb";

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, $"--embed {embedPath}");
        result.ExitCode.Should().Be(0, $"because it should succeed. Error: ${result.StdErr}. Output: ${result.StdOut}");

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);

        OleAut32.LoadTypeLibEx(embedPath, REGKIND.NONE, out var embeddedTypeLib);
        OleAut32.LoadTypeLibEx(tlbFilePath, REGKIND.NONE, out var sourceTypeLib);

        embeddedTypeLib.GetDocumentation(-1, out var embeddedTypeLibName, out _, out _, out _);
        sourceTypeLib.GetDocumentation(-1, out var sourceTypeLibName, out _, out _, out _);

        Assert.Equal(sourceTypeLibName, embeddedTypeLibName);
    }

    [StaFact]
    public void TlbExportCreateMissingDependentTLBsFalseAndOverrideTlbId_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        var parameters = new[] { "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "false", "--overridetlbid", "12345678-1234-1234-1234-123456789012" };

        var result = Execute(DSComPath, parameters);
        result.ExitCode.Should().Be(0);
        var fileName = Path.GetFileNameWithoutExtension(TestAssemblyPath);

        result.StdOut.Should().NotContain($"{fileName} does not have a type library");
        result.StdErr.Should().NotContain($"{fileName} does not have a type library");

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath);
    }

    [StaFact]
    public void TlbExportCreateMissingDependentTLBsFalse_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "false");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeFalse($"File {dependentTlbPath} should not be available.");

        result.StdErr.Should().Contain("auto generation of dependent type libs is disabled");
        result.StdErr.Should().Contain(Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath));

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportCreateMissingDependentTLBsTrue_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs", "true");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportCreateMissingDependentTLBsNoValue_ExitCodeIs0()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyPath)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(TestAssemblyDependencyPath)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--createmissingdependenttlbs");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportAndOptionSilent_StdOutAndStdErrIsEmpty()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Guid.NewGuid()}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", tlbFileName, "--silent");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        result.StdOut.Trim().Should().BeNullOrEmpty();
        result.StdErr.Trim().Should().BeNullOrEmpty();

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportAndOptionSilenceTX801311A6andTX0013116F_StdOutAndStdErrIsEmpty()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Guid.NewGuid()}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        // HACK: new test assemblies needs also to silence TX8013117D, TX800288C6,
        // as long as they are referencing WPF libs
        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", tlbFileName, "--silence", "TX801311A6", "--silence", "TX0013116F", "--silence", "TX8013117D", "--silence", "TX800288C6");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        result.StdOut.Trim().Should().BeNullOrEmpty();
        result.StdErr.Trim().Should().BeNullOrEmpty();

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);
    }

    [StaFact]
    public void TlbExportAndOptionOverrideTLBId_TLBIdIsCorrect()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var guid = Guid.NewGuid().ToString();

        var tlbFileName = $"{guid}.tlb";
        var yamlFileName = $"{guid}.yaml";

        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var yamlFilePath = Path.Combine(Environment.CurrentDirectory, yamlFileName);

        var result = Execute(DSComPath, "tlbexport", TestAssemblyPath, "--out", tlbFilePath, "--overridetlbid", guid);

        result.ExitCode.Should().Be(0);
        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", TestAssemblyPath, TestAssemblyDependencyPath);

        var dumpResult = Execute(DSComPath, "tlbdump", tlbFilePath, "/out", yamlFilePath);
        dumpResult.ExitCode.Should().Be(0);

        File.Exists(yamlFilePath).Should().BeTrue($"File {yamlFilePath} should be available.");

        var yamlContent = File.ReadAllText(yamlFilePath);
        yamlContent.Should().Contain($"guid: {guid}");

        TlbCommandShouldNotRegisterTypeLib("tlbdump", TestAssemblyPath, TestAssemblyDependencyPath);
    }
}
