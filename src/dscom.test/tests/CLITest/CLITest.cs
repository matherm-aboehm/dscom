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
using System.Runtime.Loader;
using System.Text;

namespace dSPACE.Runtime.InteropServices.Tests;

// The CLI is not available for .NET Framework
public class CLITest : IClassFixture<CompileReleaseFixture>
{
    private const string ErrorNoCommandOrOptions = "Required command was not provided.";

    internal record struct ProcessOutput(string StdOut, string StdErr, int ExitCode);

    internal CompileReleaseFixture CompileFixture { get; }

    internal string DSComPath { get; set; } = string.Empty;

    internal string DemoProjectAssembly1Path { get; }

    internal string DemoProjectAssembly2Path { get; }

    internal string DemoProjectAssembly3Path { get; }

    public CLITest(CompileReleaseFixture compileFixture)
    {
        CompileFixture = compileFixture;
        DSComPath = compileFixture.DSComPath;
        DemoProjectAssembly1Path = compileFixture.DemoProjectAssembly1Path;
        DemoProjectAssembly2Path = compileFixture.DemoProjectAssembly2Path;
        DemoProjectAssembly3Path = compileFixture.DemoProjectAssembly3Path;
    }

    [Fact]
    public void CallWithoutCommandOrOption_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath);

        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Be(ErrorNoCommandOrOptions);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void CallWithoutCommandABC_ExitCodeIs1AndStdOutIsHelpStringAndStdErrIsUsed()
    {
        var result = Execute(DSComPath, "ABC");

        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain(ErrorNoCommandOrOptions);
        result.StdErr.Trim().Should().Contain("Unrecognized command or argument 'ABC'");
    }

    [Fact]
    public void CallWithVersionOption_VersionIsAssemblyInformationalVersionAttributeValue()
    {
        var assemblyInformationalVersion = typeof(TypeLibConverter).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
        assemblyInformationalVersion.Should().NotBeNull("AssemblyInformationalVersionAttribute is not set");
        var versionFromLib = assemblyInformationalVersion!.InformationalVersion;

        var result = Execute(DSComPath, "--version");
        result.ExitCode.Should().Be(0);
        versionFromLib.Should().StartWith(result.StdOut.Trim());
    }

    [Fact]
    public void CallWithHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void TlbExportAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbexport", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void TlbDumpAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbdump", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void TlbRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbregister", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void TlbUnRegisterAndHelpOption_StdOutIsHelpStringAndExitCodeIsZero()
    {
        var result = Execute(DSComPath, "tlbunregister", "--help");
        result.ExitCode.Should().Be(0);
        result.StdOut.Trim().Should().Contain("Description");
    }

    [Fact]
    public void TlbUnRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbunregister", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [Fact]
    public void TlbRegisterAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbregister", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [Fact]
    public void TlbDumpAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbdump", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    [Fact]
    public void TlbExportAndFileNotExist_StdErrIsFileNotFoundAndExitCodeIs1()
    {
        var result = Execute(DSComPath, "tlbexport", "abc");
        result.ExitCode.Should().Be(1);
        result.StdErr.Trim().Should().Contain("not found");
    }

    private static void TlbCommandShouldNotRegisterTypeLib(string command, params string[] assemblyPaths)
    {
        var loadContext = new AssemblyLoadContext("load-ctx-clitest", true);
        try
        {
            foreach (var assemblyPath in assemblyPaths)
            {
                //var anDemoProjectAssembly2 = AssemblyName.GetAssemblyName(DemoProjectAssembly2Path);
                var identifier = loadContext.LoadFromAssemblyPath(assemblyPath).GetLibIdentifier();
                //identifier.LibID.Should().Be("6dfe7b4e-9400-3aae-ac08-c120a280ef6b");
                var hr = OleAut32.QueryPathOfRegTypeLib(identifier.LibID, identifier.MajorVersion, identifier.MinorVersion,
                    Constants.LCID_NEUTRAL, out var typeLibPathFromRegistry);
                hr.Should().NotBe(0, "{0} should not register any type lib, but it did with path '{1}'", command, typeLibPathFromRegistry);
            }
        }
        finally
        {
            loadContext.Unload();
        }
    }

    [Fact]
    public void TlbExportAndDemoAssemblyAndCallWithTlbDump_ExitCodeIs0AndTlbIsAvailableAndValid()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly1Path)}.tlb";
        var yamlFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly1Path)}.yaml";

        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var yamlFilePath = Path.Combine(Environment.CurrentDirectory, yamlFileName);
        var dependentTlbPath = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly2Path)}.tlb";

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path);
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);

        var dumpResult = Execute(DSComPath, "tlbdump", tlbFilePath);
        dumpResult.ExitCode.Should().Be(0);

        File.Exists(yamlFilePath).Should().BeTrue($"File {yamlFilePath} should be available.");

        if (Environment.OSVersion.Version == new Version(10, 0, 20348, 0)) //Windows Server 2022 Version 21H2
        {
            var unregisterResult = Execute(DSComPath, "tlbunregister", dependentTlbPath);
            unregisterResult.ExitCode.Should().Be(0);
        }
        TlbCommandShouldNotRegisterTypeLib("tlbdump", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsFalseAndOverrideTlbId_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly3Path)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        var parameters = new[] { "tlbexport", DemoProjectAssembly3Path, "--createmissingdependenttlbs", "false", "--overridetlbid", "12345678-1234-1234-1234-123456789012" };

        var result = Execute(DSComPath, parameters);
        result.ExitCode.Should().Be(0);
        var fileName = Path.GetFileNameWithoutExtension(DemoProjectAssembly3Path);

        result.StdOut.Should().NotContain($"{fileName} does not have a type library");
        result.StdErr.Should().NotContain($"{fileName} does not have a type library");

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly3Path);
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsFalse_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly1Path)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly2Path)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--createmissingdependenttlbs", "false");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeFalse($"File {dependentTlbPath} should not be available.");

        result.StdErr.Should().Contain("auto generation of dependent type libs is disabled");
        result.StdErr.Should().Contain(Path.GetFileNameWithoutExtension(DemoProjectAssembly2Path));

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsTrue_ExitCodeIs0AndTlbIsAvailableAndDependentTlbIsNot()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly1Path)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly2Path)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--createmissingdependenttlbs", "true");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportCreateMissingDependentTLBsNoValue_ExitCodeIs0()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly1Path)}.tlb";
        var dependentFileName = $"{Path.GetFileNameWithoutExtension(DemoProjectAssembly2Path)}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var dependentTlbPath = Path.Combine(Environment.CurrentDirectory, dependentFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--createmissingdependenttlbs");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");
        File.Exists(dependentTlbPath).Should().BeTrue($"File {dependentTlbPath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportAndOptionSilent_StdOutAndStdErrIsEmpty()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Guid.NewGuid()}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--out", tlbFileName, "--silent");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        result.StdOut.Trim().Should().BeNullOrEmpty();
        result.StdErr.Trim().Should().BeNullOrEmpty();

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportAndOptionSilenceTX801311A6_StdOutAndStdErrIsEmpty()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var tlbFileName = $"{Guid.NewGuid()}.tlb";
        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--out", tlbFileName, "--silence", "TX801311A6");
        result.ExitCode.Should().Be(0);

        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        result.StdOut.Trim().Should().BeNullOrEmpty();
        result.StdErr.Trim().Should().BeNullOrEmpty();

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
    }

    [Fact]
    public void TlbExportAndOptionOverrideTLBId_TLBIdIsCorrect()
    {
        using var context = CompileFixture.GetPreparedTestDirectoryContext();
        var guid = Guid.NewGuid().ToString();

        var tlbFileName = $"{guid}.tlb";
        var yamlFileName = $"{guid}.yaml";

        var tlbFilePath = Path.Combine(Environment.CurrentDirectory, tlbFileName);
        var yamlFilePath = Path.Combine(Environment.CurrentDirectory, yamlFileName);

        var result = Execute(DSComPath, "tlbexport", DemoProjectAssembly1Path, "--out", tlbFilePath, "--overridetlbid", guid);

        result.ExitCode.Should().Be(0);
        File.Exists(tlbFilePath).Should().BeTrue($"File {tlbFilePath} should be available.");

        TlbCommandShouldNotRegisterTypeLib("tlbexport", DemoProjectAssembly1Path, DemoProjectAssembly2Path);

        var dumpResult = Execute(DSComPath, "tlbdump", tlbFilePath, "/out", yamlFilePath);
        dumpResult.ExitCode.Should().Be(0);

        File.Exists(yamlFilePath).Should().BeTrue($"File {yamlFilePath} should be available.");

        var yamlContent = File.ReadAllText(yamlFilePath);
        yamlContent.Should().Contain($"guid: {guid}");

        TlbCommandShouldNotRegisterTypeLib("tlbdump", DemoProjectAssembly1Path, DemoProjectAssembly2Path);
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
}
