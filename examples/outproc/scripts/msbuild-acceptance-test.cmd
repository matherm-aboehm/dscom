@ECHO on

SET root=%~dp0\..\..\..\
SET logs=%~dp0\..\logs
IF NOT EXIST %logs% MKDIR %logs%

PUSHD %root%

dotnet build-server shutdown

dotnet pack -p:Configuration=Release -bl:%logs%\pack.binlog
SET ERPACK=%ERRORLEVEL%

IF NOT EXIST %root%\_packages MKDIR _packages

XCOPY /Y /I /C /F .\src\dscom\bin\Release\*.nupkg _packages\
XCOPY /Y /I /C /F .\src\dscom.build\bin\Release\*.nupkg _packages\

dotnet build-server shutdown

dotnet nuget locals global-packages --clear

POPD

PUSHD %~dp0\..

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Clean -p:Configuration=Release -p:PerformAcceptanceTest=Runtime -p:SkipResolvePackageAssets=true

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Restore -p:Configuration=Release -p:PerformAcceptanceTest=Runtime

dotnet build-server shutdown

dotnet add server\common\contract.csproj package --prerelease -s %root%\_packages dSPACE.Runtime.InteropServices.BuildTasks

dotnet msbuild -nodeReuse:False -t:Restore -p:Configuration=Release -p:PerformAcceptanceTest=Runtime

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Build -p:Configuration=Release -p:PerformAcceptanceTest=Runtime -bl:%logs%\build-runtime.binlog
SET ERRUNTIME=%ERRORLEVEL%

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Clean -p:Configuration=Release -p:PerformAcceptanceTest=NetStandard  -p:SkipResolvePackageAssets=true

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Restore -p:Configuration=Release -p:PerformAcceptanceTest=NetStandard

dotnet build-server shutdown

dotnet msbuild -nodeReuse:False -t:Build -p:Configuration=Release -p:PerformAcceptanceTest=NetStandard -bl:%logs%\build-standard.binlog

SET ERSTANDARD=%ERRORLEVEL%

dotnet build-server shutdown

dotnet remove server\common\contract.csproj package dSPACE.Runtime.InteropServices.BuildTasks

dotnet build-server shutdown

POPD

SetLocal EnableDelayedExpansion

SET EXITCODE=0

IF NOT "%ERPACK%" == "0" (
  SET EXITCODE=1
  ECHO "::warning:: dotnet pack failed."
)

IF NOT "%ERRUNTIME%" == "0" (
  SET EXITCODE=1
  ECHO "::warning::Runtime specific acceptance test failed."
)

IF NOT "%ERSTANDARD%" == "0" (
  SET EXITCODE=1
  ECHO "::warning::.NET Standard 2.0 specific acceptance test failed."
)

IF "%EXITCODE%" == "0" (
  ECHO "Acceptance test completed successfully."
)

EXIT /B %EXITCODE%
