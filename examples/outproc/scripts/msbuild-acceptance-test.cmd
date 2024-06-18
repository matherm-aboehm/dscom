@ECHO on

SET root=%~dp0\..\..\..\
SET logs=%~dp0\..\logs
IF NOT EXIST %logs% MKDIR %logs%

PUSHD %~dp0\..

dotnet build-server shutdown

dotnet build -nodeReuse:False -bl:%logs%\build-runtime.binlog
SET ERRUNTIME=%ERRORLEVEL%

dotnet build-server shutdown

POPD

SetLocal EnableDelayedExpansion

SET EXITCODE=0

IF NOT "%ERRUNTIME%" == "0" (
  SET EXITCODE=1
  ECHO "::warning::acceptance test failed."
)

IF "%EXITCODE%" == "0" (
  ECHO "Acceptance test completed successfully."
)

EXIT /B %EXITCODE%
