::Tips:
::Set the CSIGNTOOL as your SignPackPath
@echo off
path %CSIGNTOOL%;%path%

cd /d "%~dp0"

set OUTDIR=..\Src\slmgr\bin

nuget.exe pack "..\Src\slmgr\slmgr.nuspec" -OutputDirectory "%OUTDIR%" -NoPackageAnalysis
