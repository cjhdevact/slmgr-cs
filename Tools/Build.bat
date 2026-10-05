@echo off

if /i "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    set "PATH=C:\Windows\Microsoft.NET\Framework64\v4.0.30319;%PATH%"
) else (
    set "PATH=C:\Windows\Microsoft.NET\Framework\v4.0.30319;%PATH%"
)

cd /d "%~dp0..\Src\slmgr"
msbuild slmgr.csproj /p:Configuration=Release
msbuild slmgr.csproj /p:Configuration=Release-Console
msbuild slmgr.csproj /p:Configuration=Release-Library
cd /d "%~dp0"