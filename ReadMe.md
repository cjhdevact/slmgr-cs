# slmgr-cs

A C# reimplementation of the Windows Software Licensing Management Tool (`slmgr.vbs`), built for environments where VBScript is no longer available. The project ships the same command set and behaviour as the original script, and additionally provides a library build so external programs can invoke it programmatically.

This project is a rewrite based on the Windows 11 26H2 (10.0.26300.9457) version of `slmgr.vbs`.

## Background

Microsoft has deprecated the VBScript engine on Windows. `slmgr.vbs` — the standard tool for installing product keys, activating Windows, configuring KMS, and inspecting license state — will no longer run once the engine is removed.

This project is a drop-in replacement written in C# targeting .NET Framework 4.5.2. It mirrors the option names, output strings and side effects of the original `slmgr.vbs`, and can be compiled into three different outputs from a single source tree.

## Features

- **Full command coverage** of `slmgr.vbs`:
  - Product key and activation: `/ipk`, `/upk`, `/ato`, `/atp`, `/dti`
  - License information: `/dli`, `/dlv`, `/xpr`
  - Advanced: `/cpky`, `/ilc`, `/rilc`, `/rearm`, `/rearm-app`, `/rearm-sku`
  - KMS client: `/skms`, `/ckms`, `/skms-domain`, `/ckms-domain`, `/skhc`, `/ckhc`
  - KMS host: `/sprt`, `/sai`, `/sri`, `/sdns`, `/cdns`, `/spri`, `/cpri`, `/act-type`
  - Token-based activation: `/lil`, `/ril`, `/ltc`, `/fta`
  - Active Directory activation: `/ad-activation-online`, `/ad-activation-get-iid`, `/ad-activation-apply-cid`, `/ao-list`, `/del-ao`
- **Three build outputs from one source tree**:
  - `slmgr.exe` — Windows GUI host (output via MessageBox)
  - `slmgrc.exe` — console host (output via Console)
  - `slmgrlib.dll` — class library for external programs (output returned as a string)

## Build Requirements

- Visual Studio 2015 or later (MSBuild 14.0+)
- .NET Framework 4.5.2
- Administrator privileges at runtime for most commands

## Project Layout

```text
slmgr/
├── slmgr.csproj               Single project, multiple build configurations
├── App.config
├── Program.cs                 Entry point (compiled only in exe builds)
├── Commands.cs                Command dispatch and implementation
├── WmiHelper.cs               WMI wrapper for SoftwareLicensingService / Product
├── AdHelper.cs                Active Directory activation helpers
├── TkaHelper.cs               Token-based activation helpers
├── OutputMode.cs              Output routing (Console / MessageBox / redirect)
├── Resources.cs               ResourceManager wrapper
├── SlmgrApi.cs                Public entry point: slmgr.Slmgr.Run(...)
├── SlmgrResult.cs             Return type for the API
├── ExitSignalException.cs     Internal control-flow exception
├── WindowsAppId.cs            Shared constants
├── Resources/                 Language resources
└── Properties/
    ├── AssemblyInfo.cs
    ├── Resources.resx
    └── Settings.settings
```

### In Visual Studio IDE

1. Open the solution.
2. **Build → Configuration Manager**.
3. Under **Active solution configuration**, create `Debug-Console`, `Release-Console`, `Debug-Library`, `Release-Library` by copying from `Debug` / `Release`.
4. Ensure **Build** is checked for the `slmgr` project in every configuration.
5. Switch configurations in the toolbar to build each artifact.

### From the command line

```bat
msbuild slmgr.csproj /p:Configuration=Release
msbuild slmgr.csproj /p:Configuration=Release-Console
msbuild slmgr.csproj /p:Configuration=Release-Library
```

After all three commands, `bin\Release\` contains:

```text
bin\Release\
├── slmgr.exe
├── slmgrc.exe
├── slmgrlib.dll
├── en-US\
...
```

Each artifact carries its own satellite assemblies (the .NET resource system keys satellites by assembly name). The satellite contents are identical; the duplication is a consequence of the single-project multi-configuration approach. To share a single satellite across all three artifacts, extract the resources into a separate class library project and reference it from all three.

## Usage

### `slmgr.exe` — GUI host

```console
slmgr.exe /dlv
slmgr.exe /ipk XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
slmgr.exe /skms kms.example.com:1688
slmgr.exe /ato
```

Output is shown in a MessageBox. Intended for interactive use.

### `slmgrc.exe` — console host

```console
slmgrc.exe /dlv
slmgrc.exe /ipk XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
slmgrc.exe /ato
```

Output goes to the console. Intended for scripting and automation.

### Remote machine syntax

Both executables accept up to three leading arguments for remote operation:

```console
slmgrc.exe \\remote-pc admin password /dlv
```

`/ao-list` and `/del-ao` do not support remote execution, matching the original script.

## Library Usage

Reference `slmgrlib.dll` from your project and call `slmgr.Slmgr.Run(...)`:

```csharp
using slmgr;

SlmgrResult r = Slmgr.Run("/dlv");
Console.WriteLine("ExitCode: " + r.ExitCode);
Console.WriteLine(r.Output);

SlmgrResult r2 = Slmgr.Run("/ipk W269N-WFGWX-YVC9B-4J6C9-T83GX");
if (!r2.Success)
{
    Console.WriteLine("Failed: " + r2.Output);
}
```

### API reference

```csharp
namespace slmgr
{
    public static class Slmgr
    {
        public static SlmgrResult Run(string commandLine);
        public static SlmgrResult Run(string[] args);
    }

    public class SlmgrResult
    {
        public string Output   { get; set; }   // captured output text
        public int    ExitCode { get; set; }   // exit code
        public bool   Success  { get; }        // true when ExitCode == 0
    }
}
```

- `Run(string)` splits the command line on whitespace, honouring double quotes.
- `Run(string[])` takes a pre-parsed argument array.
- Both overloads are thread-safe; concurrent calls are serialized.
- The API does not call `Environment.Exit`; the exit code is returned through `SlmgrResult.ExitCode`.

### PowerShell

```powershell
Add-Type -Path 'slmgrlib.dll'
$r = [slmgr.Slmgr]::Run('/dlv')
Write-Host "ExitCode: $($r.ExitCode)"
Write-Host $r.Output
```

### VB.NET

```vbnet
Imports slmgr

Dim r As SlmgrResult = Slmgr.Run("/dlv")
Console.WriteLine("ExitCode: " & r.ExitCode)
Console.WriteLine(r.Output)
```

## Localization

The UI is localized through standard .NET resources.

`ResourceManager` selects the appropriate resource based on `Thread.CurrentThread.CurrentUICulture`. To force a specific language:

```csharp
System.Threading.Thread.CurrentThread.CurrentUICulture =
    new System.Globalization.CultureInfo("zh-CN");
```

### Adding a new language

1. Copy `Resources\Strings.resx` to `Resources\Strings.<culture>.resx` (for example `Strings.zh-TW.resx`).
2. Translate the values.
3. Add an `<EmbeddedResource Include="Resources\Strings.zh-TW.resx" />` entry to `slmgr.csproj`.
4. Rebuild. A satellite assembly `zh-TW\slmgr.resources.dll` will be produced.

## Behavioural Notes

The following behaviours intentionally match `slmgr.vbs`:

- Option names are identical, including the historically unusual `/cpri` for clearing KMS priority to Low.
- Output text uses the same message templates, including placeholder tokens such as `%PKEY%`, `%ACTID%`, `%ERRCODE%`.
- Hexadecimal error codes are printed without leading zeros, mirroring VBScript's `Hex()` function.
- `/ao-list` formats the binary `msSPP-CSVLKSkuId` attribute with the same byte ordering as the script's `GuidToString` helper.
- Remote connections use `Impersonate` + `PacketPrivacy`, matching the script's WMI security settings.

Two small differences from the script:

- The subscription-status block in `/dli` and `/dlv` is suppressed when `SubscriptionType` is `Unknown`, instead of printing a placeholder. This avoids noisy output on Windows builds that predate the subscription-related WMI properties.
- `/ao-list` and `/del-ao` fail fast on remote connections with a clear message, rather than a generic WMI error.

## Limitations

- **Remote WMI**: supported for most commands. `/ao-list` and `/del-ao` require local execution.
- **Remote registry**: the KMS version registry write-back is local only. The original script used `StdRegProv`; this implementation uses `Microsoft.Win32.RegistryKey` against the local machine.
- **Token-based activation** (`/fta`, `/ltc`, `/lil`, `/ril`) depends on the `SPPWMI.SppWmiTokenActivationSigner` COM object being registered. This is present on Windows installations that support Token-based Activation; on other systems the commands will fail with a clear error.
- **Administrator privileges**: nearly all commands require elevation. Running without it returns `0xC004F025` (access denied), matching the script's behaviour.

## License

This project is licensed under the MIT License.

## Acknowledgements

The command set, output strings, and internal behaviour are based on the original `slmgr.vbs` shipped with Windows. This project is an independent reimplementation and is not affiliated with or endorsed by Microsoft.