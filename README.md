# SharpMole: Deep Dive into Windows File Systems and Shares

**SharpMole** is a directory and share permissions auditing tool designed to identify exposed secrets and interesting files. It provides insights into file permissions, helping both defensive and offensive security professionals in their operations. The tool is intended to be run locally on a target file system, from an authenticated perspective.

## Requirements

- **Visual Studio 2022** (Community Edition is sufficient)
- **.NET 6.0 or later** installed
- **EPPlus 7.0.9 or later** (Can be installed via NuGet or the .NET package manager. For more information, visit [EPPlus on NuGet](https://www.nuget.org/packages/EPPlus/))

## Prerequisites

To use **SharpMole**, you need to install **EPPlus** as it is a key dependency for generating XLSX files. Install EPPlus by running the following command in your Visual Studio project's developer console:

```
dotnet add package EPPlus --version 7.0.9
```

## Build Instructions

To build **SharpMole** from source:

1. Open the `.SLN` file within Visual Studio.
2. Select `Release` configuration and change the CPU type to the desired architecture (most likely `x64`). If not already set up, just copy the settings from `AnyCPU`.
3. Open the project's developer console and type the following command to publish the application:
```
dotnet publish -c Release -r win-x64 --self-contained
```
4. You should now have a self-contained and portable `.exe` file and are ready to go!

## Usage and examples
- Basic and general directory scan
```
.\SharpMole.exe --directory="C:\Path\To\Directory"
```
- Directory scan with exclusions
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --exclude="C:\Path\To\Exclude"
```
- Filtering by file type
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --type="txt"
```
- Filtering by file name
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --name="targetString"
```
- Verbosity (display each filename being excluded or enumerated)
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --verbose
```
- Suppress file output
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --suppress
```
- Opsec flag - attempts to avoid generating unauthorised file access events
```
.\SharpMole.exe --directory="C:\Path\To\Directory" --opsec
```

### Operational Security (--opsec Flag)

Use the `--opsec` flag to minimize the risk of generating unauthorized access events in Windows event logs. This mode ensures **SharpMole** checks permissions before accessing files or directories, reducing its footprint and the likelihood of detection. When `--opsec` mode is active, all files identified by **SharpMole** are expected to be accessible, typically resulting in permissions being displayed as `Y Y Y` in the output spreadsheet.

#### Limitations of the --opsec Flag

- **Opsec is not guaranteed**: Opsec mode attempts to avoid generating unauthorized file access events by checking the current users permission before attempting to traverse a directory. However, within hardened environments, SharpMole is likely to generate alerts when more complex ACL rules are in use. TLDR: this flag is not a silver bullet and modification to the codebase may be required in order for SharpMole to work in the target environment you are assessing.
- **Performance Impact**: These additional permission checks add overhead and the time it will take to complete is highly dependent on the size of the target directory/share.
