![NanoXLSX](https://raw.githubusercontent.com/rabanti-github/NanoXLSX/refs/heads/master/Documentation/icons/NanoXLSX.png)

# NanoXLSX

![NuGet Version](https://img.shields.io/nuget/v/NanoXLSX)
![NuGet Downloads](https://img.shields.io/nuget/dt/NanoXLSX)
![GitHub License](https://img.shields.io/github/license/rabanti-github/NanoXLSX)
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX?ref=badge_shield)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

* :white_check_mark: **Minimum of dependencies** (\*
* :x: No need for an installation of Microsoft Office
* :x: No need for Office interop libraries
* :x: No need for proprietary 3rd party libraries
* :x: No need for an installation of the Microsoft Open Office XML SDK (OOXML)

:globe_with_meridians: Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

:page_facing_up: See the **[Change Log](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md)** for recent updates.

## :package: Modules

NanoXLSX v3 is split into modular NuGet packages:

| Module | Status | Description |
|--------|--------|-------------|
| **[NanoXLSX.Core](https://www.nuget.org/packages/NanoXLSX.Core)** | :green_circle: Mandatory, Bundled | Core library with workbooks, worksheets, cells, styles. No external dependencies |
| **[NanoXLSX.Reader](https://www.nuget.org/packages/NanoXLSX.Reader)** | :large_blue_circle: Optional, Bundled | Extension methods to read/load XLSX files. Depends on Core |
| **[NanoXLSX.Writer](https://www.nuget.org/packages/NanoXLSX.Writer)** | :large_blue_circle: Optional, Bundled | Extension methods to write/save XLSX files. Depends on Core |
| **[NanoXLSX.Formatting](https://www.nuget.org/packages/NanoXLSX.Formatting)** | :large_blue_circle: Optional, Bundled | In-line cell formatting (rich text). [External repo](https://github.com/rabanti-github/NanoXLSX.Formatting). Depends on Core |
| **[NanoXLSX](https://www.nuget.org/packages/NanoXLSX)** | :star: Meta-Package | Bundles all of the above. **Recommended for most users** |

> **Note:** All bundled modules are included when you install the `NanoXLSX` meta-package. There are currently no non-bundled (standalone) modules.

For advanced scenarios, you can install only the specific packages you need (e.g. `NanoXLSX.Core` + `NanoXLSX.Writer` for write-only applications).

## :sparkles: What's new in version 3.x

NanoXLSX v3 is a major release with significant architectural changes:

* **Modular architecture** - Split into separate NuGet packages (Core, Reader, Writer, Formatting) with a plugin system
* **New Color system** - Unified `Color` class supporting RGB, ARGB, indexed, theme and system colors
* **Redesigned Font and Fill** - Font properties now use proper enums; Fill supports flexible color definitions with tint
* **PascalCase naming** - All enums and constants follow C# naming conventions
* **Immutable value types** - `Address` and `Range` structs are now immutable
* **In-line formatting** - Rich text cell formatting via the NanoXLSX.Formatting module
* **Utils reorganization** - `Utils` class split into `DataUtils`, `ParserUtils`, `Validators`

:warning: **Breaking changes from v2.x** - There are breaking changes between NanoXLSX v2.6.7 and v3.0.0, mostly related to namespace changes and renamed enum values. See the **[Migration Guide](MigrationGuide.md)** for detailed upgrade instructions.

## :world_map: Roadmap

NanoXLSX v3.x is planned as the **long-term supported version**. Possible future enhancements include:

* :lock: Modern password handling (e.g. SHA-256 for worksheet protection)
* :art: Auto-formatting capabilities
* :1234: Formula assistant for easier formula creation
* :paintbrush: Modern Style builder API
* :speech_balloon: Support for cell comments
* :framed_picture: Embedded images and charts
* :rocket: Performance optimizations

## :book: Reader Support

The reader of NanoXLSX follows the principle of "What You Can Write Is What You Can Read". Therefore, all information about workbooks, worksheets, cells and styles that can be written into an XLSX file by NanoXLSX, can also be read by it.
There are some limitations:

* A workbook or worksheet password cannot be recovered, only its hash
* Information that is not supported by the library will be discarded
* There are some approximations for floating point numbers. These values (e.g. pane split widths) may deviate from the originally written values
* Numeric values are cast to the appropriate .NET types with best effort. There are import options available to enforce specific types
* No support of other objects than spreadsheet data at the moment
* Due to the potential high complexity, custom number format codes are currently not automatically escaped on writing or un-escaped on reading

## :gear: Requirements

NanoXLSX is originally based on PicoXLSX. However, NanoXLSX is now in the development lead, whereas PicoXLSX is a subset of it. The library is currently on compatibility level with .NET version 4.5 and .NET Standard 2.0. Newer versions should of course work as well. Older versions, like .NET 3.5 have only limited support, since newer language features were used.

### .NET 4.5 or newer

\*)The only requirement to compile the library besides .NET (v4.5 or newer) is the assembly **WindowsBase**, as well as **System.IO.Compression**. These assemblies are **standard components in all Microsoft Windows systems** (except Windows RT systems). If your IDE of choice supports referencing assemblies from the Global Assembly Cache (**GAC**) of Windows, select WindowsBase and Compression from there. If you want so select the DLLs manually and Microsoft Visual Studio is installed on your system, the DLL of WindowsBase can be found most likely under "c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll", as well as System.IO.Compression under "c:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.dll". Otherwise you find them in the GAC, under "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\WindowsBase" and "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.IO.Compression"

The NuGet package **does not require dependencies**

### .NET Standard

.NET Standard v2.0 resolves the dependency System.IO.Compression automatically, using NuGet and does not rely anymore on WindowsBase in the development environment. In contrast to the .NET >=4.5 version, **no manually added dependencies necessary** (as assembly references) to compile the library.

### Utility dependencies

The Test project and GitHub Actions may also require dependencies like unit testing frameworks or workflow steps. However, **none of these dependencies are essential to build the library**. They are just utilities. The test dependencies ensure efficient unit testing and code coverage. The GitHub Actions dependencies are used for the automatization of releases and API documentation

## :inbox_tray: Installation

### Using NuGet (recommended)

By package Manager (PM):

```sh
Install-Package NanoXLSX
```

By .NET CLI:

```sh
dotnet add package NanoXLSX
```

:information_source: **Note**: Other methods like adding DLLs or source files directly into your project are technically still possible, but **not recommended** anymore. Use dependency management, whenever possible

## :bulb: Usage

### Quick Start (shortened syntax)

```c#
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.WS.Value("Some Data");                                        // Add cell A1
 workbook.WS.Formula("=A1");                                            // Add formula to cell B1
 workbook.WS.Down();                                                    // Go to row 2
 workbook.WS.Value(DateTime.Now, Style.BasicStyles.Bold);               // Add formatted value to cell A2
 workbook.Save();                                                       // Save the workbook as myWorkbook.xlsx
```

### Quick Start (regular syntax)

```c#
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.CurrentWorksheet.AddNextCell("Some Data");                    // Add cell A1
 workbook.CurrentWorksheet.AddNextCell(42);                             // Add cell B1
 workbook.CurrentWorksheet.GoToNextRow();                               // Go to row 2
 workbook.CurrentWorksheet.AddNextCell(DateTime.Now);                   // Add cell A2
 workbook.Save();                                                       // Save the workbook as myWorkbook.xlsx
```

### Quick Start (read)

```c#
 using NanoXLSX.Extensions;

 Workbook wb = WorkbookReader.Load("basic.xlsx");                       // Read the workbook
 System.Console.WriteLine("contains worksheet name: " + wb.CurrentWorksheet.SheetName);
 foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)
 {
    System.Console.WriteLine("Cell address: " + cell.Key + ": content:'" + cell.Value.Value + "'");
 }
```

## :link: Further References

See the full **API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX/](https://rabanti-github.github.io/NanoXLSX/).

The **[Demo Project](https://github.com/rabanti-github/NanoXLSX.Demo)** contains 27 examples covering various use cases. The demo project is maintained in a separate repository.
See the section **[NanoXLSX](https://github.com/rabanti-github/NanoXLSX.Demo/tree/main/NanoXLSX)** for the specific examples related to NanoXLSX.

See also: [Getting started in the Wiki](https://github.com/rabanti-github/NanoXLSX/wiki/Getting-started)

Hint: You will find most certainly any function, and the way how to use it, in the [Unit Test Project](https://github.com/rabanti-github/NanoXLSX/tree/master/NanoXlsx%20Test)

## :balance_scale: License

NanoXLSX is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

This library claims to be free of any dependencies on proprietary software or libraries.
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX?ref=badge_large)
