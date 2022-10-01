# NanoXLSX ![NanoXLSX](https://github.com/rabanti-github/NanoXLSX/blob/master/Documentation/icons/NanoXLSX.png)

![nuget](https://img.shields.io/nuget/v/NanoXLSX.svg?maxAge=86400)![license](https://img.shields.io/github/license/rabanti-github/NanoXlsx.svg)[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX?ref=badge_shield)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

* **Minimum of dependencies** (\*
* No need for an installation of Microsoft Office
* No need for Office interop libraries
* No need for proprietary 3rd party libraries
* No need for an installation of the Microsoft Open Office XML SDK (OOXML)

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md)** for recent updates.

## What's new in version 2.x

There are some additional functions for workbooks and worksheets, as well as support of further data types.
The biggest change is the full capable reader support for workbook, worksheet and style information. Also, all features are now fully unit tested. This means, that NanoXLSX is no longer in Beta status, but production ready. Some key features are:

* Full reader support for styles, workbooks, worksheets and workbook metadata
* Copy functions for worksheets
* Advance import options for the reader
* Several additional checks, exception handling and updated documentation

## Road map

Version 2.x of NanoXLSX was completely overhauled and a high number of (partially parametrized) unit tests with a code coverage of >99% were written to improve the quality of the library.
However, it is not planned as a LTS version. The upcoming v3.x is supposed to introduce some important functions, like in-line cell formatting, better formula handling and additional worksheet features.
Furthermore, it is planned to introduce more modern OOXML features like the SHA256 implementation of worksheet passwords.
One of the main aspects of this upcoming version is a better modularization, as well as the consolidation with PicoXLS to one single code base.

## Reader Support

The reader of NanoXLS follows the principle of "What You Can Write Is What You Can Read". Therefore, all information about workbooks, worksheets, cells and styles that can be written into an XLSX file by NanoXLSX, can also be read by it.
There are some limitations:

* A workbook or worksheet password cannot be recovered, only its hash
* Information that is not supported by the library will be discarded
* There are some approximations for floating point numbers. These values (e.g. pane split widths) may deviate from the originally written values
* Numeric values are cast to the appropriate .NET types with best effort. There are import options available to enforce specific types
* No support of other objects than spreadsheet data at the moment

## Requirements

NanoXLSX is originally based on PicoXLSX. However, NanoXLSX is now in the development lead, whereas PicoXLSX is a subset of it. The library is currently on compatibility level with .NET version 4.5 and .NET Standard 2.0. Newer versions should of course work as well. Older versions, like .NET 3.5 have only limited support, since newer language features were used.

### .NET 4.5 or newer

*)The only requirement to compile the library besides .NET (v4.5 or newer) is the assembly **WindowsBase**, as well as **System.IO.Compression**. These assemblies are **standard components in all Microsoft Windows systems** (except Windows RT systems). If your IDE of choice supports referencing assemblies from the Global Assembly Cache (**GAC**) of Windows, select WindowsBase and Compression from there. If you want so select the DLLs manually and Microsoft Visual Studio is installed on your system, the DLL of WindowsBase can be found most likely under "c:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll", as well as System.IO.Compression under "c:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.5\System.IO.Compression.dll". Otherwise you find them in the GAC, under "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\WindowsBase" and "c:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.IO.Compression"

The NuGet package **does not require dependencies**

### .NET Standard

.NET Standard v2.0 resolves the dependency System.IO.Compression automatically, using NuGet and does not rely anymore on WindowsBase in the development environment. In contrast to the .NET >=4.5 version, **no manually added dependencies necessary** (as assembly references) to compile the library.

Please note that the demo project of the .NET Standard version will not work in Visual Studio 2017. To get the build working, unload the demo project of the .NET Standard version.

### Documentation project

If you want to compile the documentation project (folder: Documentation; project file: shfbproj), you need also the **[Sandcastle Help File Builder (SHFB)](https://github.com/EWSoftware/SHFB)**. It is also freely available. But you don't need the documentation project to build the NanoXLSX library.

The .NET version of the documentation may vary, based on the installation. If v4.5 is not available, upgrade to target to a newer version, like v4.6

## Installation

### Using NuGet

By package Manager (PM):

```sh
Install-Package NanoXLSX
```

By .NET CLI:

```sh
dotnet add package NanoXLSX
```

### As DLL

Simply place the NanoXLSX DLL into your .NET project and add a reference to it. Please keep in mind that the .NET version of your solution must match with the runtime version of the NanoXLSX DLL (currently compiled with 4.5 and .NET Standard 2.0).

### As source files

Place all .CS files from the NanoXLSX source folder and its sub-folders into your project. In case of the .NET >=4.5 version, the necessary dependencies have to be referenced as well.

## Usage

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
 Workbook wb = Workbook.Load("basic.xlsx");                             // Read the workbook
 System.Console.WriteLine("contains worksheet name: " + wb.CurrentWorksheet.SheetName);
 foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)
 {
    System.Console.WriteLine("Cell address: " + cell.Key + ": content:'" + cell.Value.Value + "'");
 }
```

## Further References

See the full **API-Documentation** at: [https://rabanti-github.github.io/NanoXLSX/](https://rabanti-github.github.io/NanoXLSX/).

The [demo project](https://github.com/rabanti-github/NanoXLSX/tree/master/Demo) contains 18 simple use cases. You can find also the full documentation in the [Documentation-Folder](https://github.com/rabanti-github/NanoXLSX/tree/master/docs) (html files or single chm file) or as C# documentation in the particular .CS files.

Note: The demo project of the .NET Standard version is identical and only links to the .NET >=4.5 version files.

See also: [Getting started in the Wiki](https://github.com/rabanti-github/NanoXLSX/wiki/Getting-started)

Hint: You will find most certainly any function, and the way how to use it, in the [Unit Test Project](https://github.com/rabanti-github/NanoXLSX/tree/master/NanoXlsx%20Test)



## License
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Frabanti-github%2FNanoXLSX?ref=badge_large)