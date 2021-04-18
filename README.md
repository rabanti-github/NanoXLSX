# NanoXLSX

![NanoXLSX](https://rabanti-github.github.io/NanoXLSX/icons/NanoXLSX.png)

NanoXLSX is a small .NET library written in C#, to create and read Microsoft Excel files in the XLSX format (Microsoft Excel 2007 or newer) in an easy and native way

* **Minimum of dependencies** (\*
* No need for an installation of Microsoft Office
* No need for Office interop libraries
* No need for proprietary 3rd party libraries
* No need for an installation of the Microsoft Open Office XML SDK (OOXML)

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md)** for recent updates.

## Reader Support

Currently, only basic reader functionality is available:

* Reading and casting of cell values into the appropriate data types
* Reading of several worksheets in on workbook with names
* Limited processing of styles (when reading) at the moment
* No support of other objects than spreadsheet data at the moment

**Note: Styles in loaded files are only considering number formats (to determine date and time values), as well as custom formats. The scope of reader functionality may change with future versions.**

## Requirements

NanoXLSX is based on PicoXLSX and was created with .NET version 4.5. Newer versions like 4.6 are working and tested. Furthermore, .NET Standard 2.0 is supported since v1.6. Older versions of.NET like 3.5 and 4.0 may also work with minor changes. Some functions introduced in .NET 4.5 were used and must be adapted in this case. 

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

### Using Nuget

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

The [demo project](https://github.com/rabanti-github/NanoXLSX/tree/master/Demo) contains 16 simple use cases. You can find also the full documentation in the [Documentation-Folder](https://github.com/rabanti-github/NanoXLSX/tree/master/docs) (html files or single chm file) or as C# documentation in the particular .CS files.

Note: The demo project of the .NET Standard version is identical and only links to the .NET >=4.5 version files.

See also: [Getting started in the Wiki](https://github.com/rabanti-github/NanoXLSX/wiki/Getting-started)
