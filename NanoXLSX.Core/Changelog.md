# Change Log - NanoXLSX.Core

## v3.1.0

---
Release Date: **(04.05.2026)** <sup>(DMY)</sup>

- **Breaking change:** `Worksheet.Cells` now returns `IReadOnlyDictionary<string, Cell>` instead of `Dictionary<string, Cell>`. All read operations (`["A1"]`, `ContainsKey`, `TryGetValue`, `foreach`, `Count`, `Keys`, `Values`) work unchanged. Mutating properties of an existing cell in-place (e.g. `Cells["A1"].Value = x`) continues to work. Adding or removing cells through `Cells` directly (e.g. `Cells.Add(...)`, `Cells.Remove(...)`, `Cells["A1"] = new Cell(...)`) is no longer possible; use `Worksheet.AddCell(...)` and `Worksheet.RemoveCell(...)` instead.
- Replaced internal cell storage with a `Dictionary<CellKey, Cell>` keyed by `(column, row)` integer pair. This eliminates per-cell address string allocation on every `AddCell` call, reducing memory usage and improving write performance for large workbooks.
- Added internal `CellKey` struct (`NanoXLSX.Internal`) implementing an efficient hash function collision-free across the full Excel address space (16 384 columns × 1 048 576 rows).
- Added `Worksheet.CellValues` property (`IEnumerable<Cell>`): allocation-free enumeration over all cells in a worksheet, preferred for hot iteration paths.

Note: Direct manipulation of the cell dictionary through `Worksheet.Cells` (e.g. `Cells.Add(...)`, `Cells.Remove(...)`, `Cells["A1"] = new Cell(...)`) was never an intended or documented API. Doing so bypassed address validation, style normalisation and internal bookkeeping performed by `AddCell` and `RemoveCell`, and could silently produce invalid workbooks. The change to `IReadOnlyDictionary` makes this boundary explicit.

## v3.0.0

---
Release Date: **(28.02.2026)** <sup>(DMY)</sup>

- Final release of NanoXLSX.Core
- See the [main changelog](https://github.com/rabanti-github/NanoXLSX/blob/master/Changelog.md) for a comprehensive summary of all changes since v2.6.7

## v3.0.0-rc.7

---
Release Date: **15.02.2026** <sup>(DMY)</sup>

- Fixed a bug in the PlugInLoader

## v3.0.0-rc.5

---
Release Date: **22.01.2026** <sup>(DMY)</sup>

- Fixed plug-in handling when loading plug-ins from NuGet packages
- Added Constructor to create a ThemeColor by index
- Added Baseline as Value for vertical alignments
- Moved internal interfaces of the Reader and Writer package to NanoXLSX.Core (namespace `NanoXLSX.Interfaces.Reader` and `NanoXLSX.Interfaces.Writer`)
- Moved and consolidated enums of password types to NanoXLSX.Core (namespace `NanoXLSX.Enums.Password`)

## v3.0.0-rc.4

---
Release Date: **07.01.2026** <sup>(DMY)</sup>

- Added Font properties: `Font.Outline`, `Font.Shadow`, `Font.Condense` and `Font.Extend` (optional font properties)

## v3.0.0-rc.3

---
Release Date: **04.01.2026** <sup>(DMY)</sup>

- Internal change of structured text handling
- Formal change of the `Color` and `ThemeColor` classes
- Removed the property `ColorTheme` from the `Font` class
- Changed the type of the property `ColorVlaue` of the Font class from `string` to `Color` (namespace `NanoXLES.Colors`)

## v3.0.0-rc.2

---
Release Date: **27.11.2025** <sup>(DMY)</sup>

- Refactoring of several enums from lowercase start to uppercase start for better consistency

## v3.0.0-rc.1

---
Release Date: **25.11.2025** <sup>(DMY)</sup>

- Initial release of the core library
