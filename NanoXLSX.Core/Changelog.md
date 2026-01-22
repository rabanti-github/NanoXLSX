# Change Log - NanoXLSX.Core

## v3.0.0-rc.5

---
Release Date: **22.01.2026**

- Fixed plug-in handling when loading plug-ins from NuGet packages
- Added Constructor to create a ThemeColor by index
- Added Baseline as Value for vertical alignments
- Moved internal interfaces of the Reader and Writer package to NanoXLSX.Core (namespace `NanoXLSX.Interfaces.Reader` and `NanoXLSX.Interfaces.Writer`)
- Moved and consolidated enums of password types to NanoXLSX.Core (namespace `NanoXLSX.Enums.Password`)


## v3.0.0-rc.4

---
Release Date: **07.01.2026**

- Added Font properties: `Font.Outline`, `Font.Shadow`, `Font.Condense` and `Font.Extend` (optional font properties)

## v3.0.0-rc.3

---
Release Date: **04.01.2026**

- Internal change of structured text handling
- Formal change of the `Color` and `ThemeColor` classes
- Removed the property `ColorTheme` from the `Font` class
- Changed the type of the property `ColorVlaue` of the Font class from `string` to `Color` (namespace `NanoXLES.Colors`)


## v3.0.0-rc.2

---
Release Date: **27.11.2025**

- Refactoring of several enums from lowercase start to uppercase start for better consistency

## v3.0.0-rc.1

---
Release Date: **25.11.2025**

- Initial release of the core library