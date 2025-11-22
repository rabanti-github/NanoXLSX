# Migration Guide v2.x to v3.0.0

## Core classes

### Workbook

  - The method `Workbook.Load(...)` and `Workbook.LoadAsync(...)` were removed due to the modularization of NanoXLSX. Please see the  **Reader section** for more details. 
  - The method `Workbook.AddStyle(Style)` was completely removed, after marked as obsolete in version 2.x. Styles should be added directly to cells or ranges. 
  - The method `Workbook.AddStyleComponent(Style, AbstractStyle)` was completely removed, after marked as obsolete in version 2.x. Styles should be modified directly on cells, e.g. `workbook.CurrentWorksheet.Cells["A1"].CellStyle.CurrentFont.Bold = true;` or `workbook.CurrentWorksheet.Cells["A1"].CellStyle.Append(fontStyle)`. 
  - The methods `Workbook.RemoveStyle(Style)`, `Workbook.RemoveStyle(Style, bool)`, `Workbook.RemoveStyle(string)` and `Workbook.RemoveStyle(string, bool)` were completely removed, after marked as obsolete in version 2.x. Styles should be removed directly from cells (e.g. `workbook.CurrentWorksheet.Cells["A1"].RemoveStyle()`). 

---

### Worksheet

  - The method `Worksheet.SetSelectedCells(Range)`(and overloads) was replaced by three methods: `Worksheet.AddSelectedCell(Range)`, `Worksheet.AddSelectedCells(String)` (address or range), `Worksheet.AddSelectedCells(Address)` and `Worksheet.AddSelectedCells(Address, Address)`.
  - The method `Worksheet.RemoveSelectedCells()` was replaced by `Worksheet.ClearSelectedCells()`. 
  - The methods `Worksheet.RemoveSelectedCells(Range)`, `Worksheet.RemoveSelectedCells(String)` (address or range), `Worksheet.RemoveSelectedCells(Address)` and `Worksheet.RemoveSelectedCells(Address, Address)` were introduced to remove specific selected cells or ranges from the selection.
  - The property `Worksheet.SelectedCells` was changed from type `Range` to `List<Range>`, to allow multiple selected ranges in a worksheet.
  - The property `Workshet.SelectedCellRanges` was replaced by `Worksheet.SelectedCells`.
  - The property `Workshet.SelectedCellRange` was replaced by `Worksheet.SelectedCells`, where the value would be the only entry in the list.
  - The property `Worksheet.SheetProtectionPassword` was changed from type `string` to the interface `IPassword` (namespace `NanoXLSX.Interfaces`), by default implemented by `LegacyPassword`. `IPassword` contains several methods to get or set a password, and to  get or set its hash value. The property is instantiated by default on every worksheet. The property has to be adapted in the code, e.g. `string plainTextPassword = worksheet.SheetProtectionPassword.GetPassword()`. 
  - The property `Worksheet.SheetProtectionPasswordHash` was removed, as the password handling was changed. The property can now be found in the property: `Worksheet.SheetProtectionPassword.PasswordHash`). 
  
  - The following property and method behaviors have changed:

| Property/Method               | Old Behavior                | New Behavior             |
|-------------------------------|-----------------------------|--------------------------|
| `Worksheet.SelectedCells`     | List was null by default    | List is empty by default |
| `Worksheet.SelectedCells`     | Every added range was present | If ranges are overlapping, or even enclosed in other ranges, the ranges are automatically recalculated, so that every address only occurs in one particular range |
| `Worksheet.SetSelectedCells(string)` | Method accepted ranges or addresses with fixed cell ranges (e.g. "$A$1:$R$1") | Method transforms addresses or ranges automatically in neutral addresses or ranges (e.g. "A1:R1") |
| `Worksheet.SheetProtectionPassword`  | Was null (string) by default         | Is instantiated by default with a structured object of the type `LegacyPassword`  |

- The public constant values of the `Worksheet` class were renamed, according to the following overview:

| Old Constant                | New Constant             |
|-----------------------------|---------------------------
| `MAX_WORKSHEET_NAME_LENGTH` | `MaxWorksheetNameLength` |
| `DEFAULT_COLUMN_WIDTH`      | `DefaultWorksheetColumnWidth`     |
| `DEFAULT_ROW_HEIGHT`        | `DefaultWorksheetRowHeight`       |
| `MAX_COLUMN_NUMBER`         | `MaxColumnNumber`        |
| `MIN_COLUMN_NUMBER`         | `MinColumnNumber`        |
| `MIN_COLUMN_WIDTH`          | `MinColumnWidth`         |
| `MIN_ROW_HEIGHT`            | `MinRowHeight`           |
| `MAX_COLUMN_WIDTH`          | `MaxColumnWidth`         |
| `MAX_ROW_NUMBER`            | `MaxRowNumber`           |
| `MIN_ROW_NUMBER`            | `MinRowNumber`           |
| `MAX_ROW_HEIGHT`            | `MaxRowHeight`           |
| `AUTO_ZOOM_FACTOR`          | `AutoZoomFactor`         |
| `MIN_ZOOM_FACTOR`           | `MinZoomFactor`          |
| `MAX_ZOOM_FACTOR`           | `MaxZoomFactor`          |
	
---

### Styles 

Styles were undergoing several changes in version 3.0.0, to improve usability and consistency.
Especially the `Font` class was completely redesigned, according to the Excel specifications.
Furthermore, a lot of constants were renamed to follow the C# naming conventions.

---

#### Font

- The public constant values of the `Font` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|-----------------------------------|
| `DEFAULT_MAJOR_FONT`     | `DefaultMajorFont`     |          |
| `DEFAULT_MINOR_FONT`     | `DefaultMinorFont`     |          |
| `DEFAULT_FONT_NAME`      | `DefaultFontName`      |          |
| `DEFAULT_FONT_SCHEME`    | `DefaultFontScheme`    |          |
| `MIN_FONT_SIZE`          | `MinFontSize`          |          |
| `MAX_FONT_SIZE`          | `MaxFontSize`          |          |
| `DEFAULT_FONT_SIZE`      | `DefaultFontSize`      |          |
| `DEFAULT_FONT_FAMILY`    | `DefaultFontSize`      | The type was changed from `int` to the enum `Font.FontFamilyValue` |
| `DEFAULT_VERTICAL_ALIGN` | `DefaultVerticalAlign` | The type was changed from enum `Font.VerticalAlignValue` to `Font.VerticalTextAlignValue` |

- The property `Font.Family` was changed from type `string` to the enum `Font.FontFamilyValue`.The value has to be replaced by one of the following available values:
```cs
NotApplicable, Roman, Swiss, Modern, Script, Decorative, Reserved1, Reserved2, Reserved3, Reserved4, Reserved5, Reserved6, Reserved7, Reserved8, Reserved9
// Mostly used: Roman, Swiss, Modern, Script, Decorative
```

- The property `Font.Charset` was changed from type `string` to the enum `Font.CharsetValue`. The initialization default value is `CharsetValue.Default` The value has to be replaced by one of the following available values:
```cs
ApplicationDefined, ANSI, Default, Symbols, Mac, ShiftJIS, Hangul, Johab, GBK, Big5, Greek, Turkish, Vietnamese, Hebrew, Arabic, Baltic, Cyrillic, Thai, EasternEuropean, OEM
// ApplicableDefined is usually ignored, and Default may be used instead
```

- The property `Font.ColorScheme` was changed from type `int` to the enum `Theme.ColorSchemeElement`. The value has to be replaced by one of the available values (See **Theme section** ). The initialization default value is `Theme.ColorSchemeElement.light1`.
- The property `Font.VerticalAlign` was changed from type `Font.VerticalAlignValue` to the enum `Font.VerticalTextAlignValue`. Only the enum name has to be changed (see below):
- The enum `Font.VerticalAlignValue` was renamed to `Font.VerticalTextAlignValue`. The available values remain unchanged


#### Border

- The public constant values of the `Border` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|-----------------------------------|
| `DEFAULT_BORDER_STYLE`   | `DefaultBorderStyle`   |          |
| `DEFAULT_COLOR`          | `DefaultBorderColor`   |          |

#### Fill

- The public constant values of the `Fill` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|-----------------------------------|
| `DEFAULT_COLOR`          | `DefaultColor`         |          |
| `DEFAULT_INDEXED_COLOR`  | `DefaultIndexedColor`  |          |
| `DEFAULT_PATTERN_FILL`   | `DefaultPatternFill`   |          |

- The static method `Fill.ValidateColr(string,bool, bool)` was moved to the utils class `Validators.ValidateColr(string,bool, bool)` in namespace `NanoXLSX.Utils`. The class has to be changed in the code, but the method signature remains unchanged.

#### CellXf

- The public constant values of the `CellXf` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|-----------------------------------|
| `DEFAULT_HORIZONTAL_ALIGNMENT` | `DefaultHorizontalAlignment`|          |
| `DEFAULT_ALIGNMENT`      | `DefaultAlignment`     |          |
| `DEFAULT_TEXT_DIRECTION` | `DefaultTextDirection` |          |
| `DEFAULT_VERTICAL_ALIGNMENT`   | `DefaultVerticalAlignment`  |          |

#### NumberFormat

- The public constant values of the `NumberFormat` class were renamed, according to the following overview:

| Old Constant             | New Constant           | Remarks  |
|--------------------------|-----------------------------------|
| `CUSTOMFORMAT_START_NUMBER` | `CustomFormatStartNumber`|          |
| `DEFAULT_NUMBER`         | `DefaultNumber`        |          |

- The enum values of `NumberFormat.FormatRange` were renamed, according to the following overview:

| Old Enum Value           | New Enum Value         | Remarks  |
|--------------------------|-----------------------------------|
| `FormatRange.defined_format` | `FormatRange.DefinedFormat`   |          |
| `FormatRange.custom_format_` | `FormatRange.CustomFormat`    |          |
| `FormatRange.invalied`   | `FormatRange.Inavlid`  |          |
| `FormatRange.undefined`  | `FormatRange.Undefined`|          |


---

### Theme
The `Theme` class was introduced with NanoXLSX v3. It represents the theme of a workbook, which contains several color schemes and font schemes.
The class can mostly be ignored unless specific stylings are required.
Theme may be references ind Styles, especially in Fonts.
- The enum `Theme.ColorSchemeElement` was introduced to represent the color scheme elements of a theme. The available values are:
```cs
 dark1, light1, dark2, light2, accent1, accent2, accent3, accent4, accent5, accent6, hyperlink, followedHyperlink
```

---

## Reader

When it comes to reading workbooks, the reader was completely separated form NanoXLSX, as an own package: `NanoXLSX.Reader`. This package can be added to `NanoXLSX.Core` and provides several extension methods to load workbooks, worksheets, styles, etc. from files or streams.

### Workbook

  - The methods `Workbook.Load()` and `Workbook.LoadAsync()` were removed due to the modularization of NanoXLSX. To load workbooks, the new class `WorkbookReader` (in namespace `NanoXLSX.Extensions`) was introduced. Sample usage:

 ```csharp
  using NanoXLSX.Extensions;

  Workbook workbook = WorkbookReader.Load("path_to_file.xlsx");
  Workbook workbook = WorkbookReader.Load("path_to_file.xlsx", new ReaderOptions(){ DateTimeFormat = "yyyy.MM.dd hh:mm:ss" }); // Using options
  Workbook workbook = WorkbookReader.Load(stream); // Using a stream
  Workbook workbook = await WorkbookReader.LoadAsync("path_to_file.xlsx"); // Using async method
  ```

  - If a workbook does not contain Metadata information, the property `Workbook.WorkbookMetadata` returned null. Now, an empty Metadata object is created by default in this case. The default object may also contain default properties, like a defined Application


  ---

### ImportOptions

- The class `ImportOptions` was renamed to `ReaderOptions`, to better reflect the purpose of the class. The class name has to be changed in the code.
- The public constants of the former `ImportOptions` class were moved to the new `ReaderOptions` class, according to the following overview:

| Old ImportOptions Constant  | New ReaderOptions Constant      | Remarks        |
|-------------------------|----------------------|------------------------|
| `DEFAULT_TIMESPAN_FORMAT`   | `DefaultTimeSpanFormat`         |         |
| `DEFAULT_DATETIME_FORMAT`   | `DefaultDateTimeFormat`         |         |

- The following method of the class `ReaderOptions` were renamed:

| Old method name | New method name | Remarks |
|-----------------|-----------------|---------|
| `EnforceValidColumnDimensions` | `EnforceStrictValidation` | Variants were combined in one flag |
| `EnforceValidRowDimensions` | `EnforceStrictValidation` | Variants were combined in one flag |

- The Enum value `ImportOptions.GlobalType.AllSingleToDecimal` (new in class `ReaderOptions`) was completely removed. If the behavior is required, the value `ReaderOptions.GlobalType.AllNumbersToDecimal` can be used instead.

---

### Cell

  - When reading Worksheets, All sorts of new line combinations could be read in for string cell values, like `\n\r` that was transformed to `\r\n\r\n`. All new lines are transformed now to `\n`, and `\r` is always stripped in combination with `\n`



---

### Internal reader classes

- All internally used reader classes were moved from the namespace `NanoXLSX.LowLevel` to the new namespace `NanoXLSX.Internal.Readers`. If these classes were used directly in the code, the namespace has to be adapted.
- The architecture of all internally used reader classes was changed redesigned from scratch. If you have modified such classes, these modifications probably have to be redone.
- **Please note**: The reader classes are not intended to be directly modified. However, you can implement custom readers that either replaces existing readers, or can be appended at several positions during the read process. This is part of the introduced plugin architecture.



## Common

All changes related to common functions

### Utils

- The general `Utils` class was removed and replaced by several specific utils classes in the namespace `NanoXLSX.Utils`. The class name has to be adapted, according to the following method overview:

| Old Utils Method        | New Utils Class      | Remarks        |
|-------------------------|----------------------|------------------------|
| `GetOADateTimeString`   | `DataUtils`          | No changes of the signature |
| `GetOADateTime`         | `DataUtils`          | No changes of the signature |
| `GetOATimeString`       | `DataUtils`          | No changes of the signature |
| `GetOATime`             | `DataUtils`          | No changes of the signature |
| `GetDateFromOA`   	  | `DataUtils`          | No changes of the signature |
| `GetInternalColumnWidth`| `DataUtils`          | No changes of the signature |
| `GetInternalRowHeight`  | `DataUtils`          | No changes of the signature |
| `GetInternalPaneSplitWidth` | `DataUtils`      | No changes of the signature |
| `GetInternalPaneSplitHeight`| `DataUtils`      | No changes of the signature |
| `GetPaneSplitHeight`    | `DataUtils`          | No changes of the signature |
| `GetPaneSplitWidth`     | `DataUtils`          | No changes of the signature |
| `ToUpper`               | `ParserUtils`        | No changes of the signature |
| `ToString`              | `ParserUtils`        | No changes of the signature |
| `GeneratePasswordHash`  | `NanoXLSX.LegacyPassword` - new method name: `GenerateLegacyPasswordHash(string)` | No longer an utils method |

- The public constant values of the former `Utils` class were moved to specific utils classes in the name space `NanoXLSX.Utils.Constants`. The class names have to be adapted, according to the following overview:

| Old Constant                | New Class and Constant       | Remarks        |
|-----------------------------|------------------------------|------------------------|
| `Utils.MIN_OA_DATE_VALUE`   | `DataUtils.MinOADateValue`   |
| `Utils.MAX_OA_DATE_VALUE`   | `DataUtils.MaxOADateValue`   |
| `Utils.FIRST_ALLOWED_EXCEL_DATE` | `DataUtils.FirstAllowedExcelDate`|
| `Utils.LAST_ALLOWED_EXCEL_DATE`  | `DataUtils.LastAllowedExcelDate` |
| `Utils.INVARIANT_CULTURE`   | `DataUtils.InvariantCulture` |





- Workbook.Load() vs WorkbokReader.Load()
- ImportOptions vs ReaderOptions
  + EnforceValidRowDimensions and EnforceValidColumnDimensions are replaced by EnforceStrictValidation (default = false)

## Workbook
- Deprecated methods removed: AddStyle, AddStyleComponent, RemoveStyle (several overloads)

## Worksheet

- Password handling
  - LegacyPassword is always defined (should never be null) 

### Methods
- Method SetSelectedCells was replaced by AddSelectedCell, RemoveSelectedCells and ClearSelectedCells. Further overload methods were added to the first two methods

## Address
- Address fields `Row`, `Column` and `Type` are now read-only (immutable) properties. To change one of the properties, a new Address object has to be created


## Range
- Range fields `StartAddress`, and `EndAddress` are now read-only (immutable) properties. To change one of the properties, a new Range object has to be created

## Style (general)
- All (s)RGB values are automatically validated and cast to upper case. If valid hex values are used, no actions are necessary