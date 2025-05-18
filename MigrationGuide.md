# Migration Guide v2.x to v3.0.0

## Reader

- Workbook.Load() vs WorkbokReader.Load()
- ImportOptions vs ReaderOptions

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