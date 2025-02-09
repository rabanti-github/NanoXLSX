# Migration Guide v2.x to v3.0.0

## Reader

- Workbook.Load() vs WorkbokReader.Load()
- ImportOptions vs ReaderOptions

## Worksheet

- Password handling
  - LegacyPassword is always defined (should never be null) 

### Methods
- Method SetSelectedCells was replaced by AddSelectedCell, RemoveSelectedCells and ClearSelectedCells. Further overload methods were added to the first two methods

## Address
- Address fields `Row`, `Column` and `Type` are now read-only (immutable) properties. To change one of the properties, a new Address object has to be created


## Range
- Range fields `StartAddress`, and `EndAddress` are now read-only (immutable) properties. To change one of the properties, a new Range object has to be created