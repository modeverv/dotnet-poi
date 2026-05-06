# Protection

dotnet-poi supports both sheet-level and workbook-level protection. Password hashing is not implemented.

## Sheet Protection

```csharp
sheet.protectSheet("password");  // lock the sheet with a password
sheet.protectSheet(null);        // remove protection
```

When protected, users cannot modify locked cells, insert/delete rows or columns, or change sheet structure.

To check protection:

```csharp
var isProtected = sheet.getProtect();  // bool
```

## Cell-Level Locking

By default, all cells are locked when a sheet is protected. To allow editing on specific cells:

```csharp
var style = wb.createCellStyle();
style.setLocked(false);       // cell can be edited even when sheet is protected
cell.setCellStyle(style);
```

## Workbook Protection

```csharp
// Protect workbook structure (prevents add/delete/move sheets)
wb.getSheetAt(0).protectSheet("password");  // applied per-sheet
```

## Limitations

- Password hashing is not implemented. The password string is stored in plain text in the file. Excel will open the file without requiring a password if the hash is absent, but the protection flag is respected.
- Protection on/off works correctly and round-trips.

## Full Runnable Example

See `examples/UsageSamples/Program.cs`:

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/main/examples/UsageSamples)
