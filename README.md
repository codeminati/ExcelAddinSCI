Excel Add-in (VSTO / Office Add-in)

Objective:
Create an Excel Add-in that enhances user productivity by sanitizing cell content and enabling easy reversion to original data. This exercise tests knowledge of VSTO and Office integration.

Features:

Custom Ribbon Tab with Two Buttons:

Button 1 – "Convert to Alphanumeric":
Removes all non-alphanumeric characters from the selected cells.

Button 2 – "Revert to Original":
Restores the values of the selected cells prior to the last conversion.

Tools and Technologies:

Visual Studio with Office/SharePoint Development Tools

.NET Framework or .NET 6+ with VSTO support

Excel VSTO Add-in Project Template

Implementation Details:

When Button 1 is clicked:

Iterate through selected range.

Store original values in a Dictionary<CellAddress, Value>.

Use regex to remove non-alphanumeric characters.

Apply sanitized values back to the cells.

When Button 2 is clicked:

Check if original values exist for the selected cells.

If found, restore them to their original state.

State Management:

Uses an in-memory dictionary to hold original cell values.

The dictionary persists as long as the workbook is open and the add-in is active.

Edge Cases Handled:

Empty cells are skipped.

Numeric-only cells are preserved.

Only visible cells in selection are processed.

Example Flow:

User selects a range.

Clicks “Convert to Alphanumeric” → updates text.

Clicks “Revert to Original” → restores prior text.

Testing Instructions:

Test with text, numbers, special characters, and formulas.

Ensure reversion works without errors.

Check behavior across multiple selections.
