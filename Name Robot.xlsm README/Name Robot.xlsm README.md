# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Name Robot.xlsm\*\* contains definitions for:

[26 Robot Commands](#command-definitions)<BR>

<BR>

## Available Robot Commands

[Name](#name)

### Name

| Name | Description |
| --- | --- |
| [Cancel Named Range Edit](#cancel-named-range-edit) | Cancel edit named range edit mode and use named range name instead of it's refers to and delete the comment. |
| [Convert Global Named Ranges To Local](#convert-global-named-ranges-to-local) | Convert all globally scoped named ranges in selection to local. |
| [Convert Local Named Ranges To Global](#convert-local-named-ranges-to-global) | Converts all locally scoped named ranges in selection to global. |
| [Create Global Named Range](#create-global-named-range) | Creates a global named range based on selection. |
| [Create Local Named Range](#create-local-named-range) | Creates a local named range based on selection. |
| [Create Relative Column Named Range](#create-relative-column-named-range) | Create named range from ActiveCell where it will use relative range reference for row e.g. $A2. |
| [Create Relative Row Named Range](#create-relative-row-named-range) | Create named range from ActiveCell by keeping row absolute e.g A$2. |
| [Create Table Column Local Named Ranges](#create-table-column-local-named-ranges) | Creates a local named range for each table column included in selection in the form \<table name\>\_\<column name\>. |
| [Create Table Column Named Ranges](#create-table-column-named-ranges) | Creates a global named range for each table column included in selection in the form \<table name\>\_\<column name\>. |
| [Edit Named Range](#edit-named-range) | Replace ActiveCell formula which has a named range with that named range RefersTo e.g. A1 formula \= MyName then it will replace MyName with it RefersTo. |
| [Expand Named Range](#expand-named-range) | Expand previously created range reference with current selection. If only one named range reference first cell is the same as the first cell of current selection then it will update reference to use current selection instead of old selection. |
| [Name All Table Data Columns](#name-all-table-data-columns) | Find table or named range from ActiveCell and Create named range for all column. |
| [Name Local Parameter Cells](#name-local-parameter-cells) | Automatically name each cell in selection based on adjacent labels using local scope. |
| [Name Parameter Cells](#name-parameter-cells) | Automatically name each cell in selection based on adjacent labels using global scope. |
| [Name Parameter Cells As Column\_Row](#name-parameter-cells-as-column_row) | Automatically name each cell in selection as \<column label\>\_\<row label\>. |
| [Name Parameter Cells As Row\_Column](#name-parameter-cells-as-row_column) | Automatically name each cell in selection as \<row label\>\_\<column label\>. |
| [Name Parameter Column](#name-parameter-column) | Create named range for each selected column where label will be searched in upper rows only. |
| [Name Parameter Row](#name-parameter-row) | Create named range for each selected row where label will be searched in left columns only. |
| [Name Table Data Column](#name-table-data-column) | Create named range for active column of a table or named range. It will only work on table or named range cells. |
| [Reassign Global Named Range](#reassign-global-named-range) | Find global scoped named range from label and if exist then update it's reference to use current selection. |
| [Reassign Local Named Range](#reassign-local-named-range) | Find local scoped named range from label and if exist then update it's reference to use current selection. |
| [Remove All Named Ranges In Workbook](#remove-all-named-ranges-in-workbook) | Remove all named ranges from active workbook. |
| [Remove Named Range From Selection](#remove-named-range-from-selection) | Remove all named range from selected cells. |
| [Remove Named Ranges With Errors](#remove-named-ranges-with-errors) | Remove all named ranges with \#REF\! errors from active workbook. |
| [Rename Named Range](#rename-named-range) | Rename the named range associated with the current selection. |
| [Save Named Range](#save-named-range) | Save ActiveCell formula as named range. It will check cell comment and use that as Named Range name if present. |

<BR>

## Command Definitions

<BR>

### Cancel Named Range Edit

*Cancel edit named range edit mode and use named range name instead of it's refers to and delete the comment.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CancelNamedRangeEdit](./VBA/modNamedRange.bas#L124)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>cnr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert Global Named Ranges To Local

*Convert all globally scoped named ranges in selection to local.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ConvertGlobalToLocal](./VBA/modNamedRange.bas#L1364)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert Local Named Ranges To Global

*Converts all locally scoped named ranges in selection to global.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ConvertLocalToGlobal](./VBA/modNamedRange.bas#L1345)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Global Named Range

*Creates a global named range based on selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameRange](./VBA/modNamedRange.bas#L431)([[Selection]], False)</code> |
| Keyboard Shortcut | <code>ctrl+shift+n</code> |
| Launch Codes | <ol><li><code>cnr</code></li><li><code>n</code></li><li><code>nr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Local Named Range

*Creates a local named range based on selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameRange](./VBA/modNamedRange.bas#L431)([[Selection]], True,True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Relative Column Named Range

*Create named range from ActiveCell where it will use relative range reference for row e.g. $A2.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CreateRelativeColumnNamedRange](./VBA/modNamedRange.bas#L214)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Relative Row Named Range

*Create named range from ActiveCell by keeping row absolute e.g A$2.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CreateRelativeRowNamedRange](./VBA/modNamedRange.bas#L224)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Table Column Local Named Ranges

*Creates a local named range for each table column included in selection in the form \<table name\>\_\<column name\>.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddTableColumnNamedRange](./VBA/modNamedRange.bas#L1991)([[Selection]],True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Table Column Named Ranges

*Creates a global named range for each table column included in selection in the form \<table name\>\_\<column name\>.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddTableColumnNamedRange](./VBA/modNamedRange.bas#L1991)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Edit Named Range

*Replace ActiveCell formula which has a named range with that named range RefersTo e.g. A1 formula \= MyName then it will replace MyName with it RefersTo.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.EditNamedRange](./VBA/modNamedRange.bas#L88)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>enr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Expand Named Range

*Expand previously created range reference with current selection. If only one named range reference first cell is the same as the first cell of current selection then it will update reference to use current selection instead of old selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ExpandNamedRange](./VBA/modNamedRange.bas#L2113)([[Selection]])</code> |
| Launch Codes | <code>xnr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name All Table Data Columns

*Find table or named range from ActiveCell and Create named range for all column.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameAllTableDataColumns](./VBA/modNamedRange.bas#L1917)([[ActiveCell]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Local Parameter Cells

*Automatically name each cell in selection based on adjacent labels using local scope.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCells](./VBA/modNamedRange.bas#L258)([[Selection]],True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells

*Automatically name each cell in selection based on adjacent labels using global scope.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCells](./VBA/modNamedRange.bas#L258)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells As Column\_Row

*Automatically name each cell in selection as \<column label\>\_\<row label\>.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCellsByColumnRow](./VBA/modNamedRange.bas#L331)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells As Row\_Column

*Automatically name each cell in selection as \<row label\>\_\<column label\>.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCellsByRowColumn](./VBA/modNamedRange.bas#L308)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Column

*Create named range for each selected column where label will be searched in upper rows only.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameParameterColumn](./VBA/modNamedRange.bas#L160)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Row

*Create named range for each selected row where label will be searched in left columns only.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameParameterRow](./VBA/modNamedRange.bas#L184)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Table Data Column

*Create named range for active column of a table or named range. It will only work on table or named range cells.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameTableDataColumn](./VBA/modNamedRange.bas#L1949)([[ActiveCell]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Reassign Global Named Range

*Find global scoped named range from label and if exist then update it's reference to use current selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ReAssignGlobalNamedRange](./VBA/modNamedRange.bas#L150)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Reassign Local Named Range

*Find local scoped named range from label and if exist then update it's reference to use current selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ReAssignLocalNamedRange](./VBA/modNamedRange.bas#L146)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove All Named Ranges In Workbook

*Remove all named ranges from active workbook.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.RemoveAllNamedRanges](./VBA/modNamedRange.bas#L1852)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Named Range From Selection

*Remove all named range from selected cells.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.DeleteNamedRangeOnly](./VBA/modNamedRange.bas#L233)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Named Ranges With Errors

*Remove all named ranges with \#REF\! errors from active workbook.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.DeleteNamedRangesHavingError](./VBA/modNamedRange.bas#L1889)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Rename Named Range

*Rename the named range associated with the current selection.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.RenameNamedRange](./VBA/modNamedRange.bas#L1200)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Save Named Range

*Save ActiveCell formula as named range. It will check cell comment and use that as Named Range name if present.*

<sup>`@Name Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.SaveNamedRange](./VBA/modNamedRange.bas#L22)([[ActiveCell]],False)</code> |
| Launch Codes | <code>snr</code> |

[^Top](#oa-robot-definitions)
