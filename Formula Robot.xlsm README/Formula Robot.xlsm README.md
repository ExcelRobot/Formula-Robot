# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Formula Robot.xlsm\*\* contains definitions for:

[46 Robot Commands](#command-definitions)<BR>

<BR>

## Available Robot Commands

[Fill](#fill) | [Formula](#formula) | [Name](#name) | [Paste](#paste) | [Table](#table) | [Translate](#translate)

### Fill

| Name | Description |
| --- | --- |
| [Auto Fill Down](#auto-fill-down) | Fill formula or value in active cell down using smart automation. |
| [Auto Fill To Right](#auto-fill-to-right) | Fill formula or value in active cell to right using smart automation. |
| [Paste Auto Fill Down](#paste-auto-fill-down) | Fill formula or value from clipboard range down using smart automation. |
| [Paste Auto Fill To Right](#paste-auto-fill-to-right) | Fill formula or value from clipboard range to right using smart automation. |

### Formula

| Name | Description |
| --- | --- |
| [Apply Filter To Array](#apply-filter-to-array) | Create Filter formula based on top row spill range formula. |
| [Auto\-Fit Formula Bar](#auto-fit-formula-bar) | Autofit formula bar height based on formula length so that whole formula is visible. |
| [Compact Formula Format](#compact-formula-format) | Compact formula format of selected cells containing formulas. |
| [Convert Formula To Structural Ref](#convert-formula-to-structural-ref) | Replace ActiveCell formula precedency range reference with their structured form like A1:D3 is named range and ActiveCell formula use A1:D3 then it will replace A1:D3 with it's named range name. |
| [Format Formulas](#format-formulas) | Format formulas from selected cells. It will only change cell formula which has a formula. |
| [Map To Array](#map-to-array) | Convert spill parent cell formula to use Map for all the cell of the spill range. |
| [Paste As References](#paste-as-references) | Paste clipboard range address to active cell. |
| [Paste As Structured References](#paste-as-structured-references) | Paste clipboard range as dynamic address to active cell. |
| [Remove Outer Function](#remove-outer-function) | Remove the outer function from ActiveCell formula e.g. ActiveCell formula \=DROP(SEQUENCE(10,2),2,1) and after running the command it will be \=SEQUENCE(10,2) |
| [Remove Outer Function X 2](#remove-outer-function-x-2) | Run Remove Outer Function command twice. |
| [Remove Outer Function X 3](#remove-outer-function-x-3) | Run Remove Outer Function command three times. |
| [Select Spill Parent](#select-spill-parent) | Select dynamic array formula cell. If no spill in ActiveCell then do nothing. |
| [Toggle Expand Formula Bar](#toggle-expand-formula-bar) | If formula bar height is 1 then autofit it otherwise make it 1. |

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

### Paste

| Name | Description |
| --- | --- |
| [Paste As References](#paste-as-references) | Paste clipboard range address to active cell. |
| [Paste As Structured References](#paste-as-structured-references) | Paste clipboard range as dynamic address to active cell. |
| [Paste Auto Fill Down](#paste-auto-fill-down) | Fill formula or value from clipboard range down using smart automation. |
| [Paste Auto Fill To Right](#paste-auto-fill-to-right) | Fill formula or value from clipboard range to right using smart automation. |
| [Paste Translate Formula](#paste-translate-formula) | Given text in clipboard containing an Excel formula in en\-us format, translates it to local languange, and puts it in active cell. |

### Table

| Name | Description |
| --- | --- |
| [Generate Table Lookup Lambdas](#generate-table-lookup-lambdas) | Generate lambdas for each column of an Excel Table. |

### Translate

| Name | Description |
| --- | --- |
| [Copy Formula To English](#copy-formula-to-english) | Translate ActiveCell formula to en\-us locale and copy to clipboard. |
| [Paste Translate Formula](#paste-translate-formula) | Given text in clipboard containing an Excel formula in en\-us format, translates it to local languange, and puts it in active cell. |

<BR>

## Command Definitions

<BR>

### Apply Filter To Array

*Create Filter formula based on top row spill range formula.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modApplyFilterToArray.ApplyFilterToArray](./VBA/modApplyFilterToArray.bas#L10)([[ActiveCell]],[[ActiveCell.Offset(1,1)]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |

[^Top](#oa-robot-definitions)

<BR>

### Auto Fill Down

*Fill formula or value in active cell down using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.FillDown](./VBA/modFillArray.bas#L30)([[Selection]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea |
| Launch Codes | <ol><li><code>d</code></li><li><code>fd</code></li><li><code>afd</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Auto Fill To Right

*Fill formula or value in active cell to right using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.FillToRight](./VBA/modFillArray.bas#L69)([[Selection]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea |
| Launch Codes | <ol><li><code>r</code></li><li><code>fr</code></li><li><code>afr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Auto\-Fit Formula Bar

*Autofit formula bar height based on formula length so that whole formula is visible.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.AutofitFormulaBar](./VBA/modAutofitFormula.bas#L27)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula OR ExcelActiveCellIsSpillParent OR ExcelActiveCellIsNotEmpty |

[^Top](#oa-robot-definitions)

<BR>

### Cancel Named Range Edit

*Cancel edit named range edit mode and use named range name instead of it's refers to and delete the comment.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CancelNamedRangeEdit](./VBA/modNamedRange.bas#L123)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>cnr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Compact Formula Format

*Compact formula format of selected cells containing formulas.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.FormatFormulas](./VBA/modAutofitFormula.bas#L63)([[Selection]],TRUE)</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Command After | [Auto-Fit Formula Bar](#auto-fit-formula-bar) |

[^Top](#oa-robot-definitions)

<BR>

### Convert Formula To Structural Ref

*Replace ActiveCell formula precedency range reference with their structured form like A1:D3 is named range and ActiveCell formula use A1:D3 then it will replace A1:D3 with it's named range name.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

> \*\*Note:\*\* Structured reference means using Table, Named Range or \# for spill formula. Instead of using A1:D3 this will try to find possible structured reference for that range and use that.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modStructuredReference.ConvertFormulaToStructuredRef](./VBA/modStructuredReference.bas#L40)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Convert Global Named Ranges To Local

*Convert all globally scoped named ranges in selection to local.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ConvertGlobalToLocal](./VBA/modNamedRange.bas#L1363)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Convert Local Named Ranges To Global

*Converts all locally scoped named ranges in selection to global.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ConvertLocalToGlobal](./VBA/modNamedRange.bas#L1344)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Copy Formula To English

*Translate ActiveCell formula to en\-us locale and copy to clipboard.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Translate`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTranslation.CopyFormulaToEnglish](./VBA/modTranslation.bas#L36)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Outputs | <ol><li>Message To User</li><li>Save To Clipboard</li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Global Named Range

*Creates a global named range based on selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameRange](./VBA/modNamedRange.bas#L430)([[Selection]], False)</code> |
| Keyboard Shortcut | <code>ctrl+shift+n</code> |
| Launch Codes | <ol><li><code>cnr</code></li><li><code>n</code></li><li><code>nr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Create Local Named Range

*Creates a local named range based on selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameRange](./VBA/modNamedRange.bas#L430)([[Selection]], True,True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Relative Column Named Range

*Create named range from ActiveCell where it will use relative range reference for row e.g. $A2.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CreateRelativeColumnNamedRange](./VBA/modNamedRange.bas#L213)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Relative Row Named Range

*Create named range from ActiveCell by keeping row absolute e.g A$2.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.CreateRelativeRowNamedRange](./VBA/modNamedRange.bas#L223)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Table Column Local Named Ranges

*Creates a local named range for each table column included in selection in the form \<table name\>\_\<column name\>.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddTableColumnNamedRange](./VBA/modNamedRange.bas#L1990)([[Selection]],True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Create Table Column Named Ranges

*Creates a global named range for each table column included in selection in the form \<table name\>\_\<column name\>.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddTableColumnNamedRange](./VBA/modNamedRange.bas#L1990)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Edit Named Range

*Replace ActiveCell formula which has a named range with that named range RefersTo e.g. A1 formula \= MyName then it will replace MyName with it RefersTo.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.EditNamedRange](./VBA/modNamedRange.bas#L87)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Launch Codes | <code>enr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Expand Named Range

*Expand previously created range reference with current selection. If only one named range reference first cell is the same as the first cell of current selection then it will update reference to use current selection instead of old selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ExpandNamedRange](./VBA/modNamedRange.bas#L2112)([[Selection]])</code> |
| Launch Codes | <code>xnr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Formulas

*Format formulas from selected cells. It will only change cell formula which has a formula.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.FormatFormulas](./VBA/modAutofitFormula.bas#L63)([[Selection]])</code> |
| Launch Codes | <code>ff</code> |

[^Top](#oa-robot-definitions)

<BR>

### Generate Table Lookup Lambdas

*Generate lambdas for each column of an Excel Table.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Table`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTableLookupLambda.GenerateTableLookupLambdas](./VBA/modTableLookupLambda.bas#L14)([[ActiveCell]], [[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Map To Array

*Convert spill parent cell formula to use Map for all the cell of the spill range.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMapToArray.MapToArray](./VBA/modMapToArray.bas#L11)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>map</code></li><li><code>tile</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Name All Table Data Columns

*Find table or named range from ActiveCell and Create named range for all column.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameAllTableDataColumns](./VBA/modNamedRange.bas#L1916)([[ActiveCell]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Local Parameter Cells

*Automatically name each cell in selection based on adjacent labels using local scope.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCells](./VBA/modNamedRange.bas#L257)([[Selection]],True)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells

*Automatically name each cell in selection based on adjacent labels using global scope.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCells](./VBA/modNamedRange.bas#L257)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells As Column\_Row

*Automatically name each cell in selection as \<column label\>\_\<row label\>.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCellsByColumnRow](./VBA/modNamedRange.bas#L330)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Cells As Row\_Column

*Automatically name each cell in selection as \<row label\>\_\<column label\>.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.AddNameToParameterCellsByRowColumn](./VBA/modNamedRange.bas#L307)([[Selection]],False)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Column

*Create named range for each selected column where label will be searched in upper rows only.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameParameterColumn](./VBA/modNamedRange.bas#L159)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Parameter Row

*Create named range for each selected row where label will be searched in left columns only.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameParameterRow](./VBA/modNamedRange.bas#L183)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Name Table Data Column

*Create named range for active column of a table or named range. It will only work on table or named range cells.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.NameTableDataColumn](./VBA/modNamedRange.bas#L1948)([[ActiveCell]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste As References

*Paste clipboard range address to active cell.*

<sup>`@Formula Robot.xlsm` `!Excel Formula Command` `#Paste` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=\[\[Clipboard::Address\]\]</code> |
| Destination Range Address | <code>\[\[ActiveCell\]\]</code> |
| User Context Filter | ClipboardHasExcelData |
| Launch Codes | <code>pr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste As Structured References

*Paste clipboard range as dynamic address to active cell.*

<sup>`@Formula Robot.xlsm` `!Excel Formula Command` `#Paste` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=\[\[Clipboard::DynamicReference\]\]</code> |
| Destination Range Address | <code>\[\[ActiveCell\]\]</code> |
| User Context Filter | ClipboardHasExcelData |
| Launch Codes | <code>psr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Auto Fill Down

*Fill formula or value from clipboard range down using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.PasteFillDown](./VBA/modFillArray.bas#L14)([Clipboard],[[ActiveCell]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea AND ClipboardHasExcelData |
| Launch Codes | <ol><li><code>pd</code></li><li><code>pf</code></li><li><code>pfd</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Auto Fill To Right

*Fill formula or value from clipboard range to right using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.PasteFillToRight](./VBA/modFillArray.bas#L53)([Clipboard],[[ActiveCell]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea AND ClipboardHasExcelData |
| Launch Codes | <ol><li><code>pr</code></li><li><code>pfr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Translate Formula

*Given text in clipboard containing an Excel formula in en\-us format, translates it to local languange, and puts it in active cell.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Paste` `#Translate`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTranslation.PasteTranslateFormula](./VBA/modTranslation.bas#L52)({{Clipboard}},[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellIsEmpty AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Reassign Global Named Range

*Find global scoped named range from label and if exist then update it's reference to use current selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ReAssignGlobalNamedRange](./VBA/modNamedRange.bas#L149)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Reassign Local Named Range

*Find local scoped named range from label and if exist then update it's reference to use current selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.ReAssignLocalNamedRange](./VBA/modNamedRange.bas#L145)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove All Named Ranges In Workbook

*Remove all named ranges from active workbook.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.RemoveAllNamedRanges](./VBA/modNamedRange.bas#L1851)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Named Range From Selection

*Remove all named range from selected cells.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.DeleteNamedRangeOnly](./VBA/modNamedRange.bas#L232)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Named Ranges With Errors

*Remove all named ranges with \#REF\! errors from active workbook.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.DeleteNamedRangesHavingError](./VBA/modNamedRange.bas#L1888)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Outer Function

*Remove the outer function from ActiveCell formula e.g. ActiveCell formula \=DROP(SEQUENCE(10,2),2,1) and after running the command it will be \=SEQUENCE(10,2)*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRemoveOuterFunction.RemoveOuterFunction](./VBA/modRemoveOuterFunction.bas#L12)([[Selection]])</code> |
| Keyboard Shortcut | <code>^+z</code> |
| Command Before | [Select Spill Parent](#select-spill-parent) |
| Launch Codes | <ol><li><code>ro</code></li><li><code>z</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Outer Function X 2

*Run Remove Outer Function command twice.*

<sup>`@Formula Robot.xlsm` `!Sequence Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Commands | <ol><li>[Remove Outer Function](#remove-outer-function)</li><li>[Remove Outer Function](#remove-outer-function)</li></ol> |
| Launch Codes | <code>ro2</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Outer Function X 3

*Run Remove Outer Function command three times.*

<sup>`@Formula Robot.xlsm` `!Sequence Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Commands | <ol><li>[Remove Outer Function](#remove-outer-function)</li><li>[Remove Outer Function](#remove-outer-function)</li><li>[Remove Outer Function](#remove-outer-function)</li></ol> |
| Launch Codes | <code>ro3</code> |

[^Top](#oa-robot-definitions)

<BR>

### Rename Named Range

*Rename the named range associated with the current selection.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.RenameNamedRange](./VBA/modNamedRange.bas#L1199)([[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Save Named Range

*Save ActiveCell formula as named range. It will check cell comment and use that as Named Range name if present.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Name`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modNamedRange.SaveNamedRange](./VBA/modNamedRange.bas#L21)([[ActiveCell]],False)</code> |
| Launch Codes | <code>snr</code> |

[^Top](#oa-robot-definitions)

<BR>

### Select Spill Parent

*Select dynamic array formula cell. If no spill in ActiveCell then do nothing.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDynamicFormula.SelectSpillParent](./VBA/modDynamicFormula.bas#L12)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Expand Formula Bar

*If formula bar height is 1 then autofit it otherwise make it 1.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.ToggleExpandFormulaBar](./VBA/modAutofitFormula.bas#L10)([[ActiveCell]])</code> |
| Keyboard Shortcut | <code>^+u</code> |
| Launch Codes | <code>fb</code> |

[^Top](#oa-robot-definitions)
