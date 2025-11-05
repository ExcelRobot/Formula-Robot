# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Formula Robot.xlsm\*\* contains definitions for:

[21 Robot Commands](#command-definitions)<BR>[1 Robot Parameter](#parameter-definitions)<BR>[1 Robot Text](#text-definitions)<BR>

<BR>

## Available Robot Commands

[Fill](#fill) | [Formula](#formula) | [Paste](#paste) | [Search](#search) | [Table](#table) | [Translate](#translate)

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
| [Format Formulas](#format-formulas) | Format formulas from selected cells. It will only change cell formula which has a formula. |
| [Map To Array](#map-to-array) | Convert spill parent cell formula to use Map for all the cell of the spill range. |
| [Paste As References](#paste-as-references) | Paste clipboard range address to active cell. |
| [Remove Outer Function](#remove-outer-function) | Remove the outer function from ActiveCell formula e.g. ActiveCell formula \=DROP(SEQUENCE(10,2),2,1) and after running the command it will be \=SEQUENCE(10,2) |
| [Select Spill Parent](#select-spill-parent) | Select dynamic array formula cell. If no spill in ActiveCell then do nothing. |
| [Show Formula Text](#show-formula-text) | Show formula text of active cell in a specified adjacent cell. |
| [Time Formula](#time-formula) | Wraps formula in active cell with Timer lambda to evaluate calculation performance. |
| [Toggle Expand Formula Bar](#toggle-expand-formula-bar) | If formula bar height is 1 then autofit it otherwise make it 1. |

### Paste

| Name | Description |
| --- | --- |
| [Paste As References](#paste-as-references) | Paste clipboard range address to active cell. |
| [Paste Auto Fill Down](#paste-auto-fill-down) | Fill formula or value from clipboard range down using smart automation. |
| [Paste Auto Fill To Right](#paste-auto-fill-to-right) | Fill formula or value from clipboard range to right using smart automation. |
| [Paste Exact Formula](#paste-exact-formula) | Paste formula exactly as copied and number formats. |
| [Paste Translate Formula](#paste-translate-formula) | Given text in clipboard containing an Excel formula in en\-us format, translates it to local languange, and puts it in active cell. |

### Search

| Name | Description |
| --- | --- |
| [Search Workbook For Specified Functions](#search-workbook-for-specified-functions) | Searches cells, names, and conditional formatting for use of user specified functions and reports findings. |
| [Search Workbook For Volatile Functions](#search-workbook-for-volatile-functions) | Searches cells, names, and conditional formatting for use of volatile functions and reports findings. |

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

## Available Robot Parameters

| Name | Description |
| --- | --- |
| [UserSpecifiedFunctions](#userspecifiedfunctions) | Ask user for a comma seperated functions name for searching in cells, names and conditional formatting. |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [Timer.lambda](#timerlambda) | Returns time required to calculate a formula in seconds. |

<BR>

## Command Definitions

<BR>

### Apply Filter To Array

*Create Filter formula based on top row spill range formula.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modApplyFilterToArray.ApplyFilterToArray](./VBA/modApplyFilterToArray.bas#L11)([[ActiveCell]],[[ActiveCell.Offset(0,1)]])</code> |
| User Context Filter | ExcelActiveCellContainsNonSpillingFormula AND ExcelActiveCellValueIsBoolean |

[^Top](#oa-robot-definitions)

<BR>

### Auto Fill Down

*Fill formula or value in active cell down using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.FillDown](./VBA/modFillArray.bas#L31)([[Selection]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea |
| Launch Codes | <ol><li><code>d</code></li><li><code>fd</code></li><li><code>afd</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Auto Fill To Right

*Fill formula or value in active cell to right using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.FillToRight](./VBA/modFillArray.bas#L70)([[Selection]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea |
| Launch Codes | <ol><li><code>r</code></li><li><code>fr</code></li><li><code>afr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Auto\-Fit Formula Bar

*Autofit formula bar height based on formula length so that whole formula is visible.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.AutofitFormulaBar](./VBA/modAutofitFormula.bas#L28)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula OR ExcelActiveCellIsSpillParent OR ExcelActiveCellIsNotEmpty |

[^Top](#oa-robot-definitions)

<BR>

### Compact Formula Format

*Compact formula format of selected cells containing formulas.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.FormatFormulas](./VBA/modAutofitFormula.bas#L64)([[Selection]],TRUE)</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Command After | [Auto-Fit Formula Bar](#auto-fit-formula-bar) |
| Launch Codes | <code>bo</code> |

[^Top](#oa-robot-definitions)

<BR>

### Copy Formula To English

*Translate ActiveCell formula to en\-us locale and copy to clipboard.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Translate`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTranslation.CopyFormulaToEnglish](./VBA/modTranslation.bas#L37)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Outputs | <ol><li>Message To User</li><li>Save To Clipboard</li></ol> |

<BR>

#### Copy Formula To English \>\> Message To User

*Message to user.*

<sup>`!Message Box Output` </sup>

| Property | Value |
| --- | --- |
| Title | <code>Copy Formula To English</code> |
| Text Before | <code>Copied to clipboard:</code><br><code></code><br><code></code> |

<BR>

#### Copy Formula To English \>\> Save To Clipboard

*Save to clipboard.*

<sup>`!Clipboard Output` </sup>

*No Values*

[^Top](#oa-robot-definitions)

<BR>

### Format Formulas

*Format formulas from selected cells. It will only change cell formula which has a formula.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.FormatFormulas](./VBA/modAutofitFormula.bas#L64)([[Selection]])</code> |
| Launch Codes | <ol><li><code>ff</code></li><li><code>unbo</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Generate Table Lookup Lambdas

*Generate lambdas for each column of an Excel Table.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Table`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTableLookupLambda.GenerateTableLookupLambdas](./VBA/modTableLookupLambda.bas#L15)([[ActiveCell]], [[Selection]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Map To Array

*Convert spill parent cell formula to use Map for all the cell of the spill range.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modMapToArray.MapToArray](./VBA/modMapToArray.bas#L13)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |
| Launch Codes | <ol><li><code>map</code></li><li><code>tile</code></li></ol> |

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

### Paste Auto Fill Down

*Fill formula or value from clipboard range down using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.PasteFillDown](./VBA/modFillArray.bas#L15)([Clipboard],[[ActiveCell]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea AND ClipboardHasExcelData |
| Launch Codes | <ol><li><code>pd</code></li><li><code>pf</code></li><li><code>pfd</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Auto Fill To Right

*Fill formula or value from clipboard range to right using smart automation.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Fill` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFillArray.PasteFillToRight](./VBA/modFillArray.bas#L54)([Clipboard],[[ActiveCell]])</code> |
| User Context Filter | ExcelSelectionIsSingleArea AND ClipboardHasExcelData |
| Launch Codes | <ol><li><code>pr</code></li><li><code>pfr</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Exact Formula

*Paste formula exactly as copied and number formats.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Paste`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modPaste.PasteExactFormula](./VBA/modPaste.bas#L4)([[Clipboard]])</code> |
| User Context Filter | ClipboardHasExcelData |
| Launch Codes | <ol><li><code>pe</code></li><li><code>pef</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Paste Translate Formula

*Given text in clipboard containing an Excel formula in en\-us format, translates it to local languange, and puts it in active cell.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Paste` `#Translate`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modTranslation.PasteTranslateFormula](./VBA/modTranslation.bas#L53)({{Clipboard}},[[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellIsEmpty AND ExcelSelectionIsSingleCell |

[^Top](#oa-robot-definitions)

<BR>

### Remove Outer Function

*Remove the outer function from ActiveCell formula e.g. ActiveCell formula \=DROP(SEQUENCE(10,2),2,1) and after running the command it will be \=SEQUENCE(10,2)*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modRemoveOuterFunction.RemoveOuterFunction](./VBA/modRemoveOuterFunction.bas#L13)([[Selection]])</code> |
| Keyboard Shortcut | <code>^+z</code> |
| Command Before | [Select Spill Parent](#select-spill-parent) |
| Launch Codes | <ol><li><code>ro</code></li><li><code>z</code></li></ol> |

[^Top](#oa-robot-definitions)

<BR>

### Search Workbook For Specified Functions

*Searches cells, names, and conditional formatting for use of user specified functions and reports findings.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Search`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFunctionSearch.SearchFunctions](./VBA/modFunctionSearch.bas#L10)({{[UserSpecifiedFunctions](#userspecifiedfunctions)}},[[NewTableTargetToRight]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Search Workbook For Volatile Functions

*Searches cells, names, and conditional formatting for use of volatile functions and reports findings.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Search`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFunctionSearch.SearchFunctions](./VBA/modFunctionSearch.bas#L10)("NOW, TODAY, RAND, RANDBETWEEN, RANDARRAY, OFFSET, INDIRECT, CELL, INFO",[[NewTableTargetToRight]])</code> |

[^Top](#oa-robot-definitions)

<BR>

### Select Spill Parent

*Select dynamic array formula cell. If no spill in ActiveCell then do nothing.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDynamicFormula.SelectSpillParent](./VBA/modDynamicFormula.bas#L13)()</code> |

[^Top](#oa-robot-definitions)

<BR>

### Show Formula Text

*Show formula text of active cell in a specified adjacent cell.*

<sup>`@Formula Robot.xlsm` `!Excel Formula Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=FORMULATEXT(\[\[ActiveCell\]\])</code> |
| Destination Range Address | <code>\[ActiveCell.Offset({{OffsetXY::ExcelFormula()}})\]</code> |
| Scroll To Destination | ☐Yes ☑No |
| Parameters | <ol><li>[OffsetXY](#show-formula-text--offsetxy)</li></ol> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Outputs | <ol></ol> |
| Launch Codes | <code>sft</code> |

<BR>

#### Show Formula Text \>\> OffsetXY

*X,Y offset string to be inserted into Destination Range address*

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Where would you like to show the formula text?</code> |
| Validation List | <code>Above active cell,"\-1,0"</code><br><code>Two above active cell,"\-2,0"</code><br><code>Right of active cell,"0,1"</code><br><code>Below active cell,"1,0"</code> |
| Data Type | String |
| Default Value | <code>\-1,0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Time Formula

*Wraps formula in active cell with Timer lambda to evaluate calculation performance.*

<sup>`@Formula Robot.xlsm` `!Excel Formula Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Formula | <code>\=Timer(\[\[ActiveCell::Formula\]\],{{Include\_Output}})</code> |
| Formula Dependencies | [Timer.lambda](#timerlambda) |
| Parameters | <ol><li>[Include_Output](#time-formula--include_output)</li></ol> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |
| Outputs | <ol></ol> |

<BR>

#### Time Formula \>\> Include\_Output

*Ask user whether to include output.*

<sup>`!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Do you want to include the formula output with the timer results?</code> |
| Validation List | <code>"No, return time only.", FALSE</code><br><code>"Yes, return time and formula output.", TRUE</code> |
| Default Value | <code>FALSE</code> |

[^Top](#oa-robot-definitions)

<BR>

### Toggle Expand Formula Bar

*If formula bar height is 1 then autofit it otherwise make it 1.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modAutofitFormula.ToggleExpandFormulaBar](./VBA/modAutofitFormula.bas#L11)([[ActiveCell]])</code> |
| Keyboard Shortcut | <code>^+u</code> |
| Launch Codes | <code>fb</code> |

[^Top](#oa-robot-definitions)

<BR>

## Parameter Definitions

<BR>

### UserSpecifiedFunctions

*Ask user for a comma seperated functions name for searching in cells, names and conditional formatting.*

<sup>`@Formula Robot.xlsm` `!Input Parameter` </sup>

| Property | Value |
| --- | --- |
| Prompt | <code>Specify functions to search for (example: LET, LAMBDA):</code> |
| Data Type | String |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### Timer.lambda

*Returns time required to calculate a formula in seconds.*

<sup>`@Formula Robot.xlsm` `!Excel Name Text` </sup>

> \*\*Note:\*\* Because this lambda doesn't start with \=LAMBDA(, it can't be edited using Lambda Robot.

| Property | Value |
| --- | --- |
| Text | [Timer.lambda](<./Text/Timer.lambda.txt>) |
| Value | <code>Timer \= LET(\_StartTimer, NOW(),</code><br><code> LAMBDA(formula,\[include\_output\], LET(</code><br><code> \_Timing, TEXT(NOW() \- \_StartTimer, "\[s\].000\\s"),</code><br><code> \_Result, IF(include\_output,IFERROR(VSTACK(\_Timing, formula), ""),\_Timing),</code><br><code> \_Result</code><br><code> ))</code><br><code>);</code> |
| Content Type | ExcelLambda |
| Location | <code>Timer</code> |

[^Top](#oa-robot-definitions)
