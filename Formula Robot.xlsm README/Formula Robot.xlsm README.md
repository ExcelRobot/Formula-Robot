# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Formula Robot.xlsm\*\* contains definitions for:

[20 Robot Commands](#command-definitions)<BR>

<BR>

## Available Robot Commands

[Fill](#fill) | [Formula](#formula) | [Paste](#paste) | [Table](#table) | [Translate](#translate)

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
| Macro Expression | <code>[modApplyFilterToArray.ApplyFilterToArray](./VBA/modApplyFilterToArray.bas#L11)([[ActiveCell]],[[ActiveCell.Offset(1,1)]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula |

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

### Convert Formula To Structural Ref

*Replace ActiveCell formula precedency range reference with their structured form like A1:D3 is named range and ActiveCell formula use A1:D3 then it will replace A1:D3 with it's named range name.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

> \*\*Note:\*\* Structured reference means using Table, Named Range or \# for spill formula. Instead of using A1:D3 this will try to find possible structured reference for that range and use that.

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modStructuredReference.ConvertFormulaToStructuredRef](./VBA/modStructuredReference.bas#L41)([[ActiveCell]])</code> |
| User Context Filter | ExcelActiveCellContainsFormula AND ExcelSelectionIsSingleCell |

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

### Select Spill Parent

*Select dynamic array formula cell. If no spill in ActiveCell then do nothing.*

<sup>`@Formula Robot.xlsm` `!VBA Macro Command` `#Formula`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modDynamicFormula.SelectSpillParent](./VBA/modDynamicFormula.bas#L13)()</code> |

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
