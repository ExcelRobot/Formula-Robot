Attribute VB_Name = "modAutofitFormula"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Toggle Expand Formula Bar
' Description:            If formula bar height is 1 then autofit it otherwise make it 1.
' Macro Expression:       ToggleExpandFormulaBar([ActiveCell])
' Generated:              07/16/2023 02:22 PM
'----------------------------------------------------------------------------------------------------
Public Sub ToggleExpandFormulaBar(ByVal FormulaCell As Range)
    
    If Application.FormulaBarHeight = 1 Then
        AutofitFormulaBar FormulaCell
    Else
        Application.FormulaBarHeight = 1
    End If
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Autofit Formula Bar
' Description:            Autofit formula bar height based on formula length so that whole formula is visible.
' Macro Expression:       AutofitFormulaBar([ActiveCell])
' Generated:              07/16/2023 02:17 PM
'----------------------------------------------------------------------------------------------------

Public Sub AutofitFormulaBar(ByVal FormulaCell As Range)
    
    ' Constants for the minimum and maximum height of the formula bar
    Const MIN_HEIGHT As Long = 4
    Const MAX_HEIGHT As Long = 10
    
    ' Calculate the number of new lines in the cell's formula
    Dim NewLineCount As Long
    NewLineCount = Len(FormulaCell.Formula) - Len(VBA.Replace(FormulaCell.Formula, Chr$(10), vbNullString)) + 1

    On Error GoTo TryOnceAgain
    ' Adjust the height of the formula bar based on the number of new lines
    ' If the number of lines is less than the minimum height, set it to the minimum height
    If NewLineCount < MIN_HEIGHT Then
        Application.FormulaBarHeight = MIN_HEIGHT
        ' If the number of lines is more than the maximum height, set it to the maximum height
    ElseIf NewLineCount > MAX_HEIGHT Then
        Application.FormulaBarHeight = MAX_HEIGHT
        ' If the number of lines is between the minimum and maximum heights, set the height equal to the number of lines
    Else
        Application.FormulaBarHeight = NewLineCount
    End If
    Exit Sub
    
TryOnceAgain:
    
    ' After openning excel and before activating VBE if we try to run this then it doesn't work for the first time.
    ' After using Resume it may not work too. But trying.
    ' But after that every time we run this command then it will work.
    Dim ErrorCount As Long
    ErrorCount = ErrorCount + 1
    If ErrorCount = 1 Then Resume
    Err.Clear
    
End Sub

Public Sub FormatFormulas(SelectionRange As Range, Optional CompactConfig As Boolean = False)
    
    On Error Resume Next
    Dim FormulaCells As Range
    Set FormulaCells = Intersect(SelectionRange, SelectionRange.SpecialCells(xlCellTypeFormulas))
    If FormulaCells Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    
    
    Dim PreviousState As XlCalculation
    PreviousState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim CurrentCell As Range
    For Each CurrentCell In FormulaCells.Cells
        CurrentCell.Formula2 = FormatFormula(ReplaceInvalidCharFromFormulaWithValid(CurrentCell.Formula2), CompactConfig)
    Next CurrentCell
    
    Application.Calculation = PreviousState
    On Error GoTo 0
    
End Sub

