Attribute VB_Name = "modPaste"
Option Explicit

Sub PasteExactFormula(rngCopied As Range)
    Dim nRowsCopied As Long
    Dim nColumnsCopied As Long
    Dim rngSelection As Range
    Dim rngActiveCell As Range
    Dim rngTarget As Range
    Dim nAreaCtr As Long
    Dim nRowCtr As Long
    Dim nColCtr As Long
    Dim rngNewSelection As Range
    
    If rngCopied Is Nothing Then Exit Sub
    
    nRowsCopied = rngCopied.Rows.CountLarge
    nColumnsCopied = rngCopied.Columns.CountLarge
    
    Set rngSelection = Selection
    Set rngActiveCell = ActiveCell
    
    Dim nCalcMode As Integer
    nCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    For nAreaCtr = 1 To rngSelection.Areas.Count
        If rngSelection.Areas(nAreaCtr).Cells.CountLarge = 1 Then
            Set rngTarget = rngSelection.Areas(nAreaCtr).Cells.Resize(nRowsCopied, nColumnsCopied)
        Else
            Set rngTarget = rngSelection.Areas(nAreaCtr).Cells
        End If
        For nRowCtr = 1 To rngTarget.Rows.CountLarge
            For nColCtr = 1 To rngTarget.Columns.CountLarge
                rngTarget.Cells(nRowCtr, nColCtr).Formula2 = rngCopied.Cells(((nRowCtr - 1) Mod nRowsCopied) + 1, ((nColCtr - 1) Mod nColumnsCopied) + 1).Formula2
                rngTarget.Cells(nRowCtr, nColCtr).NumberFormat = rngCopied.Cells(((nRowCtr - 1) Mod nRowsCopied) + 1, ((nColCtr - 1) Mod nColumnsCopied) + 1).NumberFormat
                If rngNewSelection Is Nothing Then
                    Set rngNewSelection = rngTarget.Cells(nRowCtr, nColCtr)
                Else
                    Set rngNewSelection = Union(rngNewSelection, rngTarget.Cells(nRowCtr, nColCtr))
                End If
            Next nColCtr
        Next nRowCtr
    Next nAreaCtr
        
    rngNewSelection.Select
    If Not Union(rngNewSelection, rngActiveCell) Is Nothing Then
        rngActiveCell.Activate
    End If
    rngCopied.Copy

    Application.Calculation = nCalcMode
    Application.ScreenUpdating = True


End Sub
