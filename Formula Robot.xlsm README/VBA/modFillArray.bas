Attribute VB_Name = "modFillArray"
Option Explicit

Private FillDownUndoColl As Collection
Private FillToRightUndoColl As Collection

'*******************************************************
' Fill Down
' Available Commands:
' 1. Auto Fill Down
' 2. Paste Fill Down
'*******************************************************

Public Sub PasteFillDown(ByVal ClipboardRange As Range, ByVal DestinationRange As Range)
    
    If ClipboardRange Is Nothing Then Exit Sub
    If DestinationRange Is Nothing Then Exit Sub
    CopyPasteIfNotSameFirstCell ClipboardRange, DestinationRange
    Set DestinationRange = DestinationRange.Resize(ClipboardRange.Rows.Count, ClipboardRange.Columns.Count)
    FillDown DestinationRange
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Auto-Fill Down
' Description:            Fill down formula or value(dyanmic array formula or normal formula or native fill down).
' Macro Expression:       modFillArray.FillDown([Selection],[Selection])
' Generated:              11/02/2023 12:03 PM
'----------------------------------------------------------------------------------------------------
Public Sub FillDown(ByVal FormulaCell As Range)
    
    Dim Helper As FillDownHelper
    Set Helper = New FillDownHelper
    Helper.FillDown FormulaCell
    Set FillDownUndoColl = Helper.FillDownUndoColl
    
End Sub

Public Sub FillDown_Undo()
    
    UndoOperation FillDownUndoColl
    Set FillDownUndoColl = Nothing
    
End Sub

'*******************************************************
' Fill Array To Right
' Available Commands:
' 1. Auto Fill To Right
' 2. Paste Fill To Right
'*******************************************************

Public Sub PasteFillToRight(ByVal ClipboardRange As Range, ByVal DestinationRange As Range)
        
    If ClipboardRange Is Nothing Then Exit Sub
    If DestinationRange Is Nothing Then Exit Sub
    CopyPasteIfNotSameFirstCell ClipboardRange, DestinationRange
    Set DestinationRange = DestinationRange.Resize(ClipboardRange.Rows.Count, ClipboardRange.Columns.Count)
    FillToRight DestinationRange
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Auto Fill To Right
' Description:            Fill to right by using either Array formula behavior or vba smart fill down or native excel fill down.
' Macro Expression:       modFillArray.FillToRight([Selection],[Selection])
' Generated:              11/02/2023 12:07 PM
'----------------------------------------------------------------------------------------------------
Public Sub FillToRight(ByVal FormulaCell As Range)
    
    Dim Helper As FillToRightHelper
    Set Helper = New FillToRightHelper
    Helper.FillToRight FormulaCell
    Set FillToRightUndoColl = Helper.FillToRightUndoColl
    
End Sub

Public Sub FillToRight_Undo()
    
    UndoOperation FillToRightUndoColl
    Set FillToRightUndoColl = Nothing

End Sub


'**************************
' Helper Sub/Function
'**************************

Public Sub UndoOperation(CellColl As Collection)
    
    If CellColl Is Nothing Then Exit Sub
    If CellColl.Count = 0 Then Exit Sub
    
    Dim CurrentItem As UndoHandler
    For Each CurrentItem In CellColl
        CurrentItem.Undo
    Next CurrentItem
    
End Sub

Public Function GenerateFillWithSequence(ByVal StartFormula As String _
                                         , ByVal ValidCells As Collection _
                                          , ByVal TypeOfFill As FillType _
                                           , ByVal FormulaName As String) As String
    
    Dim StepNamePrefix As String
    StepNamePrefix = GetStepNamePrefix(StartFormula, ValidCells.Count)
    
    If ValidCells.Count = 1 Then
        GenerateFillWithSequence = GenerateFillIfOneRef(StartFormula, ValidCells, TypeOfFill, FormulaName)
        Exit Function
    End If
    
    Dim IsTile As Boolean
    IsTile = (FormulaName = TILE_FX_NAME)
    
    Dim LetPart As String
    LetPart = LET_AND_OPEN_PAREN
    Dim CurrentItemIndex As Long
    Dim CurrentItem As PrecedencyInfo
    For CurrentItemIndex = 1 To ValidCells.Count
        Set CurrentItem = ValidCells.Item(CurrentItemIndex)
        
        Dim ParamName As String
        ParamName = GetParamNameFromCounter(StepNamePrefix, CurrentItemIndex, ValidCells.Count)
    
        Dim ColOrRowIndex As Long
        ColOrRowIndex = CurrentItem.ColOrRowIndex
        
        If TypeOfFill = Fill_DOWN Then
            ' if only one col then Index is perfect.
            If CurrentItem.NameInFormulaRange.Columns.Count = 1 Then
                LetPart = LetPart & ParamName & LIST_SEPARATOR _
                          & ONE_SPACE & INDEX_FX_NAME & FIRST_PARENTHESIS_OPEN _
                          & IIf(IsTile, CurrentItem.AbsRangeRef, CurrentItem.RangeRef) _
                          & LIST_SEPARATOR & "n" _
                          & LIST_SEPARATOR & ColOrRowIndex & FIRST_PARENTHESIS_CLOSE _
                          & LIST_SEPARATOR
            Else
                LetPart = LetPart & ParamName & LIST_SEPARATOR & ONE_SPACE _
                          & CHOOSEROWS_FX_NAME & FIRST_PARENTHESIS_OPEN _
                          & IIf(IsTile, CurrentItem.AbsChoosePartFormula, CurrentItem.ChoosePartFormula) _
                          & LIST_SEPARATOR _
                          & "n" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
            End If
            
        ElseIf TypeOfFill = FILL_TO_RIGHT Then
            If CurrentItem.NameInFormulaRange.Rows.Count = 1 Then
                LetPart = LetPart & ParamName & LIST_SEPARATOR _
                          & ONE_SPACE & INDEX_FX_NAME & FIRST_PARENTHESIS_OPEN _
                          & IIf(IsTile, CurrentItem.AbsRangeRef, CurrentItem.RangeRef) _
                          & LIST_SEPARATOR & ColOrRowIndex _
                          & LIST_SEPARATOR & "n" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
                          
            Else
                LetPart = LetPart & ParamName & LIST_SEPARATOR & ONE_SPACE _
                          & CHOOSECOLS_FX_NAME & FIRST_PARENTHESIS_OPEN _
                          & IIf(IsTile, CurrentItem.AbsChoosePartFormula, CurrentItem.ChoosePartFormula) _
                          & LIST_SEPARATOR _
                          & "n" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
            End If
        End If
        StartFormula = ReplaceTokenWithNewToken(StartFormula, CurrentItem.NameInFormula, ParamName)
        
    Next CurrentItemIndex
    
    Set CurrentItem = ValidCells.Item(1)
    Dim TileStartPart As String
    If TypeOfFill = Fill_DOWN Then
        TileStartPart = "=" & FormulaName & "(" & SEQUENCE_FX_NAME & FIRST_PARENTHESIS_OPEN _
                        & ROWS_FX_NAME & FIRST_PARENTHESIS_OPEN _
                        & IIf(IsTile, CurrentItem.AbsRangeRef, CurrentItem.RangeRef) _
                        & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                        & LAMBDA_AND_OPEN_PAREN & "n" & LIST_SEPARATOR
                        
    ElseIf TypeOfFill = FILL_TO_RIGHT Then
        TileStartPart = "=" & FormulaName & "(" & SEQUENCE_FX_NAME & FIRST_PARENTHESIS_OPEN _
                        & "1" & LIST_SEPARATOR & COLUMNS_FX_NAME _
                        & FIRST_PARENTHESIS_OPEN & IIf(IsTile, CurrentItem.AbsRangeRef, CurrentItem.RangeRef) _
                        & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                        & LAMBDA_AND_OPEN_PAREN & "n" & LIST_SEPARATOR
    End If
    
    ' Three close paren because One is for LET, Another one is for Lambda and Third one is for TILE
    GenerateFillWithSequence = TileStartPart & LetPart & RemoveStartingEqualSign(StartFormula) _
                               & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    
End Function

Public Function GetStepNamePrefix(ByVal FormulaText As String, ValidCellsCount As Long) As String
    
    Dim PreviousStepsName As Collection
    Set PreviousStepsName = GetAllParamAndStepName(FormulaText)
    
    Dim StepNamePrefix As String
    StepNamePrefix = "x"
    Do While IsAnyItemExistInCollection(PreviousStepsName, GetAllInitialStepNames(StepNamePrefix, ValidCellsCount))
        StepNamePrefix = StepNamePrefix & "x"
    Loop
    
    GetStepNamePrefix = StepNamePrefix
    
End Function

Private Function GetAllInitialStepNames(ByVal StartStepName As String _
                                        , ByVal TotalStep As Long) As Variant
    
    Dim Result As Variant
    ReDim Result(1 To TotalStep) As String
    
    Dim Counter As Long
    For Counter = 1 To TotalStep
        Result(Counter) = GetParamNameFromCounter(StartStepName, Counter, TotalStep)
    Next Counter
    GetAllInitialStepNames = Result
    
End Function

Private Function IsAnyItemExistInCollection(ByVal SearchInColl As Collection, SearchKeys As Variant) As Boolean
    
    Dim Key As Variant
    For Each Key In SearchKeys
        If IsExistInCollection(SearchInColl, CStr(Key)) Then
            IsAnyItemExistInCollection = True
            Exit Function
        End If
    Next Key
    
    IsAnyItemExistInCollection = False
    
End Function

Private Function GenerateFillIfOneRef(ByVal StartFormula As String _
                                      , ByVal ValidCells As Collection _
                                       , ByVal TypeOfFill As FillType _
                                        , ByVal FormulaName As String) As String

    If ValidCells.Count <> 1 Then Exit Function
    
    Dim CurrentItem As PrecedencyInfo
    Set CurrentItem = ValidCells.Item(1)
    
    Dim StepName As String
    StepName = GetStepNamePrefix(StartFormula, ValidCells.Count)
    
    Dim ChoosePart As String
    If FormulaName = TILE_FX_NAME Then
        ChoosePart = CurrentItem.AbsChoosePartFormula
    Else
        ChoosePart = CurrentItem.ChoosePartFormula
    End If
    
    Dim Formula As String
    GenerateFillIfOneRef = "=" & FormulaName & "(" & ChoosePart & LIST_SEPARATOR _
                           & LAMBDA_AND_OPEN_PAREN & StepName & LIST_SEPARATOR _
                           & RemoveStartingEqualSign(ReplaceTokenWithNewToken(StartFormula, CurrentItem.NameInFormula, StepName)) _
                           & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE

End Function

Public Function MaxColCount(ByVal ValidCells As Collection) As Long
    
    Dim Count As Long
    Dim CurrentItem As PrecedencyInfo
    For Each CurrentItem In ValidCells
        If Count < CurrentItem.ColCount Then Count = CurrentItem.ColCount
    Next CurrentItem
    MaxColCount = Count
    
End Function

Public Function ConsistentFormulaCount(ByVal CheckOnRange As Range, IsCheckOnCol As Boolean) As Long
    
    If IsNothing(CheckOnRange) Then
        ConsistentFormulaCount = 0
        Exit Function
    End If
    
    Dim TopCellFormula As String
    TopCellFormula = CheckOnRange.Cells(1).Formula2R1C1
    
    Dim Counter As Long
    Dim Formulas As Variant
    If IsCheckOnCol Then
        Formulas = CheckOnRange.Columns(1).Formula2R1C1
    Else
        Formulas = CheckOnRange.Rows(1).Formula2R1C1
    End If
    Dim Formula As Variant
    For Each Formula In Formulas
        If Formula = TopCellFormula Then
            Counter = Counter + 1
        Else
            Exit For
        End If
    Next Formula
    
    ConsistentFormulaCount = Counter
    
End Function

Public Sub CopyPasteIfNotSameFirstCell(ByVal ClipboardRange As Range, ByVal DestinationRange As Range)
    
    If ClipboardRange Is Nothing Then Exit Sub
    If Not modUtility.IsStartCellSame(ClipboardRange, DestinationRange) Then
        With ClipboardRange
            .Copy DestinationRange
'            Set DestinationRange = DestinationRange.Resize(.Rows.Count, .Columns.Count)
'            Dim RowIndex As Long
'            For RowIndex = 1 To .Rows.Count
'                Dim ColIndex As Long
'                For ColIndex = 1 To .Columns.Count
'                    DestinationRange.Cells(RowIndex, ColIndex).Formula2 = .Cells(RowIndex, ColIndex).Formula2
'                Next ColIndex
'            Next RowIndex
        End With
    End If

End Sub


