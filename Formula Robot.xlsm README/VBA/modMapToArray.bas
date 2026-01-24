Attribute VB_Name = "modMapToArray"
Option Explicit
Option Private Module

Private MapToArrayUndoColl As Collection

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Map To Array
' Description:            Convert Spill parent cell formula to use Map for all the cell of the spill range.
' Macro Expression:       modMapToArray.MapToArray([ActiveCell])
' Generated:              10/22/2023 04:09 PM
'----------------------------------------------------------------------------------------------------
Public Sub MapToArray(ByVal FormulaCell As Range, Optional ByVal PlaceFormulaToCell As Range)
    
    Const COMMAND_NAME As String = "Map To Array Command"
    If Not FormulaCell.HasFormula Then
        MsgBox "No formula found in cell: " & FormulaCell.Address, vbCritical + vbInformation, COMMAND_NAME
        Exit Sub
    End If
    
    ' Only Try on the first cell.
    Set FormulaCell = FormulaCell.Cells(1)
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = GetDirectPrecedents(FormulaCell.Formula2, FormulaCell.Worksheet)
    
    If Not IsArray(DirectPrecedents) Then
        MsgBox "No direct precedent cell has been found in the formula.", vbCritical + vbInformation, COMMAND_NAME
        Exit Sub
    End If
    
    If PlaceFormulaToCell Is Nothing Then Set PlaceFormulaToCell = FormulaCell
    Set PlaceFormulaToCell = PlaceFormulaToCell.Cells(1)
    
    Set MapToArrayUndoColl = New Collection
    
    Dim ValidSpillParentCellsForMap As New Collection
    
    Dim ValidMultiCellsRef As New Collection
    
    Dim MaxValidCellsRowCount As Long
    Dim MaxValidCellsColCount As Long
    
    Dim CurrentPrecedency As Variant
    Dim CurrentRange As Range
    For Each CurrentPrecedency In DirectPrecedents
        Dim PrecedentCellAsText As String
        PrecedentCellAsText = CStr(CurrentPrecedency)
        Set CurrentRange = RangeResolver.GetRangeForDependency(PrecedentCellAsText, FormulaCell)
        If Not IsNothing(CurrentRange) Then
            If CurrentRange.Cells(1).HasSpill Then
                If IsSpillParent(CurrentRange) Then
                    ' If it is in spill parent then it could be One Spill parent or it can be in chooserows or choose cols. So update all three.
                    modUtility.UpdateValidCells ValidSpillParentCellsForMap, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE, Nothing
                    
                    If MaxValidCellsRowCount < CurrentRange.SpillingToRange.Rows.CountLarge Then
                        MaxValidCellsRowCount = CurrentRange.SpillingToRange.Rows.CountLarge
                    End If
                    
                    If MaxValidCellsColCount < CurrentRange.SpillingToRange.Columns.CountLarge Then
                        MaxValidCellsColCount = CurrentRange.SpillingToRange.Columns.CountLarge
                    End If
                
                ElseIf CurrentRange.Cells(1).HasSpill And CurrentRange.Cells.CountLarge > 1 And Not Text.IsEndsWith(PrecedentCellAsText, HASH_SIGN) Then
                    modUtility.UpdateValidCells ValidMultiCellsRef, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE, Nothing
                End If
            End If
        End If
    Next CurrentPrecedency
    
    If ValidSpillParentCellsForMap.Count = 0 And ValidMultiCellsRef.Count <> 1 Then
        If IsStartCellSame(FormulaCell, PlaceFormulaToCell) Then
            MsgBox "No Valid cell has been found to do map operation.", vbCritical + vbInformation, COMMAND_NAME
            Exit Sub
        Else
            PlaceFormulaToCell.Formula2 = FormulaCell.Formula2
            Exit Sub
        End If
    End If
    
    Dim FullFormula As String
    If ValidSpillParentCellsForMap.Count > 0 Then
        FullFormula = GetFormulaForMapToArray(ValidSpillParentCellsForMap, FormulaCell, FormulaCell.Formula2, MaxValidCellsRowCount, MaxValidCellsColCount)
    ElseIf IsValidForUntile(ValidMultiCellsRef) Then
        FullFormula = GenerateUntileFormula(ValidMultiCellsRef, FormulaCell)
    Else
        Exit Sub
    End If
    
    AddCustomLambdaIfNeeded FormulaCell.Worksheet.Parent, FullFormula
    MapToArrayUndoColl.Add UndoHandler.Create(DYNAMIC_ARRAY_VERSION, PlaceFormulaToCell, PlaceFormulaToCell.Formula2)
    
    If FullFormula <> vbNullString Then
        
        Dim CopyFromRange As Range
        
        If FormulaCell.HasSpill Then
            Set CopyFromRange = FormulaCell.Cells(1).SpillingToRange
        Else
            Set CopyFromRange = FormulaCell
        End If
        
        PlaceFormulaToCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FullFormula)
        CopyFormatToWholeSpillRange CopyFromRange, PlaceFormulaToCell
    End If
    
    AssingOnUndo "MapToArray"
    
End Sub

Private Sub CopyFormatToWholeSpillRange(ByVal CopyFromRange As Range, ByVal PasteToRange As Range)
        
        
    If IsNotNothing(CopyFromRange) And IsNotNothing(PasteToRange) Then
        
        Dim PrevSelection As Variant
        Set PrevSelection = Selection
        Dim CutCopyMode As XlCutCopyMode
        CutCopyMode = Application.CutCopyMode
        CopyFromRange.Copy
        PasteToRange.Cells(1).SpillingToRange.PasteSpecial xlPasteFormats
        Application.CutCopyMode = CutCopyMode
        If Not PrevSelection Is Nothing Then PrevSelection.Select
        
    End If
    
End Sub

Private Function GenerateUntileFormula(ByVal ValidMultiCellsRef As Collection, FormulaCell As Range) As String
    
    Dim CurrentPrecedent As PrecedencyInfo
    Set CurrentPrecedent = ValidMultiCellsRef.Item(1)
    
    Dim Formula As String
    With CurrentPrecedent.NameInFormulaRange
        
        Formula = ReplaceTokenWithNewToken(FormulaCell.Formula2, CurrentPrecedent.NameInFormula, "x")
        Formula = "=" & UN_TILE_FN_NAME & "(" & .Cells(1).Address(False, False) & "#," _
                  & .Rows.CountLarge & " ," & .Columns.CountLarge _
                  & ",LAMBDA(x," & Text.RemoveFromStartIfPresent(Formula, EQUAL_SIGN) & "))"
    End With
    
    GenerateUntileFormula = Formula
    
End Function

Private Function IsValidForUntile(ByVal ValidMultiCellsRef As Collection) As Boolean
    
    Dim IsValid As Boolean
    If ValidMultiCellsRef.Count <> 1 Then
        IsValid = False
    Else
        Dim CurrentPrecedent As PrecedencyInfo
        Set CurrentPrecedent = ValidMultiCellsRef.Item(1)
        
        If Not CurrentPrecedent.HasSpill Then
            IsValid = False
        ElseIf Not IsSpillParent(CurrentPrecedent.NameInFormulaRange.Cells(1)) Then
            IsValid = False
        ElseIf CurrentPrecedent.NameInFormulaRange.Rows.CountLarge > 1 Or CurrentPrecedent.NameInFormulaRange.Columns.CountLarge > 1 Then
            IsValid = True
        End If
        
    End If
    
    IsValidForUntile = IsValid
    
End Function

Public Sub MapToArray_Undo()
    
    If MapToArrayUndoColl Is Nothing Then Exit Sub
    If MapToArrayUndoColl.Count = 0 Then Exit Sub
    Dim Item As UndoHandler
    Set Item = MapToArrayUndoColl.Item(1)
    
    Dim PlaceOnCell As Range
    Set PlaceOnCell = Item.ClearRange
    Dim OldFormula As String
    OldFormula = Item.FirstCellOldFormula
    If Not PlaceOnCell Is Nothing Then
        PlaceOnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    End If
    Set MapToArrayUndoColl = Nothing
    
End Sub

Private Function GetFormulaForMapToArray(ByVal ValidSpillParentCellsForMap As Collection _
                                         , ByVal FormulaCell As Range _
                                         , ByVal StructuredFormula As String _
                                         , MaxValidCellsRowCount As Long _
                                         , MaxValidCellsColCount As Long) As String
    
    ' We need to use TILE function
    Dim Formula As String
    If FormulaCell.HasSpill Then
        Formula = GetMapToArrayForTile(StructuredFormula, ValidSpillParentCellsForMap)
        Formula = ValidateGeneratedTileFormula(FormulaCell, Formula)
    Else
        
        Dim KeyToAnswer As Variant
        KeyToAnswer = GetMapToArrayKeyToAnswerByFillingFormula(FormulaCell, MaxValidCellsRowCount, MaxValidCellsColCount)
        Formula = GenerateFormulaByReplacingRef(FormulaCell.Formula2, ValidSpillParentCellsForMap)
        
        Dim FormulaResult As Variant
        FormulaResult = GetFormulaResult(Formula, FormulaCell)
        
        
        If Not IsBothSame(KeyToAnswer, FormulaResult) Then
            Formula = modUtility.GenerateFormulaForMapToArrayExceptTile(ValidSpillParentCellsForMap _
                                                                        , MAP_FN_NAME, StructuredFormula)
        End If
    End If
    
    GetFormulaForMapToArray = Formula
    
End Function

Private Function ValidateGeneratedTileFormula(ByVal FormulaCell As Range, ByVal Formula As String) As String
    
    Dim CurrentResult As Variant
    CurrentResult = FormulaCell.Cells(1).SpillingToRange.Value
    
    Dim RowCount As Long
    RowCount = FormulaCell.Cells(1).SpillingToRange.Rows.CountLarge
    
    Dim ColCount As Long
    ColCount = FormulaCell.Cells(1).SpillingToRange.Columns.CountLarge
    
    On Error GoTo HandleError
    Dim OldFormula As String
    OldFormula = FormulaCell.Cells(1).Formula2
    FormulaCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(Formula)
    
    Dim GeneratedResult As Variant
    GeneratedResult = FormulaCell.Resize(RowCount, ColCount).Value
    
    Dim CorrectFormula As String
    If IsBothSame(CurrentResult, GeneratedResult) Then
        CorrectFormula = Formula
    Else
        CorrectFormula = vbNullString
    End If
    ValidateGeneratedTileFormula = CorrectFormula
    
CleanExit:
    FormulaCell.Formula2 = OldFormula
    Exit Function
    
HandleError:
    ValidateGeneratedTileFormula = vbNullString
    GoTo CleanExit
    
End Function

Private Function GetMapToArrayForTile(ByVal StartFormula As String _
                                      , ByVal ValidCells As Collection) As String
    
    Dim Formula As String
    Formula = "=" & TILE_FN_NAME & "("
    Dim FirstItem As PrecedencyInfo
    Set FirstItem = ValidCells.Item(1)
    Dim FirstCellAddress As String
    FirstCellAddress = FirstItem.AbsRangeRef
    
    If ValidCells.Count = 1 Then
        ' If only one then use x as param name and replace that cell ref with x
        Formula = Formula & FirstItem.AbsChoosePartFormula & LIST_SEPARATOR _
                  & LAMBDA_AND_OPEN_PAREN & "x" & LIST_SEPARATOR
        Formula = Formula & RemoveStartingEqualSign(ReplaceTokenWithNewToken(StartFormula, FirstItem.NameInFormula, "x")) _
                  & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    Else
        ' if more than one then use n as lambda param name.
        Formula = Formula & SEQUENCE_FN_NAME & FIRST_PARENTHESIS_OPEN _
                  & ROWS_FN_NAME & FIRST_PARENTHESIS_OPEN & FirstCellAddress _
                  & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & COLUMNS_FN_NAME _
                  & FIRST_PARENTHESIS_OPEN & FirstCellAddress & FIRST_PARENTHESIS_CLOSE _
                  & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & LAMBDA_AND_OPEN_PAREN _
                  & "n" & LIST_SEPARATOR
                  
        Formula = CreateLambdaPartForTile(StartFormula, Formula, ValidCells)
    End If
    
    GetMapToArrayForTile = Formula
    
End Function

Private Function CreateLambdaPartForTile(ByVal BaseFormula As String _
                                         , ByVal TileFirstPart As String _
                                          , ByVal ValidCells As Collection) As String
    
    
    Dim StepNamePrefix As String
    StepNamePrefix = GetStepNamePrefix(BaseFormula, ValidCells.Count)
    Dim LetPart As String
    LetPart = LET_AND_OPEN_PAREN
    Dim CurrentItemIndex As Long
    Dim CurrentItem As PrecedencyInfo
    For CurrentItemIndex = 1 To ValidCells.Count
        Set CurrentItem = ValidCells.Item(CurrentItemIndex)
        Dim ParamName As String
        ParamName = GetParamNameFromCounter(StepNamePrefix, CurrentItemIndex, ValidCells.Count)
        
        Dim ParentCellRef As String
        ParentCellRef = CurrentItem.AbsRangeRef
        LetPart = LetPart & ParamName & LIST_SEPARATOR & ONE_SPACE _
                  & INDEX_FN_NAME & FIRST_PARENTHESIS_OPEN & TOROW_FN_NAME _
                  & FIRST_PARENTHESIS_OPEN & ParentCellRef & FIRST_PARENTHESIS_CLOSE _
                  & LIST_SEPARATOR & "1" & LIST_SEPARATOR _
                  & "n" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
                  
        BaseFormula = ReplaceTokenWithNewToken(BaseFormula, CurrentItem.NameInFormula, ParamName)
    Next CurrentItemIndex
    
    ' Three close paren because One is for LET, Another one is for Lambda and Third one is for TILE
    CreateLambdaPartForTile = TileFirstPart & LetPart & RemoveStartingEqualSign(BaseFormula) _
                              & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    
End Function

Private Function GetMapToArrayKeyToAnswerByFillingFormula(ByVal FormulaCell As Range _
                                                          , MaxValidCellsRowCount As Long _
                                                           , MaxValidCellsColCount As Long) As Variant
    
    Dim OldFormula As String
    OldFormula = FormulaCell.Cells(1).Formula2
    On Error GoTo ResetFormula
    Dim FilledFormulaRange As Range
    Set FilledFormulaRange = FormulaCell.Resize(MaxValidCellsRowCount, MaxValidCellsColCount)
    
    FilledFormulaRange.Formula2 = ReplaceInvalidCharFromFormulaWithValid(ConvertSpillRangeDependencyToAbsRef(FormulaCell.Cells(1), True))
    
    Dim DataRange As Range
    Set DataRange = FilledFormulaRange

    Dim Result As Variant
    Result = DataRange.Value
    GetMapToArrayKeyToAnswerByFillingFormula = Result
    
    FilledFormulaRange.Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Exit Function
    
ResetFormula:
    
    FilledFormulaRange.Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function




