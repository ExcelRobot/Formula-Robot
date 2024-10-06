Attribute VB_Name = "modMapToArray"
Option Explicit
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
                    modUtility.UpdateValidCells ValidSpillParentCellsForMap, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE
                End If
            End If
        End If
    Next CurrentPrecedency
    
    If ValidSpillParentCellsForMap.Count = 0 Then
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
        FullFormula = GetFormulaForMapToArray(ValidSpillParentCellsForMap, FormulaCell)
    Else
        Exit Sub
    End If
    
    If IsTileFormula(FullFormula) Then AddTILEIfNotPresent FormulaCell.Worksheet.Parent
    MapToArrayUndoColl.Add UndoHandler.Create(DYNAMIC_ARRAY_VERSION, PlaceFormulaToCell, PlaceFormulaToCell.Formula2)
    PlaceFormulaToCell.Formula2 = FullFormula
    AssingOnUndo "MapToArray"
    
End Sub

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
        PlaceOnCell.Formula2 = OldFormula
    End If
    Set MapToArrayUndoColl = Nothing
    
End Sub

Private Function GetFormulaForMapToArray(ByVal ValidSpillParentCellsForMap As Collection _
                                         , ByVal FormulaCell As Range) As String
    
    ' We need to use TILE function
    If FormulaCell.HasSpill Then
        GetFormulaForMapToArray = GetMapToArrayForTile(FormulaCell.Formula2 _
                                                       , ValidSpillParentCellsForMap)
    Else
        Dim Formula As String
        Formula = GenerateFormulaByReplacingRef(FormulaCell.Formula2, ValidSpillParentCellsForMap)
        Dim ResultArrayRowCount As Long
        ResultArrayRowCount = MaxRowCount(ValidSpillParentCellsForMap)
        If Not IsSpillRowCountSame(Formula, FormulaCell, ResultArrayRowCount) Then
            Formula = modUtility.GenerateFormulaForMapToArrayExceptTile(ValidSpillParentCellsForMap _
                                                                        , FormulaCell, MAP_FX_NAME)
        End If
        
        GetFormulaForMapToArray = Formula
    End If
    
End Function

Private Function GetMapToArrayForTile(ByVal StartFormula As String _
                                      , ByVal ValidCells As Collection) As String
    
    Dim Formula As String
    Formula = "=" & TILE_FX_NAME & "("
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
        Formula = Formula & SEQUENCE_FX_NAME & FIRST_PARENTHESIS_OPEN _
                  & ROWS_FX_NAME & FIRST_PARENTHESIS_OPEN & FirstCellAddress _
                  & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & COLUMNS_FX_NAME _
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
                  & INDEX_FX_NAME & FIRST_PARENTHESIS_OPEN & TOROW_FX_NAME _
                  & FIRST_PARENTHESIS_OPEN & ParentCellRef & FIRST_PARENTHESIS_CLOSE _
                  & LIST_SEPARATOR & "1" & LIST_SEPARATOR _
                  & "n" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
                  
        BaseFormula = ReplaceTokenWithNewToken(BaseFormula, CurrentItem.NameInFormula, ParamName)
    Next CurrentItemIndex
    
    ' Three close paren because One is for LET, Another one is for Lambda and Third one is for TILE
    CreateLambdaPartForTile = TileFirstPart & LetPart & RemoveStartingEqualSign(BaseFormula) _
                              & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    
End Function
