VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FillDownHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type TFillDownHelper
    FillDownUndoColl As Collection
    IsInvalidForOpt As Boolean
    InvalidForCell As Range
End Type

Private this As TFillDownHelper

Public Property Get FillDownUndoColl() As Collection
    Set FillDownUndoColl = this.FillDownUndoColl
End Property

Public Property Set FillDownUndoColl(ByVal RHS As Collection)
    Set this.FillDownUndoColl = RHS
End Property

Public Sub FillDown(ByVal FormulaCell As Range)

    Dim FillDownBeans As Collection
    Set FillDownBeans = New Collection
    Dim FillHeight As Long
    Dim Bean As FillBean
    If FormulaCell.Rows.Count = 1 And FormulaCell.Columns.Count > 1 Then
    
        Dim Counter As Long
        For Counter = 1 To FormulaCell.Columns.Count
            Set Bean = GetFillBean(FormulaCell.Columns(Counter))
        
            If IsNotNothing(Bean) Then
                FillDownBeans.Add Bean
                If FillDownBeans.Count = 1 Then
                    FillHeight = Bean.ProbableFillDimension
                ElseIf FillHeight <> Bean.ProbableFillDimension Then
                    this.IsInvalidForOpt = True
                End If
                If Not this.IsInvalidForOpt Then Bean.DoFill
            End If
            
            If this.IsInvalidForOpt Then
                Set this.InvalidForCell = FormulaCell.Columns(Counter)
                Exit For
            End If
            
        Next Counter
        
    Else
        
        Set Bean = GetFillBean(FormulaCell)
        If this.IsInvalidForOpt Then Set this.InvalidForCell = FormulaCell
        
        If IsNotNothing(Bean) And Not this.IsInvalidForOpt Then
            FillDownBeans.Add Bean
            Bean.DoFill
        End If
        
    End If
    
    If this.IsInvalidForOpt Then
        this.InvalidForCell.Activate
        Exit Sub
    End If

    If FillDownBeans.Count = 0 Then Exit Sub
    Set this.FillDownUndoColl = GetUndoCells(FillDownBeans)
    SelectFillingRegionIfNotDA FillDownBeans
    If this.FillDownUndoColl.Count > 0 Then AssingOnUndo "FillDown"
    
End Sub

Private Function GetFillBean(ByVal FormulaCell As Range) As FillBean
    
    Dim Bean As FillBean
    If FormulaCell.Cells.Count = 1 Then
        
        If FormulaCell.HasFormula Then
            
            ' If has spill formula and more than one row then do nothing.
            If FormulaCell.HasSpill Then
                If SpillRangeRowCount(FormulaCell) > 1 Then Exit Function
            End If
            
            Set Bean = FillBeanForDynamicArrayFormula(FormulaCell)
            If this.IsInvalidForOpt Then Exit Function
            If IsNothing(Bean) Then
                Set Bean = FillBeanForVBAFillDown(FormulaCell)
            End If
        ElseIf FormulaCell.HasSpill Then
            Exit Function
        ElseIf Not IsBlankCellNoError(FormulaCell) Then
            Set Bean = FillBeanForVBAFillDown(FormulaCell)
        End If
        
    ElseIf FormulaCell.Rows.Count > 1 Then
        
        Set Bean = New FillBean
        With Bean
            Set .OnRange = FormulaCell
            .ProbableFillDimension = FormulaCell.Rows.Count
            .TypeOfFill = Fill_DOWN
            Set .Filler = New NativeFiller
        End With
        
    End If
    Set GetFillBean = Bean
    
End Function

Private Function FillBeanForDynamicArrayFormula(ByVal FormulaCell As Range) As FillBean
    
    If Not FormulaCell.HasFormula Then Exit Function
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = GetDirectPrecedents(FormulaCell.Formula2, FormulaCell.Worksheet)
    
    If Not IsArray(DirectPrecedents) Then Exit Function
    If IsArrayOfNullString(DirectPrecedents) Then Exit Function
    
    Dim ValidCellsForFillDown As Collection
    Set ValidCellsForFillDown = New Collection
    
    Dim CurrentPrecedency As Variant
    Dim CurrentRange As Range
    Dim IsAtleastOnePrecedentSpilledRange As Boolean
    Dim RowHeight As Long
    For Each CurrentPrecedency In DirectPrecedents
        
        Dim PrecedentCellAsText As String
        PrecedentCellAsText = CStr(CurrentPrecedency)
        Set CurrentRange = RangeResolver.GetRangeForDependency(PrecedentCellAsText, FormulaCell)
        
        If CurrentRange Is Nothing Then GoTo Continue
        
        Dim IsValidCell As Boolean
        IsValidCell = IsValidForArrayVersion(FormulaCell, PrecedentCellAsText, Fill_DOWN)
        
        If Not IsValidCell Then GoTo Continue
        
        If CurrentRange.Cells(1).HasSpill Then
            
            IsAtleastOnePrecedentSpilledRange = True
            RowHeight = SpillRangeRowCount(CurrentRange)
            
            If IsSpillParent(CurrentRange) And IsOneColSpillRange(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillDown, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE
            ElseIf IsInFirstRowOfSpillRange(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillDown, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_COLS
            ElseIf IsValidRangeUsedForByRow(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillDown, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_COLS
            End If
            
        ElseIf Not CurrentRange.Cells(1).HasSpill Then
            
            If CurrentRange.Columns.Count = 1 Then
                modUtility.UpdateValidCells ValidCellsForFillDown, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE
            ElseIf CurrentRange.Columns.Count > 1 Then
                modUtility.UpdateValidCells ValidCellsForFillDown, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_COLS
            End If
            
        End If
        
Continue:
    Next CurrentPrecedency
    
    If IsAtleastOnePrecedentSpilledRange Then
        ChangeNonArraysDataRangeToSameSize ValidCellsForFillDown, RowHeight, FormulaCell, Fill_DOWN
    Else
        Exit Function
    End If
    
    If ValidCellsForFillDown.Count = 0 Then Exit Function
    
    UpdateValidity ValidCellsForFillDown
    
    If this.IsInvalidForOpt Then Exit Function
    
    If Not IsAllBlankAndNoFormulas(DropFirstCell(FormulaCell.Resize(RowHeight))) Then Exit Function
    
    Dim Formula As String
    Formula = GenerateFormulaForFillDown(FormulaCell, ValidCellsForFillDown)
    
    If Formula <> vbNullString Then
        Dim UndoOperator As UndoHandler
        Set UndoOperator = UndoHandler.Create(DYNAMIC_ARRAY_VERSION, FormulaCell, FormulaCell.Cells(1).Formula2)
        Set FillBeanForDynamicArrayFormula = CreateDAFillerBean(FormulaCell, Formula, RowHeight, UndoOperator, Fill_DOWN)
    ElseIf Formula = vbNullString And ValidCellsForFillDown.Count > 0 Then
        Set this.InvalidForCell = FormulaCell
        this.IsInvalidForOpt = True
    End If
    
End Function

Private Function GenerateFormulaForFillDown(ByVal FormulaCell As Range _
                                            , ByVal ValidCellsForFillDown As Collection) As String
    
    Dim RowCount As Long
    Dim CurrentPrecedencyInfo As PrecedencyInfo
    Set CurrentPrecedencyInfo = ValidCellsForFillDown.Item(1)
    RowCount = CurrentPrecedencyInfo.RowCount
    
    Dim KeyToAnswer As Variant
    KeyToAnswer = GetKeyToAnswerByFillingFormula(FormulaCell, RowCount)
    
    Dim FormulaWithoutFX As String
    FormulaWithoutFX = GenerateFormulaByReplacingRef(FormulaCell.Formula2, ValidCellsForFillDown)
    
    Dim FormulaWithFX As String
    FormulaWithFX = GenerateFormulaForFillDownIfMapByRowOrTile(FormulaCell, ValidCellsForFillDown)
    Logger.Log DEBUG_LOG, "With Function: " & FormulaWithFX
    Logger.Log DEBUG_LOG, "Without Function: " & FormulaWithoutFX
    
    AddCustomLambdaIfNeeded FormulaCell.Worksheet.Parent, FormulaWithFX
    Dim CorrectFormula As String
    CorrectFormula = GetCorrectFormula(FormulaCell, FormulaWithoutFX, FormulaWithFX, KeyToAnswer)
    
    If CorrectFormula = vbNullString Then
        CorrectFormula = GetFormulaByCheckingOutputType(FormulaCell, ValidCellsForFillDown, KeyToAnswer)
    End If
    
    GenerateFormulaForFillDown = CorrectFormula
    
End Function

Private Function GetFormulaByCheckingOutputType(ByVal FormulaCell As Range _
                                                , ByVal ValidCellsForFillDown As Collection _
                                                 , ByVal KeyToAnswer As Variant) As String
    
    Dim CorrectFormula As String
    
    Dim TypeFormula As String
    TypeFormula = EQUAL_TYPE_FIRST_PAREN & Text.RemoveFromStart(FormulaCell.Formula2, 1) & FIRST_PARENTHESIS_CLOSE
    
    Dim TypeOutput As Variant
    TypeOutput = GetFormulaResult(TypeFormula, FormulaCell)
    
    If IsNumeric(TypeOutput) Then
        If TypeOutput = TYPE_OUTPUT_FOR_ARR Then
            CorrectFormula = GenerateFormulaForFillDownIfArrayOutput(FormulaCell, ValidCellsForFillDown)
        End If
    End If
    
    AddCustomLambdaIfNeeded FormulaCell.Worksheet.Parent, CorrectFormula
    Dim FirstArray As Variant
    FirstArray = GetFormulaResult(CorrectFormula, FormulaCell)
    If Not IsBothSame(FirstArray, KeyToAnswer) Then
        CorrectFormula = vbNullString
    End If
    
    GetFormulaByCheckingOutputType = CorrectFormula
    
End Function

Private Function GetKeyToAnswerByFillingFormula(ByVal FormulaCell As Range, ByVal RowCount As Long) As Variant
    
    Dim OldFormula As String
    OldFormula = FormulaCell.Cells(1).Formula2
    On Error GoTo ResetFormula
    FormulaCell.Resize(RowCount).Formula2 = ReplaceInvalidCharFromFormulaWithValid(ConvertSpillRangeDependencyToAbsRef(FormulaCell.Cells(1)))
    
    Dim MaxColCount As Long
    MaxColCount = GetMaxSpillColCount(FormulaCell.Resize(RowCount))
    
    Dim DataRange As Range
    Set DataRange = FormulaCell.Resize(RowCount, MaxColCount)

    Dim Result As Variant
    Result = DataRange.Value
    If DataRange.Columns.Count > 1 Then
    
        Dim RowIndex As Long
        For RowIndex = LBound(Result, 1) To UBound(Result, 1)
            
            Dim SpillCellsCount As Long
            Dim FirstCell As Range
            Set FirstCell = DataRange.Cells(RowIndex - LBound(Result, 1) + 1, 1)
            If FirstCell.HasSpill Then
                SpillCellsCount = FirstCell.SpillingToRange.Cells.Count
            Else
                SpillCellsCount = 1
            End If
            
            Dim ColIndex As Long
            For ColIndex = SpillCellsCount + 1 To DataRange.Columns.Count
                Result(RowIndex, ColIndex - LBound(Result, 2) + 1) = CVErr(xlErrNA)
            Next ColIndex
            
        Next RowIndex
        
    End If
    
    GetKeyToAnswerByFillingFormula = Result
    
    FormulaCell.Resize(RowCount).Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Exit Function
    
ResetFormula:
    
    FormulaCell.Resize(RowCount).Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetMaxSpillColCount(ByVal FormulaCells As Range) As Long
    
    
    Dim MaxColCount As Long
    MaxColCount = 1
    
    Dim CurrentColCount As Long
    
    Dim CurrentCell As Range
    For Each CurrentCell In FormulaCells.Cells
        If CurrentCell.HasSpill Then
            CurrentColCount = CurrentCell.SpillingToRange.Columns.Count
            If CurrentColCount > MaxColCount Then MaxColCount = CurrentColCount
        End If
    Next CurrentCell
    
    GetMaxSpillColCount = MaxColCount
    
End Function

Private Function GenerateFormulaForFillDownIfMapByRowOrTile(ByVal FormulaCell As Range _
                                                            , ByVal ValidCellsForFillDown As Collection) As String
    
    Dim Formula As String
    
    If FormulaCell.HasSpill Then
        Formula = GenerateFormulaForFillDownIfArrayOutput(FormulaCell, ValidCellsForFillDown)
    Else
        
        If IsValidToUseByRowOrCol(ValidCellsForFillDown, Fill_DOWN) Then
            Formula = GenerateFormulaIfByRow(FormulaCell, ValidCellsForFillDown)
        ElseIf IsValidToUseMapWithoutSequence(ValidCellsForFillDown, Fill_DOWN) Then
            Formula = GenerateFormulaForMapToArrayExceptTile(ValidCellsForFillDown, FormulaCell, MAP_FN_NAME)
        Else
            Formula = modFillArray.GenerateFillWithSequence(FormulaCell.Formula2, ValidCellsForFillDown, Fill_DOWN, MAP_FN_NAME)
        End If
    
    End If
    
    GenerateFormulaForFillDownIfMapByRowOrTile = Formula
    
End Function

Private Function GenerateFormulaForFillDownIfArrayOutput(ByVal FormulaCell As Range _
                                                            , ByVal ValidCellsForFillDown As Collection)
    
    Dim Formula As String
    If IsValidToUseByRowOrCol(ValidCellsForFillDown, Fill_DOWN) Then
        Formula = GenerateFillWhenByRowAOrColA(FormulaCell.Formula2, ValidCellsForFillDown, True)
    Else
        Formula = modFillArray.GenerateFillWithSequence(FormulaCell.Formula2, ValidCellsForFillDown, Fill_DOWN, TILE_FN_NAME)
    End If
    
    GenerateFormulaForFillDownIfArrayOutput = Formula
    
End Function

Private Sub UpdateValidity(ByVal ValidCells As Collection)
    
    If Not IsAllPrecedencyAreSameSize(ValidCells, Fill_DOWN) Then
        this.IsInvalidForOpt = True
        Exit Sub
    End If
    
End Sub

Private Function FillBeanForVBAFillDown(ByVal FormulaCell As Range) As FillBean
    
    Set FormulaCell = GetBottomRangeFromCurrentReg(FormulaCell)
    
    Dim Bean As FillBean
    If FormulaCell.Cells.Count > 1 Then
        
        Dim RowCount As Long
        RowCount = FormulaCell.Rows.Count
        Dim UndoOperator As UndoHandler
        Set UndoOperator = UndoHandler.Create(VBA_VERSION, DropFirstCell(FormulaCell), vbNullString)
        If IsAllBlankAndNoFormulas(DropFirstCell(FormulaCell)) Then
            Set Bean = CreateVBAFillerBean(FormulaCell.Cells(1), RowCount, UndoOperator, Fill_DOWN)
        ElseIf ConsistentFormulaCount(DropFirstCell(FormulaCell.Resize(FormulaCell.Worksheet.Rows.Count - FormulaCell.Row)), True) <> RowCount - 1 Then
            this.IsInvalidForOpt = True
            Exit Function
        Else
            UndoOperator.SecondCellFormulaR1C1 = FormulaCell.Cells(2).Formula2R1C1
            Set Bean = CreateVBAFillerBean(FormulaCell.Cells(1), RowCount, UndoOperator, Fill_DOWN)
        End If
    Else
        this.IsInvalidForOpt = True
    End If
    Set FillBeanForVBAFillDown = Bean
            
End Function

Private Function GetBottomRangeFromCurrentReg(ByVal FormulaCell As Range) As Range
    
    Dim TopLeftCell As Range
    Set TopLeftCell = FormulaCell.Cells(1)
    
    ' If no rows to fill to, we'll try to guess how far to fill
    If FormulaCell.Rows.Count = 1 Then
        Dim ReferenceArea As Range
        Dim LastRow As Range
        
        ' First try current region
        Set ReferenceArea = TopLeftCell.CurrentRegion
        Set LastRow = ReferenceArea.Rows(ReferenceArea.Rows.Count).EntireRow
        
        ' Then try current region of whatever is directly to left
        If LastRow.Row = TopLeftCell.Row Then
            Set ReferenceArea = TopLeftCell.End(xlToLeft).CurrentRegion
            ' Need to add validation to check if more than one cell.
            ' If not
            Set LastRow = ReferenceArea.Rows(ReferenceArea.Rows.Count).EntireRow
        End If
        
        Set GetBottomRangeFromCurrentReg = FormulaCell.Worksheet.Range(FormulaCell, Intersect(FormulaCell.EntireColumn, LastRow))
    Else
        Set GetBottomRangeFromCurrentReg = FormulaCell
    End If
    
End Function


