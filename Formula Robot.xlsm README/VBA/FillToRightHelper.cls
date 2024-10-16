VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FillToRightHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TFillToRightHelper
    FillToRightUndoColl As Collection
    IsInvalidForOpt As Boolean
    InvalidForCell As Range
End Type

Private This As TFillToRightHelper

Public Property Get FillToRightUndoColl() As Collection
    Set FillToRightUndoColl = This.FillToRightUndoColl
End Property

Public Property Set FillToRightUndoColl(ByVal RHS As Collection)
    Set This.FillToRightUndoColl = RHS
End Property

Public Sub FillToRight(ByVal FormulaCell As Range)
    
    Dim FillToRightBeans As Collection
    Set FillToRightBeans = New Collection
    
    Dim FillWidth As Long
    Dim Bean As FillBean
    If FormulaCell.Rows.Count > 1 And FormulaCell.Columns.Count = 1 Then
    
        Dim Counter As Long
        For Counter = 1 To FormulaCell.Rows.Count
            
            Set Bean = GetFillBean(FormulaCell.Rows(Counter))
            If IsNotNothing(Bean) Then
                FillToRightBeans.Add Bean
                
                If FillToRightBeans.Count = 1 Then
                    FillWidth = Bean.ProbableFillDimension
                ElseIf FillWidth <> Bean.ProbableFillDimension Then
                    This.IsInvalidForOpt = True
                End If
                If Not This.IsInvalidForOpt Then Bean.DoFill
            End If
            
            If This.IsInvalidForOpt Then
                Set This.InvalidForCell = FormulaCell.Rows(Counter)
                Exit For
            End If
            
        Next Counter
        
    Else
        
        Set Bean = GetFillBean(FormulaCell)
        If This.IsInvalidForOpt Then Set This.InvalidForCell = FormulaCell
        If IsNotNothing(Bean) And Not This.IsInvalidForOpt Then
            FillToRightBeans.Add Bean
            Bean.DoFill
        End If
        
    End If
    
    If This.IsInvalidForOpt Then
        This.InvalidForCell.Activate
        Exit Sub
    End If
    
    If FillToRightBeans.Count = 0 Then Exit Sub

    Set This.FillToRightUndoColl = GetUndoCells(FillToRightBeans)
    If This.FillToRightUndoColl.Count > 0 Then AssingOnUndo "FillToRight"
    
End Sub

Private Function GetFillBean(ByVal FormulaCell As Range) As FillBean
    
    Dim Bean As FillBean
    If FormulaCell.Cells.Count = 1 Then
        If FormulaCell.HasFormula Then
            
            ' If more than one col spill range then put same formula and exit
            If FormulaCell.HasSpill Then
                If SpillRangeColCount(FormulaCell) > 1 Then Exit Function
            End If
            
            Set Bean = FillBeanForDynamicArrayFormula(FormulaCell)
            If This.IsInvalidForOpt Then Exit Function
            If IsNothing(Bean) Then
                Set Bean = FillBeanForVBAFillToRight(FormulaCell)
            End If
            
        ElseIf FormulaCell.HasSpill Then
            Exit Function
        ElseIf Not IsBlankCellNoError(FormulaCell) Then
            Set Bean = FillBeanForVBAFillToRight(FormulaCell)
        End If
        
    ElseIf FormulaCell.Columns.Count > 1 Then
        
        Set Bean = New FillBean
        With Bean
            Set .OnRange = FormulaCell
            .ProbableFillDimension = FormulaCell.Columns.Count
            .TypeOfFill = FILL_TO_RIGHT
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
    
    Dim ValidCellsForFillRight As New Collection
    Dim ColWidth As Long
    Dim IsAtleastOnePrecedentSpilledRange As Boolean
    Dim CurrentPrecedency As Variant
    Dim CurrentRange As Range
    For Each CurrentPrecedency In DirectPrecedents
        Dim PrecedentCellAsText As String
        PrecedentCellAsText = CStr(CurrentPrecedency)
        Set CurrentRange = RangeResolver.GetRangeForDependency(PrecedentCellAsText, FormulaCell)
        
        Dim IsValidCell As Boolean
        IsValidCell = IsValidForArrayVersion(FormulaCell, PrecedentCellAsText, FILL_TO_RIGHT)
        
        If IsValidCell And Not CurrentRange.Cells(1).HasSpill Then
            
            If CurrentRange.Rows.Count = 1 Then
                modUtility.UpdateValidCells ValidCellsForFillRight, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE
            ElseIf CurrentRange.Rows.Count > 1 Then
                modUtility.UpdateValidCells ValidCellsForFillRight, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_ROWS
            End If
            
        ElseIf IsValidCell And CurrentRange.Cells(1).HasSpill Then
            
            IsAtleastOnePrecedentSpilledRange = True
            ColWidth = SpillRangeColCount(CurrentRange)
            If IsSpillParent(CurrentRange) And IsOneRowSpillRange(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillRight, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_NONE
            ElseIf IsInFirstColOfSpillRange(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillRight, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_ROWS
            ElseIf IsValidRangeUsedForByCol(CurrentRange) Then
                modUtility.UpdateValidCells ValidCellsForFillRight, PrecedentCellAsText, CurrentRange, FormulaCell, CHOOSE_ROWS
            End If
            
        End If
    Next CurrentPrecedency
        
    Dim Formula As String
    
    If IsAtleastOnePrecedentSpilledRange Then
       ChangeNonArraysDataRangeToSameSize ValidCellsForFillRight, ColWidth, FormulaCell, FILL_TO_RIGHT
    Else
        Exit Function
    End If
    
    If ValidCellsForFillRight.Count = 0 Then Exit Function
    UpdateValidity ValidCellsForFillRight
    If This.IsInvalidForOpt Then Exit Function
    If Not IsAllBlankAndNoFormulas(DropFirstCell(FormulaCell.Resize(, ColWidth))) Then Exit Function
    
    Formula = GenerateFormulaForFillToRight(FormulaCell, ValidCellsForFillRight)
    If Formula <> vbNullString Then
        Dim UndoOperator As UndoHandler
        Set UndoOperator = UndoHandler.Create(DYNAMIC_ARRAY_VERSION, FormulaCell, FormulaCell.Cells(1).Formula2)
        Set FillBeanForDynamicArrayFormula = CreateDAFillerBean(FormulaCell, Formula, ColWidth, UndoOperator, FILL_TO_RIGHT)
    ElseIf Formula = vbNullString And ValidCellsForFillRight.Count > 0 Then
        Set This.InvalidForCell = FormulaCell
        This.IsInvalidForOpt = True
    End If
    
End Function

Private Function GenerateFormulaForFillToRight(ByVal FormulaCell As Range, ByVal ValidCellsForFillRight As Collection) As String
    
    Dim ColCount As Long
    Dim CurrentPrecedencyInfo As PrecedencyInfo
    Set CurrentPrecedencyInfo = ValidCellsForFillRight.Item(1)
    ColCount = CurrentPrecedencyInfo.ColCount
    
    Dim KeyToAnswer As Variant
    KeyToAnswer = GetKeyToAnswerByFillingFormula(FormulaCell, ColCount)
    
    Dim FormulaWithoutFX As String
    FormulaWithoutFX = GenerateFormulaByReplacingRef(FormulaCell.Formula2, ValidCellsForFillRight)
    
    Dim FormulaWithFX As String
    FormulaWithFX = GenerateFormulaForFillToRightIfMapByRowOrTile(FormulaCell, ValidCellsForFillRight)
    Logger.Log DEBUG_LOG, "With Function: " & FormulaWithFX
    Logger.Log DEBUG_LOG, "Without Function: " & FormulaWithoutFX
    
    If IsTileFormula(FormulaWithFX) Then
        AddTILEIfNotPresent FormulaCell.Worksheet.Parent
    End If
    
    GenerateFormulaForFillToRight = GetCorrectFormula(FormulaCell, FormulaWithoutFX, FormulaWithFX, KeyToAnswer)
    
End Function

Private Function GetKeyToAnswerByFillingFormula(FormulaCell As Range, ColCount As Long) As Variant
    
    Dim OldFormula As String
    OldFormula = FormulaCell.Cells(1).Formula2
    On Error GoTo ResetFormula
    FormulaCell.Resize(, ColCount).Formula2 = ReplaceInvalidCharFromFormulaWithValid(ConvertSpillRangeDependencyToAbsRef(FormulaCell.Cells(1)))
    
    Dim MaxRowCount As Long
    MaxRowCount = GetMaxSpillRowCount(FormulaCell.Resize(, ColCount))
    
    Dim DataRange As Range
    Set DataRange = FormulaCell.Resize(MaxRowCount, ColCount)
    
    Dim Result As Variant
    Result = DataRange.Value
    
    If DataRange.Rows.Count > 1 Then
        
        Dim ColIndex As Long
        For ColIndex = LBound(Result, 2) To UBound(Result, 2)
            
            Dim SpillCellsCount As Long
            Dim FirstCell As Range
            Set FirstCell = DataRange.Cells(1, ColIndex - LBound(Result, 2) + 1)
            If FirstCell.HasSpill Then
                SpillCellsCount = FirstCell.SpillingToRange.Cells.Count
            Else
                SpillCellsCount = 1
            End If
            
            Dim RowIndex As Long
            For RowIndex = SpillCellsCount + 1 To DataRange.Rows.Count
                Result(RowIndex - LBound(Result, 1) + 1, ColIndex) = CVErr(xlErrNA)
            Next RowIndex
            
        Next ColIndex
        
    End If
    
    GetKeyToAnswerByFillingFormula = Result
    
    FormulaCell.Resize(, ColCount).Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Exit Function
    
ResetFormula:
    FormulaCell.Resize(, ColCount).Formula2 = vbNullString
    FormulaCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Function GetMaxSpillRowCount(ByVal FormulaCells As Range) As Long
    
    
    Dim MaxRowCount As Long
    MaxRowCount = 1
    
    Dim CurrentCell As Range
    For Each CurrentCell In FormulaCells.Cells
        If CurrentCell.HasSpill Then
            MaxRowCount = CurrentCell.SpillingToRange.Rows.Count
        End If
    Next CurrentCell
    
    GetMaxSpillRowCount = MaxRowCount
    
End Function

Private Function GenerateFormulaForFillToRightIfMapByRowOrTile(ByVal FormulaCell As Range _
                                                               , ByVal ValidCellsForFillRight As Collection) As String
    
    Dim Formula As String
    
    If FormulaCell.HasSpill Then
        Formula = modFillArray.GenerateFillWithSequence(FormulaCell.Formula2, ValidCellsForFillRight, FILL_TO_RIGHT, TILE_FX_NAME)
    Else
        
        If IsValidToUseByRowOrCol(ValidCellsForFillRight, FILL_TO_RIGHT) Then
            Formula = GenerateFormulaIfByCol(FormulaCell, ValidCellsForFillRight)
        ElseIf modUtility.IsValidToUseMapWithoutSequence(ValidCellsForFillRight, FILL_TO_RIGHT) Then
            Formula = GenerateFormulaForMapToArrayExceptTile(ValidCellsForFillRight, FormulaCell, MAP_FX_NAME)
        Else
            Formula = modFillArray.GenerateFillWithSequence(FormulaCell.Formula2, ValidCellsForFillRight, FILL_TO_RIGHT, MAP_FX_NAME)
        End If
    
    End If
    
    GenerateFormulaForFillToRightIfMapByRowOrTile = Formula
    
End Function

Private Sub UpdateValidity(ByVal ValidCells As Collection)
    
    If Not IsAllPrecedencyAreSameSize(ValidCells, FILL_TO_RIGHT) Then
        This.IsInvalidForOpt = True
        Exit Sub
    End If
    
End Sub

Private Function FillBeanForVBAFillToRight(ByVal FormulaCell As Range) As FillBean
     
    Set FormulaCell = GetRightRangeFromCurrentReg(FormulaCell)
    
    Dim Bean As FillBean
    If FormulaCell.Cells.Count > 1 Then
        
        Dim ColCount As Long
        ColCount = FormulaCell.Columns.Count
        Dim UndoOperator As UndoHandler
        Set UndoOperator = UndoHandler.Create(VBA_VERSION, DropFirstCell(FormulaCell), vbNullString)
        If IsAllBlankAndNoFormulas(DropFirstCell(FormulaCell)) Then
            Set Bean = CreateVBAFillerBean(FormulaCell.Cells(1), ColCount, UndoOperator, FILL_TO_RIGHT)
        ElseIf ConsistentFormulaCount(DropFirstCell(FormulaCell.Resize(, FormulaCell.Worksheet.Columns.Count - FormulaCell.Column)), False) <> ColCount - 1 Then
            This.IsInvalidForOpt = True
            Exit Function
        Else
            UndoOperator.SecondCellFormulaR1C1 = FormulaCell.Cells(2).Formula2R1C1
            Set Bean = CreateVBAFillerBean(FormulaCell.Cells(1), ColCount, UndoOperator, FILL_TO_RIGHT)
        End If
    Else
        This.IsInvalidForOpt = True
    End If
    
    Set FillBeanForVBAFillToRight = Bean
    
End Function

Public Function GetRightRangeFromCurrentReg(ByVal FormulaCell As Range) As Range
    
    Dim TopLeftCell As Range
    Set TopLeftCell = FormulaCell.Cells(1)
    
    ' If no columns to fill to, we'll try to guess how far to fill
    If FormulaCell.Columns.Count = 1 Then
        Dim ReferenceArea As Range
        Dim LastColumn As Range
        
        ' First try current region
        Set ReferenceArea = TopLeftCell.CurrentRegion
        Set LastColumn = ReferenceArea.Columns(ReferenceArea.Columns.Count).EntireColumn
        
        ' Then try current region of whatever is directly above
        If LastColumn.Column = TopLeftCell.Column Then
            Set ReferenceArea = TopLeftCell.End(xlUp).CurrentRegion
            Set LastColumn = ReferenceArea.Columns(ReferenceArea.Columns.Count).EntireColumn
        End If
        
        Set FormulaCell = Range(FormulaCell, Intersect(FormulaCell.EntireRow, LastColumn))
    End If
    Set GetRightRangeFromCurrentReg = FormulaCell
    
End Function


