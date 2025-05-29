Attribute VB_Name = "modApplyFilterToArray"
Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Apply Filter To Array
' Description:            Create Filter formula based on top row spill range formula.
' Macro Expression:       modApplyFilterToArray.ApplyFilterToArray([ActiveCell],[ActiveCell.Offset(1,1)])
' Generated:              10/24/2023 09:42 AM
'----------------------------------------------------------------------------------------------------
Public Sub ApplyFilterToArray(ByVal FormulaCell As Range, Optional ByVal PlaceFormulaToCell As Range)
    
    Const COMMAND_NAME As String = "Apply Filter To Array Command"
    
    If Not FormulaCell.HasFormula Then
        MsgBox "No formula found in cell: " & FormulaCell.Address, vbCritical + vbInformation, COMMAND_NAME
        Exit Sub
    End If
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = GetDirectPrecedents(FormulaCell.Formula2, FormulaCell.Worksheet)
    
    If Not IsArray(DirectPrecedents) Then
        MsgBox "No direct precedent cell has been found in the formula.", vbCritical + vbInformation, COMMAND_NAME
        Exit Sub
    End If
    
    Dim ValidCellsForMAP As New Collection
    Dim ValidCellsForFilter As New Collection
    
    Dim CurrentPrecedency As Variant
    Dim CurrentRange As Range
    For Each CurrentPrecedency In DirectPrecedents
        
        Set CurrentRange = RangeResolver.GetRangeForDependency(CStr(CurrentPrecedency), FormulaCell)
        
        Dim IsValidForMap As Boolean
        IsValidForMap = False
        
        If Not IsNothing(CurrentRange) Then
            If CurrentRange.Cells(1).HasSpill Then
                If Not Intersect(CurrentRange, CurrentRange.Cells(1).SpillParent.SpillingToRange.Rows(1)) Is Nothing Then
                    IsValidForMap = True
                End If
            End If
        End If
        
        
        If IsValidForMap Then
                    
            If CurrentRange.Cells.Count = 1 Then
                UpdateValidCells ValidCellsForMAP, CStr(CurrentPrecedency), CurrentRange, FormulaCell, CHOOSE_COLS
            End If
                    
            Dim SpillParentRef As String
            SpillParentRef = GetParentCellRef(FormulaCell, CurrentRange, False)
                    
            If Not IsExistInCollection(ValidCellsForFilter, SpillParentRef) Then
                Dim CurrentPrecedencyInfo As PrecedencyInfo
                Set CurrentPrecedencyInfo = New PrecedencyInfo
                CurrentPrecedencyInfo.TopLeftCellColNo = CurrentRange.Column
                CurrentPrecedencyInfo.RangeRef = SpillParentRef
                ValidCellsForFilter.Add CurrentPrecedencyInfo, SpillParentRef
            End If
                    
        End If
        
    Next CurrentPrecedency
    
    If ValidCellsForFilter.Count = 0 Then
        MsgBox "No Valid cell has been found to do filter operation." _
               , vbCritical + vbInformation, "Apply Filter ToArray Command"
        Exit Sub
    End If
    
    ' If less than or equal three then variable name will be x,y,z order otherwise a,b,c...z
    Dim MaskPartFormula As String
    MaskPartFormula = GenerateFilterFormula(ValidCellsForMAP, FormulaCell.Formula2)
    If PlaceFormulaToCell Is Nothing Then Set PlaceFormulaToCell = FormulaCell
    
    Dim FullFormula As String
    FullFormula = EQUAL_SIGN & FILTER_FN_NAME & FIRST_PARENTHESIS_OPEN _
                  & GetFilterFirstParam(ValidCellsForFilter) & LIST_SEPARATOR _
                  & ONE_SPACE & RemoveStartingEqualSign(MaskPartFormula) & FIRST_PARENTHESIS_CLOSE
                  
    PlaceFormulaToCell.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(FullFormula)
    
End Sub

Private Function GetFilterFirstParam(ByVal ValidCellsForFilter As Collection) As String
    
    Dim CurrentRef As PrecedencyInfo
    Dim Formula As String
    If ValidCellsForFilter.Count = 1 Then
        Set CurrentRef = ValidCellsForFilter.Item(1)
        Formula = CurrentRef.RangeRef
    Else
        
        Dim Result As Variant
        ReDim Result(1 To ValidCellsForFilter.Count, 1 To 2) As Variant
        
        Formula = HSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
        Dim Counter As Long
        For Each CurrentRef In ValidCellsForFilter
            Counter = Counter + 1
            Result(Counter, 1) = CurrentRef.TopLeftCellColNo
            Result(Counter, 2) = CurrentRef.RangeRef
        Next CurrentRef
        Result = Application.WorksheetFunction.Sort(Result, 1, 1)
        
        Dim FirstColumnIndex  As Long
        FirstColumnIndex = LBound(Result, 2)
        Dim CurrentRowIndex As Long
        For CurrentRowIndex = LBound(Result, 1) To UBound(Result, 1)
            Formula = Formula & Result(CurrentRowIndex, FirstColumnIndex + 1) & LIST_SEPARATOR
        Next CurrentRowIndex
        
        Formula = Left$(Formula, Len(Formula) - Len(LIST_SEPARATOR)) & FIRST_PARENTHESIS_CLOSE
        
    End If
    
    GetFilterFirstParam = Formula
    
End Function

Private Function GenerateFilterFormula(ByVal ValidCellForParam As Collection _
                                       , ByVal StartFormula As String) As String
                                                     
    Dim CalculationPart As String
    CalculationPart = StartFormula
    Dim LambdaParamPart As String
    
    LambdaParamPart = LAMBDA_AND_OPEN_PAREN
    Dim MapParamPart As String
    MapParamPart = EQUAL_SIGN & MAP_FN_NAME & FIRST_PARENTHESIS_OPEN
    
    Dim Result As Variant
    Result = SortValidCellsByColNumber(ValidCellForParam)
    
    Dim CurrentRowIndex As Long
    For CurrentRowIndex = LBound(Result, 1) To UBound(Result, 1)
        
        Dim ParamName As String
        ParamName = GetParamNameFromCounter("x", CurrentRowIndex, ValidCellForParam.Count)
        
        LambdaParamPart = LambdaParamPart & ParamName & LIST_SEPARATOR
        
        MapParamPart = MapParamPart & Result(CurrentRowIndex, 3) & LIST_SEPARATOR
        CalculationPart = ReplaceTokenWithNewToken(CalculationPart _
                                                   , CStr(Result(CurrentRowIndex, 1)), ParamName)
    Next CurrentRowIndex
    
    GenerateFilterFormula = MapParamPart & LambdaParamPart _
                            & RemoveStartingEqualSign(CalculationPart) _
                            & FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    
End Function

Private Function SortValidCellsByColNumber(ByVal ValidCellForParam As Collection) As Variant
    
    Dim CurrentPrecedencyRef As PrecedencyInfo
    
    Dim Result As Variant
    ReDim Result(1 To ValidCellForParam.Count, 1 To 4) As Variant
    Dim Counter As Long
    For Each CurrentPrecedencyRef In ValidCellForParam
        
        Counter = Counter + 1
        Result(Counter, 1) = CurrentPrecedencyRef.NameInFormula
        Result(Counter, 2) = CurrentPrecedencyRef.RangeRef
        Result(Counter, 3) = CurrentPrecedencyRef.ChoosePartFormula
        Result(Counter, 4) = CurrentPrecedencyRef.TopLeftCellColNo
        
    Next CurrentPrecedencyRef
    If ValidCellForParam.Count > 1 Then
        Result = Application.WorksheetFunction.Sort(Result, 4, 1)
    End If
    SortValidCellsByColNumber = Result
    
End Function

