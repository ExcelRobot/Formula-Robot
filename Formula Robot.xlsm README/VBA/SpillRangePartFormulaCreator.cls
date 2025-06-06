VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpillRangePartFormulaCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredID
Option Explicit

Private Type TSpillRangePartFormulaCreator
    SpillParentCell As Range
    SpillParentSheetName As String
    ParentCellFormulaRef As String
    SpillingRange As Range
    TopLeftCell As Range
    TopRightCell As Range
    BottomLeftCell As Range
    BottomRightCell As Range
    IsSpillingRangeIncluded As Boolean
    IsTopLeftCellIncluded As Boolean
    IsTopRightCellIncluded As Boolean
    IsBottomLeftCellIncluded As Boolean
    IsBottomRightCellIncluded As Boolean
    IsAnyCellOfFirstCol As Boolean
    IsAnyCellOfLastCol As Boolean
    IsAnyCellOfFirstRow As Boolean
    IsAnyCellOfLastRow As Boolean
    TopCellRowDiff As Long
    TopCellColDiff As Long
    RowCount As Long
    ColCount As Long
End Type

Private This As TSpillRangePartFormulaCreator

Public Function GetFormula(ByVal SelectedCells As Range, ByVal DestinationCell As Range) As String
    
    If Not SelectedCells.Cells(1).HasSpill Then Exit Function
    
    ExtractInfoFromSelectedCells SelectedCells, DestinationCell
    
    Dim Formula As String

    Dim TakeWithParentCellRef As String
    TakeWithParentCellRef = TAKE_FN_NAME & "(" & This.ParentCellFormulaRef
    
    Dim TakeAndDropWithParentCellRef As String
    TakeAndDropWithParentCellRef = TAKE_FN_NAME & "(" & DROP_FN_NAME & "(" & This.ParentCellFormulaRef
    
    With This
        ' Here order of these cases are important. If you want to make it independent then you have to change code
        ' and add validation. But as select case statement will go one by one, we can keep in this way.
        Select Case True
            Case .IsSpillingRangeIncluded
                Formula = This.ParentCellFormulaRef
            
            Case .IsTopLeftCellIncluded And .IsTopRightCellIncluded
                Formula = TakeWithParentCellRef & "," & .RowCount & ")"
        
            Case .IsBottomLeftCellIncluded And .IsBottomRightCellIncluded
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount & ")"
        
            Case .IsTopLeftCellIncluded And .IsBottomLeftCellIncluded
                Formula = TakeWithParentCellRef & ",," & .ColCount & ")"
            
            Case .IsTopRightCellIncluded And .IsBottomRightCellIncluded
                Formula = TakeWithParentCellRef & ",," & "-" & .ColCount & ")"
            
            Case .IsTopLeftCellIncluded
                Formula = TakeWithParentCellRef & "," & .RowCount _
                          & "," & .ColCount & ")"
                
            Case .IsTopRightCellIncluded
                Formula = TakeWithParentCellRef & "," & .RowCount _
                          & "," & "-" & .ColCount & ")"
                
            Case .IsBottomLeftCellIncluded
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount _
                          & "," & .ColCount & ")"
                
            Case .IsBottomRightCellIncluded
                Formula = TakeWithParentCellRef & "," & "-" & .RowCount _
                          & "," & "-" & .ColCount & ")"
            
            Case .IsAnyCellOfFirstCol And .IsAnyCellOfLastCol
                If .RowCount = 1 Then
                    Formula = "INDEX(" & .ParentCellFormulaRef & "," & .TopCellRowDiff + 1 & ",)"
                Else
                    Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & ")" _
                              & "," & .RowCount & ")"
                End If
                
            Case .IsAnyCellOfFirstRow And .IsAnyCellOfLastRow
                If .ColCount = 1 Then
                    Formula = "INDEX(" & .ParentCellFormulaRef & ",," & .TopCellColDiff + 1 & ")"
                Else
                    Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & ")" _
                              & ",," & .ColCount & ")"
                End If
                
            Case .IsAnyCellOfFirstCol
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & ")" _
                          & "," & .RowCount & "," & .ColCount & ")"
            
            Case .IsAnyCellOfFirstRow
                Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & ")" _
                          & "," & .RowCount & "," & .ColCount & ")"
                         
            Case .IsAnyCellOfLastCol
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & ")" _
                          & "," & .RowCount & "," & "-" & .ColCount & ")"
            
            Case .IsAnyCellOfLastRow
                Formula = TakeAndDropWithParentCellRef & ",," & .TopCellColDiff & ")" _
                          & "," & "-" & .RowCount & "," & .ColCount & ")"
            
                        
            Case .RowCount = 1 And .ColCount = 1
                Formula = "INDEX(" & .ParentCellFormulaRef & "," & .TopCellRowDiff + 1 & "," & .TopCellColDiff + 1 & ")"
            
            Case Else
                Formula = TakeAndDropWithParentCellRef & "," & .TopCellRowDiff & "," _
                          & .TopCellColDiff & ")" & "," & .RowCount & "," & .ColCount & ")"
                
        End Select
    End With
    
    GetFormula = Formula
    
End Function

Private Sub ExtractInfoFromSelectedCells(ByVal SelectedCells As Range, ByVal DestinationCell As Range)
    
    With This
        
        Set .SpillParentCell = SelectedCells.Cells(1).SpillParent
        .SpillParentSheetName = .SpillParentCell.Worksheet.Name
        
        ' We are considering that workbook is same. If need to support for different workbook as well then need
        ' to change code accordingly.
        .ParentCellFormulaRef = vbNullString
        If DestinationCell.Worksheet.Name <> .SpillParentSheetName Then
            .ParentCellFormulaRef = GetSheetRefForRangeReference(.SpillParentSheetName)
        End If
        
        .ParentCellFormulaRef = .ParentCellFormulaRef _
                              & .SpillParentCell.Address(False, False) & "#"
        
        Set .TopLeftCell = .SpillParentCell
        .IsTopLeftCellIncluded = IsSubRange(SelectedCells, .TopLeftCell)
        
        Set .SpillingRange = .SpillParentCell.SpillingToRange
        .IsSpillingRangeIncluded = (.SpillingRange.Address = SelectedCells.Address)
        Set .TopRightCell = .SpillingRange.Rows(1).Cells(.SpillingRange.Columns.CountLarge)
        .IsTopRightCellIncluded = IsSubRange(SelectedCells, .TopRightCell)
        
        Set .BottomLeftCell = .SpillingRange.Rows(.SpillingRange.Rows.Count).Cells(1)
        .IsBottomLeftCellIncluded = IsSubRange(SelectedCells, .BottomLeftCell)
        
        Set .BottomRightCell = .BottomLeftCell.Offset(, .SpillingRange.Columns.Count - 1)
        .IsBottomRightCellIncluded = IsSubRange(SelectedCells, .BottomRightCell)
        
        .IsAnyCellOfFirstCol = IsAnyCommonCell(SelectedCells, .SpillingRange.Columns(1))
        .IsAnyCellOfFirstRow = IsAnyCommonCell(SelectedCells, .SpillingRange.Rows(1))
        .IsAnyCellOfLastCol = IsAnyCommonCell(SelectedCells, .SpillingRange.Columns(.SpillingRange.Columns.Count))
        .IsAnyCellOfLastRow = IsAnyCommonCell(SelectedCells, .SpillingRange.Rows(.SpillingRange.Rows.Count))
        
        This.RowCount = SelectedCells.Rows.Count
        This.ColCount = SelectedCells.Columns.Count
        This.TopCellRowDiff = SelectedCells.Row - .SpillParentCell.Row
        This.TopCellColDiff = SelectedCells.Column - .SpillParentCell.Column
        
    End With
    
End Sub

Private Function IsAnyCommonCell(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    
    Dim IntersectingCells As Range
    Set IntersectingCells = Intersect(FirstRange, SecondRange)
    
    IsAnyCommonCell = (Not IntersectingCells Is Nothing)
    
End Function


