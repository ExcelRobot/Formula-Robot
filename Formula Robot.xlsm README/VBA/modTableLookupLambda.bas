Attribute VB_Name = "modTableLookupLambda"
Option Explicit

Private Sub Test()
    GenerateTableLookupLambdas ActiveCell, Selection
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Generate Table Lookup Lambdas
' Description:            Generate lambdas for each column of an Excel Table.
' Macro Expression:       modTableLookupLambda.GenerateTableLookupLambdas([ActiveCell])
' Generated:              04/26/2024 04:11 PM
'----------------------------------------------------------------------------------------------------
Public Sub GenerateTableLookupLambdas(ByVal TableCell As Range, SelectedCells As Range)
    
    ' If not table nor spill range
    If TableCell.ListObject Is Nothing And Not TableCell.HasSpill Then
        Exit Sub
    End If
    
    If TableCell.ListObject Is Nothing Then
        If Not IsInsideNamedRange(TableCell.SpillParent.SpillingToRange) Then
            modNamedRange.AddNameRange TableCell.SpillParent.SpillingToRange, vbNullString
        End If
    End If
    
    Dim DataRangeWithHeader As Range
    Dim TableName As String
    Dim IsTable As Boolean
    
    If Not TableCell.ListObject Is Nothing Then
        IsTable = True
        TableName = TableCell.ListObject.Name
        Set DataRangeWithHeader = TableCell.ListObject.Range
    ElseIf IsInsideNamedRange(TableCell.SpillParent.SpillingToRange) Then
        Dim CurrentName As Name
        Set CurrentName = FindNamedRangeFromSubCell(TableCell.SpillParent.SpillingToRange)
        IsTable = False
        TableName = CurrentName.Name
        Set DataRangeWithHeader = CurrentName.RefersToRange
    Else
        Exit Sub
    End If
    
    If Not IsValidDefinedName(TableName) Then
        TableName = MakeValidDefinedName(TableName, False, True)
    End If
    
    Dim ValidHeaders As Range
    Set ValidHeaders = GetSelectedHeaders(SelectedCells, DataRangeWithHeader)
    
    If ValidHeaders Is Nothing Then Exit Sub
    
    Dim LambdaGenerator As TableLookupLambdaGenerator
    Set LambdaGenerator = New TableLookupLambdaGenerator
    LambdaGenerator.GenerateTemplateLambda IsTable, TableName, ValidHeaders
    
    Dim Book As Workbook
    Set Book = TableCell.Worksheet.Parent
    
    With Book
        .Names.Add TableName & "." & "Select", LambdaGenerator.FilterLambda
    End With
    
    Dim DefPart As String
    DefPart = LambdaGenerator.DefPart
    
    Dim FilterInvocationPart As String
    FilterInvocationPart = Text.RemoveFromEndIfPresent(Replace(Replace(DefPart, "[", vbNullString), "]", vbNullString), LIST_SEPARATOR)
    
    Dim ColIndex As Long
    For ColIndex = 1 To DataRangeWithHeader.Columns.Count
        
        Dim CurrentColParamName As String
        CurrentColParamName = DataRangeWithHeader.Cells(1, ColIndex).Value
        
        Dim Lambda As String
        
        Dim ReturnColDataPart As String
        ReturnColDataPart = DOUBLE_QUOTE & EscapeDoubleQuote(CurrentColParamName) & DOUBLE_QUOTE
        
        Lambda = EQUAL_SIGN & LAMBDA_AND_OPEN_PAREN & DefPart & ONE_SPACE & LET_AND_OPEN_PAREN & NEW_LINE & _
                 THREE_SPACE & "_FilteredDataWithHeader" & LIST_SEPARATOR & ONE_SPACE & TableName & ".Select" & FIRST_PARENTHESIS_OPEN _
                 & ReturnColDataPart & LIST_SEPARATOR & FilterInvocationPart & FIRST_PARENTHESIS_CLOSE _
                 & LIST_SEPARATOR & vbNewLine & _
                 THREE_SPACE & "_Result" & LIST_SEPARATOR & ONE_SPACE & IF_FX_NAME & FIRST_PARENTHESIS_OPEN & ROWS_FX_NAME _
                 & FIRST_PARENTHESIS_OPEN & "_FilteredDataWithHeader" & FIRST_PARENTHESIS_CLOSE & EQUAL_SIGN & "1" & LIST_SEPARATOR _
                 & NA_FX_NAME & FIRST_PARENTHESIS_OPEN & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & DROP_FX_NAME _
                 & FIRST_PARENTHESIS_OPEN & "_FilteredDataWithHeader" & LIST_SEPARATOR & "1" & FIRST_PARENTHESIS_CLOSE _
                 & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE & _
                 THREE_SPACE & "_Result" & NEW_LINE & _
                 FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
        
        CurrentColParamName = MakeValidDefinedName(CurrentColParamName, False, True)
        
        With Book
            If CurrentColParamName <> "Select" Then
                .Names.Add TableName & "." & CurrentColParamName, Replace(Lambda, vbNewLine, Chr(10))
            End If
        End With
        
    Next ColIndex
    
End Sub

Private Function GetSelectedHeaders(ByVal SelectedCells As Range, ByVal DataRangeWithHeader As Range) As Range
    
    Dim ValidHeaders As Range
    Set ValidHeaders = Intersect(SelectedCells.EntireColumn, DataRangeWithHeader.Rows(1))
    If ValidHeaders Is Nothing Then Exit Function
    
    Dim Temp As Variant
    ReDim Temp(1 To ValidHeaders.Cells.Count, 1 To 2)
    
    Dim Counter As Long
    Dim CurrentCell As Range
    For Each CurrentCell In ValidHeaders.Cells
        Counter = Counter + 1
        Temp(Counter, 1) = CurrentCell.Address
        Temp(Counter, 2) = CurrentCell.Column
    Next CurrentCell
    
    If ValidHeaders.Cells.Count > 1 Then
        ' We need sorting only we have more than one valid headers. Otherwise sort convert 2D array to vector.
        Temp = Application.WorksheetFunction.Sort(Temp, 2, 1)
    End If
    
    Dim SortedValidHeaders As Range
    
    Dim FirstColumnIndex  As Long
    FirstColumnIndex = LBound(Temp, 2)
    Dim RowIndex As Long
    For RowIndex = LBound(Temp, 1) To UBound(Temp, 1)
        
        Dim TempCell As Range
        Set TempCell = SelectedCells.Worksheet.Range(Temp(RowIndex, FirstColumnIndex))
        
        ' In case of spill range blank shows as zero.
        Dim IsValidHeaderCell As Boolean
        IsValidHeaderCell = (Trim(TempCell.Value) <> vbNullString And Trim(TempCell.Value) <> 0)
        
        If IsValidHeaderCell Then
            Set SortedValidHeaders = UnionOfNonExistableRange(SortedValidHeaders, TempCell)
        End If
        
    Next RowIndex
    
    Set GetSelectedHeaders = SortedValidHeaders
    
End Function

Public Function EscapeDoubleQuote(ByVal UnEscapePart As String) As String
    EscapeDoubleQuote = Replace(UnEscapePart, DOUBLE_QUOTE, DOUBLE_QUOTE & DOUBLE_QUOTE)
End Function
