Attribute VB_Name = "modTableLookupLambda"
Option Explicit
Option Private Module

Private Sub Test()
    GenerateTableLookupLambdas ActiveCell, Selection
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Generate Table Lookup Lambdas
' Description:            Generate lambdas for each column of an Excel Table.
' Macro Expression:       modTableLookupLambda.GenerateTableLookupLambdas([ActiveCell])
' Generated:              04/26/2024 04:11 PM
'----------------------------------------------------------------------------------------------------
Public Sub GenerateTableLookupLambdas(ByVal TableCell As Range, ByVal SelectedCells As Range)
    
    ' If not table nor spill range
    If IsNothing(TableCell.ListObject) And Not TableCell.HasSpill Then
        Exit Sub
    End If
    
    CreateTableOrNamedRangeIfSpillRange TableCell
    
    Dim DataRangeWithHeader As Range
    Dim TableName As String
    Dim IsTable As Boolean
    
    If IsNotNothing(TableCell.ListObject) Then
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
    
    Dim NewlyCreatedLambdas As Collection
    Set NewlyCreatedLambdas = New Collection
    
    ' Add method will replace refersto if already exist.
    With Book
        NewlyCreatedLambdas.Add 1, TableName & "." & "Select"
        .Names.Add TableName & "." & "Select", LambdaGenerator.FilterLambda
    End With
    
    Dim DefPart As String
    DefPart = LambdaGenerator.DefPart
    
    Dim FilterInvocationPart As String
    FilterInvocationPart = Text.RemoveFromEndIfPresent(Replace(Replace(DefPart, "[", vbNullString), "]", vbNullString), LIST_SEPARATOR)
    
    Dim ColIndex As Long
    For ColIndex = 1 To DataRangeWithHeader.Columns.CountLarge
        
        Dim CurrentColParamName As String
        CurrentColParamName = DataRangeWithHeader.Cells(1, ColIndex).Value
        
        Dim ReturnColDataPart As String
        ReturnColDataPart = DOUBLE_QUOTE & EscapeDoubleQuote(CurrentColParamName) & DOUBLE_QUOTE
        
        Dim Lambda As String
        Lambda = GetCurrentColLambda(DefPart, TableName, ReturnColDataPart, FilterInvocationPart)
        
        CurrentColParamName = MakeValidDefinedName(CurrentColParamName, False, True)
        
        If CurrentColParamName <> "Select" Then
            NewlyCreatedLambdas.Add 1, TableName & "." & CurrentColParamName
            Book.Names.Add TableName & "." & CurrentColParamName, Replace(Lambda, vbNewLine, Chr$(10))
        End If
        
    Next ColIndex
    
    DeletePreviouslyCreatedLambdas NewlyCreatedLambdas, Book, TableName
    
End Sub

Private Sub DeletePreviouslyCreatedLambdas(ByVal NewlyCreatedLambdas As Collection _
                                           , ByVal SourceBook As Workbook _
                                            , ByVal TableName As String)
    Dim CurrentName As Name
    For Each CurrentName In SourceBook.Names
        
        If Not IsBuiltInName(CurrentName) Then
        
            If Text.IsStartsWith(CurrentName.Name, TableName & ".") Then
                If IsLambdaFunction(CurrentName.RefersTo) _
                   And Not IsExistInCollection(NewlyCreatedLambdas, CurrentName.Name) Then
                    CurrentName.Delete
                End If
            End If
        
        End If
        
    Next CurrentName
    
End Sub

Public Function GetCurrentColLambda(ByVal DefPart As String _
                                     , ByVal TableName As String _
                                      , ByVal ReturnColDataPart As String _
                                       , ByVal FilterInvocationPart As String) As String

    Dim Code As String
    Code = "=LAMBDA(" & DefPart & " LET(" & vbNewLine _
           & "   _FilteredDataWithHeader, " & TableName & ".Select(" & ReturnColDataPart & "," & FilterInvocationPart & ")," & vbNewLine _
           & "   _RowCount,ROWS(_FilteredDataWithHeader)," & vbNewLine _
           & "   _Result, IF(_RowCount=1,NA(),IF(_RowCount=2,INDEX(_FilteredDataWithHeader,2,1),DROP(_FilteredDataWithHeader,1)))," & vbNewLine _
           & "   _Result" & vbNewLine _
           & "))"

    GetCurrentColLambda = Code

End Function

Private Sub CreateTableOrNamedRangeIfSpillRange(ByVal TableCell As Range)
    
    If IsNotNothing(TableCell.ListObject) Then Exit Sub
    
    If IsInsideNamedRange(TableCell.SpillParent.SpillingToRange) Then
        Exit Sub
    End If
    
    Dim Answer As VbMsgBoxResult
    Answer = MsgBox("Would you like to convert to values as Excel Table for performance purposes?", vbYesNo, "Excel Table")
    
    If Answer = vbYes Then
        
        Dim TableName As String
        TableName = Application.InputBox("Please provide the table name: ", "Table Name", Type:=2)
        
        If TableName = "False" Then Exit Sub
        
        ' Replace with values
        Dim DataRange As Range
        Set DataRange = TableCell.SpillParent.SpillingToRange
        DataRange.Value = DataRange.Value
        
        Dim Table As ListObject
        Set Table = TableCell.Worksheet.ListObjects.Add(xlSrcRange, DataRange, , xlYes)
        
        On Error Resume Next
        Table.Name = TableName
        On Error GoTo 0
        
    ElseIf Answer = vbNo And Not IsInsideNamedRange(TableCell.SpillParent.SpillingToRange) Then
        modNamedRange.AddNameRange TableCell.SpillParent.SpillingToRange, vbNullString
    End If
    
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
    
    If ValidHeaders.Cells.CountLarge > 1 Then
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
        If TempCell.HasSpill Then
            IsValidHeaderCell = (Trim(TempCell.Value) <> vbNullString And Trim(TempCell.Value) <> 0)
        Else
            IsValidHeaderCell = (Trim(TempCell.Value) <> vbNullString)
        End If
        
        If IsValidHeaderCell Then
            Set SortedValidHeaders = UnionOfNonExistableRange(SortedValidHeaders, TempCell)
        End If
        
    Next RowIndex
    
    Set GetSelectedHeaders = SortedValidHeaders
    
End Function

Public Function EscapeDoubleQuote(ByVal UnEscapePart As String) As String
    EscapeDoubleQuote = Replace(UnEscapePart, DOUBLE_QUOTE, DOUBLE_QUOTE & DOUBLE_QUOTE)
End Function


