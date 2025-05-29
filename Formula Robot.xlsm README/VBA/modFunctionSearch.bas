Attribute VB_Name = "modFunctionSearch"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Search Workbook For Specified Functions
' Description:            Searches cells, names, and conditional formatting for use of user specified functions and reports findings.
' Macro Expression:       modFunctionSearch.SearchFunctions({{UserSpecifiedFunctions}},[[NewTableTargetToRight]])
' Generated:              2025-05-17 09:21 AM
'----------------------------------------------------------------------------------------------------
Public Sub SearchFunctions(ByVal ListOfFnsCSV As String, ByVal DumpToCell As Range)
    
    Dim Data As Variant
    
    With New FunctionsFinder
        .ListOfFnsCSV = ListOfFnsCSV
        Set .SearchInBook = DumpToCell.Worksheet.Parent
        .SearchVolatileFunctions
        
        If .IsNoFunctionFound Then
            
            Dim Message As String
            If .SearchFnCount <= 1 Then
                Message = "The function was not found in " & DumpToCell.Worksheet.Parent.Name & "."
            Else
                Message = "None of the functions were found in " & DumpToCell.Worksheet.Parent.Name & "."
            End If
            
            MsgBox Message, vbInformation + vbOKOnly, "Search Workbook For Functions"
            Exit Sub
            
        End If
            
        Data = .SearchOutput
    End With
    
    ' Ignore formula column.
    Set DumpToCell = DumpToCell.Resize(UBound(Data, 1) - LBound(Data, 1) + 1, UBound(Data, 2) - LBound(Data, 1))
    
    If IsBlankRange(DumpToCell) Then
        DumpToCell.Value = Data
        
        Dim FirstColumnIndex  As Long
        FirstColumnIndex = LBound(Data, 2)
        Dim RowIndex As Long
        For RowIndex = LBound(Data, 1) + 1 To UBound(Data, 1)
            
            If Data(RowIndex, FirstColumnIndex + 2) <> "Name" Then
                Dim CurrentCell As Range
                Set CurrentCell = DumpToCell.Cells(RowIndex, 1)
                CurrentCell.Worksheet.Hyperlinks.Add CurrentCell, Address:=vbNullString, SubAddress:=Data(RowIndex, FirstColumnIndex)
            End If
            
        Next RowIndex
        
        DumpToCell.Worksheet.ListObjects.Add xlSrcRange, DumpToCell, , xlYes
        
        AutoFitRange DumpToCell, 100, 14
        
    Else
        MsgBox "Not enough blank area. Needed blank cell area: " & DumpToCell.Address, vbInformation + vbOKOnly, "Search Workbook For Functions"
    End If
    
End Sub




