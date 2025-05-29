VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionsFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type TFunctionsFinder
    SearchInBook As Workbook
    SearchFnsList As Scripting.Dictionary
    SearchResult As Collection
    SearchOutput As Variant
    ListOfFnsCSV As String
    IsNoFunctionFound As Boolean
    SearchFnCount As Long
End Type

Private This As TFunctionsFinder

Public Property Get SearchFnCount() As Long
    SearchFnCount = This.SearchFnCount
End Property

Public Property Get IsNoFunctionFound() As Boolean
    IsNoFunctionFound = This.IsNoFunctionFound
End Property

Public Property Get ListOfFnsCSV() As String
    ListOfFnsCSV = This.ListOfFnsCSV
End Property

Public Property Let ListOfFnsCSV(ByVal RHS As String)
    
    This.ListOfFnsCSV = RHS
    Dim Fns As Variant
    Fns = Split(RHS, COMMA)
    
    Set This.SearchFnsList = New Scripting.Dictionary
    This.SearchFnsList.CompareMode = TextCompare
    
    Dim CurrentFn As Variant
    For Each CurrentFn In Fns
        This.SearchFnsList.Add Trim$(CurrentFn), Trim$(CurrentFn)
    Next CurrentFn
    
    This.SearchFnCount = This.SearchFnsList.Count
    
End Property

Public Property Get SearchOutput() As Variant
    
    If IsObject(This.SearchOutput) Then
        Set SearchOutput = This.SearchOutput
    Else
        SearchOutput = This.SearchOutput
    End If
    
End Property

Public Property Get SearchInBook() As Workbook
    Set SearchInBook = This.SearchInBook
End Property

Public Property Set SearchInBook(ByVal RHS As Workbook)
    Set This.SearchInBook = RHS
End Property

Public Sub SearchVolatileFunctions()
    
    SearchInNameManager
    SearchInCellsFormula
    SearchInConditionalFormattings
    
    Dim Result As Variant
    ReDim Result(1 To This.SearchResult.Count + 1, 1 To 4)
    Result(1, 1) = "Location"
    Result(1, 2) = "Functions"
    Result(1, 3) = "Type"
    Result(1, 4) = "Formula"
    
    Dim RowIndex As Long
    RowIndex = 2
    
    This.IsNoFunctionFound = (This.SearchResult.Count = 0)
    
    Dim CurrentBean As SearchFunctionData
    For Each CurrentBean In This.SearchResult
        
        With CurrentBean
            Result(RowIndex, 1) = .Location
            Result(RowIndex, 2) = .Functions
            Result(RowIndex, 3) = .SearchSection
            Result(RowIndex, 4) = .Formula
        End With
        
        RowIndex = RowIndex + 1
        
    Next CurrentBean
    
    This.SearchOutput = Result
    
End Sub

Private Sub SearchInConditionalFormattings()
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In This.SearchInBook.Worksheets
        
        Dim FormulaCells As Range
        On Error Resume Next
        Set FormulaCells = FilterUsingSpecialCells(CurrentSheet.UsedRange, xlCellTypeAllFormatConditions)
        On Error GoTo 0
        If IsNotNothing(FormulaCells) Then
            
            Dim FormulaGroups As Scripting.Dictionary
            Set FormulaGroups = GroupConditionalFormattingsByFormula(FormulaCells)
            
            Dim CurrentKey As Variant
            For Each CurrentKey In FormulaGroups.Keys
                Dim Fns As String
                Fns = ExtractUsedSearchFnsFromFormula(CurrentKey, True)
                
                Dim Area As Range
                Set Area = FormulaGroups.Item(CurrentKey)
                
                If Fns <> vbNullString Then
                    This.SearchResult.Add SearchFunctionData.Create(GetRangeRefWithSheetName(Area), Fns, "Conditional Formatting", CurrentKey)
                End If
                
            Next CurrentKey
            
        End If
        
        DoEvents
    Next CurrentSheet
    
End Sub

Private Function GroupConditionalFormattingsByFormula(ByVal FormulaCells As Range) As Scripting.Dictionary
    
    Dim Groups As Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    
    Dim CurrentArea As Range
    For Each CurrentArea In FormulaCells.Areas
        
        Dim CFs As FormatConditions
        Dim CurrentCell As Range
        For Each CurrentCell In CurrentArea
            
            Set CFs = CurrentCell.FormatConditions
            
            Dim CurrentCF As Object
            For Each CurrentCF In CFs
                If TypeName(CurrentCF) = "FormatCondition" Then
                    AddOrUpdateGroup Groups, CurrentCF.Formula1, CurrentCell
                End If
            Next CurrentCF
            
        Next CurrentCell
        
    Next CurrentArea
    
    Set GroupConditionalFormattingsByFormula = Groups
    
End Function

Private Sub SearchInCellsFormula()
    
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In This.SearchInBook.Worksheets
        
        Dim FormulaCells As Range
        On Error Resume Next
        Set FormulaCells = FilterUsingSpecialCells(CurrentSheet.UsedRange, xlCellTypeFormulas)
        On Error GoTo 0
        If IsNotNothing(FormulaCells) Then
            
            Dim FormulaGroups As Scripting.Dictionary
            Set FormulaGroups = GroupsRangeByFormula(FormulaCells)
            
            Dim CurrentKey As Variant
            For Each CurrentKey In FormulaGroups.Keys
                Dim Fns As String
                Fns = ExtractUsedSearchFnsFromFormula(CurrentKey, True)
                
                Dim Area As Range
                Set Area = FormulaGroups.Item(CurrentKey)
                
                If Fns <> vbNullString Then
                    This.SearchResult.Add SearchFunctionData.Create(GetRangeRefWithSheetName(Area), Fns, "Cell Formula", Area.Cells(1).Formula2)
                End If
                
            Next CurrentKey
            
        End If
        
        DoEvents
        
    Next CurrentSheet

End Sub

Private Function GroupsRangeByFormula(ByVal FormulaCells As Range) As Scripting.Dictionary
    
    Dim Groups As Scripting.Dictionary
    Set Groups = New Scripting.Dictionary
    
    Dim CurrentArea As Range
    For Each CurrentArea In FormulaCells.Areas
        
        Dim Formulas As Variant
        Formulas = CurrentArea.Formula2R1C1
        If Not IsArray(Formulas) Then
            AddOrUpdateGroup Groups, Formulas, CurrentArea
        Else

            Dim RowIndex As Long
            For RowIndex = LBound(Formulas, 1) To UBound(Formulas, 1)
                Dim ColumnIndex As Long
                For ColumnIndex = LBound(Formulas, 2) To UBound(Formulas, 2)
                    AddOrUpdateGroup Groups, Formulas(RowIndex, ColumnIndex), CurrentArea.Cells(RowIndex, ColumnIndex)
                Next ColumnIndex
            Next RowIndex
        End If
        
    Next CurrentArea
    
    Set GroupsRangeByFormula = Groups
    
End Function

Private Sub AddOrUpdateGroup(ByRef Groups As Scripting.Dictionary _
                             , ByVal Formula As String _
                              , ByVal CellRange As Range)
                              
    If Groups.Exists(Formula) Then
        Set Groups(Formula) = UnionOfNonExistableRange(Groups(Formula), CellRange)
    Else
        Groups.Add Formula, CellRange
    End If
    
End Sub

Private Sub SearchInNameManager()
    
    Dim CurrentName As Name
    For Each CurrentName In This.SearchInBook.Names
        If Not IsBuiltInName(CurrentName) Then
            
            Dim Fns As String
            Fns = ExtractUsedSearchFnsFromFormula(CurrentName.RefersTo, False)
            
            If Fns <> vbNullString Then
                This.SearchResult.Add SearchFunctionData.Create(CurrentName.Name, Fns, "Name", CurrentName.RefersTo)
            End If
            
        End If
    Next CurrentName
    
End Sub

Private Function ExtractUsedSearchFnsFromFormula(ByVal Formula As String, ByVal IsR1C1 As Boolean) As String
    
    Dim Result As String
    Result = vbNullString
    If IsAnySearchFnFoundByStringSearch(Formula) Then
        
        Dim UsedFns As Variant
        UsedFns = GetUsedFunctions(Formula, IsR1C1)
        
        Dim CurrentFn As Variant
        For Each CurrentFn In UsedFns
            If This.SearchFnsList.Exists(CurrentFn) Then
                Result = Result & CurrentFn & ", "
            End If
        Next CurrentFn
        
    End If
    
    If Result <> vbNullString Then
        Result = Left$(Result, Len(Result) - 2)
    End If
    
    ExtractUsedSearchFnsFromFormula = Result
    
End Function

Private Function IsAnySearchFnFoundByStringSearch(ByVal Formula As String) As Boolean
    
    Dim Result As Boolean
    Result = False
    
    Dim FnName As Variant
    For Each FnName In This.SearchFnsList.Keys
        If Text.Contains(Formula, FnName) Then
            Result = True
            Exit For
        End If
    Next FnName
    
    IsAnySearchFnFoundByStringSearch = Result
    
End Function

Private Sub Class_Initialize()
    Set This.SearchResult = New Collection
End Sub

Private Sub Class_Terminate()
    Set This.SearchFnsList = Nothing
    Set This.SearchResult = Nothing
End Sub


