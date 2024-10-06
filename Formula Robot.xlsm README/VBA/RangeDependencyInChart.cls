VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeDependencyInChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UndeclaredVariable, ImplicitActiveSheetReference
'@PredeclaredId

Option Explicit

Private Type TRangeDependencyInChart
    Book As Workbook
    SheetNamePrefix As String
    BookNamePrefix As String
    AllDependency As Collection
    PutDependencyOnTable As ListObject
    NamedRangeNameOnly As String
End Type

Private This As TRangeDependencyInChart

Public Property Get PutDependencyOnTable() As ListObject
    Set PutDependencyOnTable = This.PutDependencyOnTable
End Property

Public Property Set PutDependencyOnTable(ByVal RHS As ListObject)
    Set This.PutDependencyOnTable = RHS
End Property

Public Property Get AllDependency() As Collection
    Set AllDependency = This.AllDependency
End Property

Public Property Set AllDependency(ByVal RHS As Collection)
    Set This.AllDependency = RHS
End Property

Public Property Get Book() As Workbook
    Set Book = This.Book
End Property

Public Property Set Book(ByVal RHS As Workbook)
    Set This.Book = RHS
End Property

Public Property Get SheetNamePrefix() As String
    SheetNamePrefix = This.SheetNamePrefix
End Property

Public Property Let SheetNamePrefix(ByVal RHS As String)
    This.SheetNamePrefix = RHS
End Property

Private Sub UpdateBasicStat(ByVal GivenRange As Range)
    
    ' Set up basic statistics for the given range
    
    Set This.Book = GivenRange.Worksheet.Parent
    This.SheetNamePrefix = AddSingleQuoteIfAnyNonAlphanumeric(GivenRange.Worksheet.Name)
    This.BookNamePrefix = AddSingleQuoteIfAnyNonAlphanumeric(WorkbookNameFromRange(GivenRange))
    Set This.AllDependency = New Collection
    
End Sub

Private Sub ClearCatches()
    
    ' Clear all the cached variables and objects
    
    Set This.Book = Nothing
    This.SheetNamePrefix = vbNullString
    Set This.AllDependency = Nothing

End Sub

Private Function AddSingleQuoteIfAnyNonAlphanumeric(ByVal ToText As String) As String
    
    Dim Result As String
    If IsAnyNonAlphanumeric(ToText) Then
        Result = SINGLE_QUOTE & Replace(ToText, SINGLE_QUOTE, SINGLE_QUOTE & SINGLE_QUOTE) _
                 & SINGLE_QUOTE
    Else
        Result = ToText
    End If
    
    AddSingleQuoteIfAnyNonAlphanumeric = Result
    
End Function

Public Sub ExtractAllNamedRangeRef(ByVal ForName As Name, ByVal FromBook As Workbook)
    
    Set This.Book = FromBook
    If IsLocalScopeNamedRange(ForName) Then
        This.SheetNamePrefix = Text.BeforeDelimiter(ForName.NameLocal, EXCLAMATION_SIGN)
    ElseIf Not ForName.RefersToRange Is Nothing Then
        This.SheetNamePrefix = AddSingleQuoteIfAnyNonAlphanumeric(ForName.RefersToRange.Worksheet.Name)
    End If
    This.NamedRangeNameOnly = ExtractNameFromLocalNameRange(ForName.NameLocal)
    This.BookNamePrefix = AddSingleQuoteIfAnyNonAlphanumeric(FromBook.Name)
    Set This.AllDependency = New Collection
    FindAllNamedRangeDependency ForName
    
End Sub

Public Sub RenameNamedRange(ByVal NewName As String)
    
    Dim CurrentItem As ChartInfo
    For Each CurrentItem In This.AllDependency
        Dim NewFormulaText As String
        NewFormulaText = CallByName(CurrentItem.FoundOnObject, CurrentItem.PropertyName, VbGet)
        NewFormulaText = ReplaceTokenWithNewToken(NewFormulaText, This.SheetNamePrefix & EXCLAMATION_SIGN & This.NamedRangeNameOnly _
                                                                 , This.SheetNamePrefix & EXCLAMATION_SIGN & NewName)
        NewFormulaText = ReplaceTokenWithNewToken(NewFormulaText, This.BookNamePrefix & EXCLAMATION_SIGN & This.NamedRangeNameOnly _
                                                                 , This.BookNamePrefix & EXCLAMATION_SIGN & NewName)
        CallByName CurrentItem.FoundOnObject, CurrentItem.PropertyName, VbLet, NewFormulaText
    Next CurrentItem
    
End Sub

Public Sub ApplyNameRange(ByVal CurrentName As Name)
    
    ' Apply the named range to the charts that reference it
    
    Dim RefersToRange As Range
    Set RefersToRange = CurrentName.RefersToRange
    FindAllRangeDependency RefersToRange
    If This.AllDependency.Count = 0 Then Exit Sub
    Dim ReplaceWith As String
    ReplaceWith = GetNamedRangeRefText(CurrentName)
    Dim CurrentItem As ChartInfo
    For Each CurrentItem In This.AllDependency
        If CurrentItem.IsWholeSearchRangeBeingUsed Then
            Dim FindAddress As String
            If CurrentItem.OnSheet = RefersToRange.Worksheet.Name Then
                FindAddress = GetSheetRefForRangeReference(CurrentItem.OnSheet) & CurrentItem.RangeAddress
            Else
                FindAddress = CurrentItem.RangeAddress
            End If
            Dim PropertyValue As String
            PropertyValue = CallByName(CurrentItem.FoundOnObject, CurrentItem.PropertyName, VbGet)
            CallByName CurrentItem.FoundOnObject, CurrentItem.PropertyName _
                                                 , VbLet, VBA.Replace(PropertyValue, FindAddress, ReplaceWith)
        End If
    Next CurrentItem
    
End Sub

Private Function GetNamedRangeRefText(ByVal CurrentName As Name) As String
    
    ' Get the reference text for the named range based on its scope (local or global)
    
    Dim ReplaceWith As String
    If IsLocalScopeNamedRange(CurrentName.NameLocal) Then
        ReplaceWith = CurrentName.NameLocal
    Else
        Dim WorkbookName As String
        WorkbookName = CurrentName.RefersToRange.Worksheet.Parent.Name
        If Text.Contains(WorkbookName, ONE_SPACE) Then
            ReplaceWith = SINGLE_QUOTE & WorkbookName & SINGLE_QUOTE & EXCLAMATION_SIGN & CurrentName.Name
        Else
            ReplaceWith = WorkbookName & EXCLAMATION_SIGN & CurrentName.Name
        End If
    End If
    GetNamedRangeRefText = ReplaceWith
    
End Function

Private Sub FindAllRangeDependency(ByVal GivenRange As Range)

    ' Find all the chart objects that reference the given range and update basic statistics
    
    UpdateBasicStat GivenRange
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Book.Worksheets
        ProcessSheetForSeriesForRange CurrentSheet, GivenRange
    Next CurrentSheet
    
End Sub

Private Sub FindAllNamedRangeDependency(ByVal ForName As Name)

    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In Book.Worksheets
        ProcessSheetForSeriesForNamedRange CurrentSheet, ForName
    Next CurrentSheet
    
End Sub

Public Sub SendDataToSheet(ByVal GivenRange As Range, ByVal PutToRange As Range)
    
    ' Send the dependency data to a specified range on a sheet
    
    FindAllRangeDependency GivenRange
    If This.AllDependency.Count = 0 Then
        MsgBox "No Reference found on any chart", vbInformation, "SendDataToSheet"
        Exit Sub
    End If
    Dim AllData As Variant
    AllData = GetDependencyData(This.AllDependency, True)
    Dim PutDependencyDataOnRange As Range
    Set PutDependencyDataOnRange = PutToRange.Resize(AllDependency.Count + 1, UBound(AllData, 2))
    If Not IsAllCellBlank(PutDependencyDataOnRange) Then
        MsgBox "Unable to write dependency info to worksheet. Existing data in range " _
               & PutDependencyDataOnRange.Address & _
               " would be overwritten. Please clear enough space and run again.", vbExclamation + vbOKOnly, APP_NAME
        Exit Sub
    End If
    
    PutDependencyDataOnRange.Value = AllData
    Set This.PutDependencyOnTable = PutToRange.Parent.ListObjects.Add(xlSrcRange, PutDependencyDataOnRange, , xlYes)
    This.PutDependencyOnTable.ListColumns(3).Range.ColumnWidth = 50
    This.PutDependencyOnTable.ListColumns(5).Range.ColumnWidth = 50
    This.PutDependencyOnTable.Range.WrapText = False
    ScrollToDependencyDataRange This.PutDependencyOnTable
    ClearCatches
    
End Sub

Private Function GetDependencyData(ByVal AllDependency As Collection, ByVal IsWithHeader As Boolean) As Variant
    
    ' Prepare the dependency data in a 2D array
    
    Dim AllData As Variant
    If IsWithHeader Then
        ReDim AllData(0 To AllDependency.Count, 1 To 6)
        AllData(0, 1) = "On Sheet"
        AllData(0, 2) = "Chart Object Name"
        AllData(0, 3) = "Formula"
        AllData(0, 4) = "VBA Syntax"
        AllData(0, 5) = "Is Whole Search Range Being Used?"
        AllData(0, 6) = "Range Address"
    Else
        ReDim AllData(1 To AllDependency.Count, 1 To 6)
    End If
    
    Dim CurrentChartInfo As ChartInfo
    Dim CurrentItemIndex As Long
    For CurrentItemIndex = 1 To AllDependency.Count
        Set CurrentChartInfo = AllDependency.Item(CurrentItemIndex)
        AllData(CurrentItemIndex, 1) = CurrentChartInfo.OnSheet
        AllData(CurrentItemIndex, 2) = CurrentChartInfo.ChartObjectName
        AllData(CurrentItemIndex, 3) = SINGLE_QUOTE & CurrentChartInfo.Formula
        AllData(CurrentItemIndex, 4) = CurrentChartInfo.VBASyntax
        AllData(CurrentItemIndex, 5) = CurrentChartInfo.IsWholeSearchRangeBeingUsed
        AllData(CurrentItemIndex, 6) = CurrentChartInfo.RangeAddress
    Next CurrentItemIndex
    GetDependencyData = AllData
    
End Function

Private Sub ProcessSheetForSeriesForRange(ByVal CurrentSheet As Worksheet, ByVal GivenRange As Range)
    
    ' Process the worksheet for chart series
    
    Dim CurrentChartObject As ChartObject
    For Each CurrentChartObject In CurrentSheet.ChartObjects
        Dim CurrentSeries As Series
        Dim Counter As Long
        For Counter = 1 To CurrentChartObject.Chart.SeriesCollection.Count
            Set CurrentSeries = CurrentChartObject.Chart.SeriesCollection.Item(Counter)
            Dim VBASyntax As String
            VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name, "FormulaLocal", Counter)
            Dim FormulaText As String
            FormulaText = GetSeriesFormulaIfErrorVBNullString(CurrentSeries)
            FindAndCreateDependencyForRange FormulaText, GivenRange, CurrentSheet, VBASyntax, CurrentChartObject, CurrentSeries
            ProcessSheetForDataLabels CurrentSheet, GivenRange, CurrentChartObject, CurrentSeries, Counter
        Next Counter
    Next CurrentChartObject
    
End Sub

Private Sub ProcessSheetForSeriesForNamedRange(ByVal CurrentSheet As Worksheet, ByVal ForName As Name)
    
    ' Process the worksheet for chart series
    
    Dim CurrentChartObject As ChartObject
    For Each CurrentChartObject In CurrentSheet.ChartObjects
        Dim CurrentSeries As Series
        Dim Counter As Long
        For Counter = 1 To CurrentChartObject.Chart.SeriesCollection.Count
            Set CurrentSeries = CurrentChartObject.Chart.SeriesCollection.Item(Counter)
            Dim VBASyntax As String
            VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name, "FormulaLocal", Counter)
            Dim FormulaText As String
            FormulaText = GetSeriesFormulaIfErrorVBNullString(CurrentSeries)
            FindAndCreateDependencyForNamedRange FormulaText, ForName, CurrentSheet, VBASyntax, CurrentChartObject, CurrentSeries
            ProcessSheetForDataLabelsForNamedRange CurrentSheet, ForName, CurrentChartObject, CurrentSeries, Counter
        Next Counter
    Next CurrentChartObject
    
End Sub

Private Sub FindAndCreateDependencyForRange(ByVal FormulaText As String, ByVal GivenRange As Range _
                                                                , ByVal CurrentSheet As Worksheet _
                                                                 , ByVal VBASyntax As String _
                                                                  , ByVal CurrentChartObject As ChartObject _
                                                                   , ByVal FoundOnObject As Object)
    
    ' Find and create the dependency info for the given formula text
    
    If Text.Contains(FormulaText, This.SheetNamePrefix) Then
        Dim AllRangeAddress As Variant
        AllRangeAddress = GetDirectPrecedents(FormulaText, GivenRange.Worksheet)
        Dim TempRange As Range
        Dim RangeAddress As Variant
        
        For Each RangeAddress In AllRangeAddress
            If RangeAddress <> vbNullString Then
                Set TempRange = Range(CStr(RangeAddress))
                If HasIntersection(TempRange, GivenRange) Then
                    Dim CurrentChartInfo As ChartInfo
                    Set CurrentChartInfo = CreateNewChartInfo(CurrentSheet, CurrentChartObject _
                                                                           , TempRange.Address = GivenRange.Address _
                                                                            , CStr(RangeAddress) _
                                                                             , FoundOnObject, "FormulaLocal", VBASyntax)
                    This.AllDependency.Add CurrentChartInfo
                End If
            End If
        Next RangeAddress
    End If
    
End Sub

Private Sub FindAndCreateDependencyForNamedRange(ByVal FormulaText As String, ByVal ForName As Name _
                                                                             , ByVal CurrentSheet As Worksheet _
                                                                              , ByVal VBASyntax As String _
                                                                               , ByVal CurrentChartObject As ChartObject _
                                                                                , ByVal FoundOnObject As Object)
    
    ' Find and create the dependency info for the given formula text
    
    If Not (Text.Contains(FormulaText, This.SheetNamePrefix) Or Text.Contains(FormulaText, This.BookNamePrefix)) Then
        Exit Sub
    End If
    
    Dim AllRangeAddress As Variant
    AllRangeAddress = GetDirectPrecedents(FormulaText, CurrentSheet)
    Dim RangeAddress As Variant
        
    For Each RangeAddress In AllRangeAddress
            
        Dim IsNamedRangeFound As Boolean
        ' If local scoped named range then range ref would be equal to the name of the named range.
        ' I have seen both workbook name and sheet name as prefix.
        IsNamedRangeFound = ( _
                            RangeAddress = ForName.Name _
                            Or RangeAddress = This.SheetNamePrefix & EXCLAMATION_SIGN & ForName.Name _
                            Or RangeAddress = This.BookNamePrefix & EXCLAMATION_SIGN & ForName.Name _
                            )
            
        If IsNamedRangeFound Then
            Dim CurrentChartInfo As ChartInfo
            Set CurrentChartInfo = CreateNewChartInfo(CurrentSheet, CurrentChartObject _
                                                                   , True _
                                                                    , CStr(RangeAddress) _
                                                                     , FoundOnObject, "FormulaLocal", VBASyntax)
            This.AllDependency.Add CurrentChartInfo
        End If
    Next RangeAddress
    
End Sub

Public Sub ProcessSheetForDataLabels(ByVal CurrentSheet As Worksheet, ByVal GivenRange As Range _
                                                                     , ByVal CurrentChartObject As ChartObject _
                                                                      , ByVal CurrentSeries As Series _
                                                                       , ByVal SeriesCounter As Long)
    
    ' Process the worksheet for data labels in the chart series
    
    If GivenRange.Cells.Count > 1 Then Exit Sub
    If Not CurrentSeries.HasDataLabels Then Exit Sub
    Dim Counter As Long
    For Counter = 1 To CurrentSeries.DataLabels.Count
        Dim CurrentDataLabel As DataLabel
        Set CurrentDataLabel = CurrentSeries.DataLabels(Counter)
        Dim VBASyntax As String
        VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name _
                                                                 , "DataLabels(" & Counter & ").FormulaLocal" _
                                                                  , SeriesCounter)
        FindAndCreateDependencyForRange CurrentDataLabel.FormulaLocal, GivenRange _
                                                              , CurrentSheet _
                                                               , VBASyntax _
                                                                , CurrentChartObject _
                                                                 , CurrentDataLabel
    Next Counter
    
End Sub

Public Sub ProcessSheetForDataLabelsForNamedRange(ByVal CurrentSheet As Worksheet, ByVal ForName As Name _
                                                                     , ByVal CurrentChartObject As ChartObject _
                                                                      , ByVal CurrentSeries As Series _
                                                                       , ByVal SeriesCounter As Long)
    
    ' Process the worksheet for data labels in the chart series
    If Not CurrentSeries.HasDataLabels Then Exit Sub
    Dim Counter As Long
    For Counter = 1 To CurrentSeries.DataLabels.Count
        Dim CurrentDataLabel As DataLabel
        Set CurrentDataLabel = CurrentSeries.DataLabels(Counter)
        Dim VBASyntax As String
        VBASyntax = GetChartSeriesProperty(CurrentSheet.CodeName, CurrentChartObject.Name _
                                                                 , "DataLabels(" & Counter & ").FormulaLocal" _
                                                                  , SeriesCounter)
        FindAndCreateDependencyForNamedRange CurrentDataLabel.FormulaLocal, ForName _
                                                              , CurrentSheet _
                                                               , VBASyntax _
                                                                , CurrentChartObject _
                                                                 , CurrentDataLabel
    Next Counter
    
End Sub

Private Function GetSeriesFormulaIfErrorVBNullString(ByVal CurrentSeries As Series) As String
    
    ' Get the formula of the series or return empty string in case of an error
    
    On Error Resume Next
    GetSeriesFormulaIfErrorVBNullString = CurrentSeries.FormulaLocal
    On Error GoTo 0
    
End Function

Private Function CreateNewChartInfo(ByVal CurrentSheet As Worksheet, ByVal CurrentChartObject As ChartObject _
                                                                    , ByVal IsWholeSearchRangeBeingUsed As Boolean _
                                                                     , ByVal RangeAddress As String _
                                                                      , ByVal FoundOnObject As Object _
                                                                       , ByVal PropertyName As String _
                                                                        , ByVal VBASyntax As String) As ChartInfo
    
    ' Create a new ChartInfo object with the provided information
    
    Dim CurrentChartInfo As ChartInfo
    Set CurrentChartInfo = New ChartInfo
    With CurrentChartInfo
        .OnSheet = CurrentSheet.Name
        .ChartObjectName = CurrentChartObject.Name
        .Formula = CallByName(FoundOnObject, PropertyName, VbGet)
        .VBASyntax = VBASyntax
        .IsWholeSearchRangeBeingUsed = IsWholeSearchRangeBeingUsed
        .RangeAddress = RangeAddress
        .PropertyName = PropertyName
        Set .FoundOnObject = FoundOnObject
    End With
    
    Set CreateNewChartInfo = CurrentChartInfo

End Function

Public Function GetChartSeriesProperty(ByVal SheetCodeName As String _
                                       , ByVal ChartObjectName As String _
                                        , ByVal PropertyName As String, ByVal Counter As Long) As String
    
    ' Get the property of the chart series
    
    If Counter = 0 Then
        GetChartSeriesProperty = Me.GetChartProperty(SheetCodeName, ChartObjectName, "SeriesCollection." & PropertyName)
    Else
        GetChartSeriesProperty = Me.GetChartProperty(SheetCodeName, ChartObjectName, "SeriesCollection.Item(" & Counter & ")." & PropertyName)
    End If
    
End Function

Public Function GetChartProperty(ByVal SheetCodeName As String _
                                 , ByVal ChartObjectName As String _
                                  , ByVal PropertyName As String) As String
    
    ' Get the chart property
    
    GetChartProperty = SheetCodeName & ".ChartObjects(""" & ChartObjectName & QUOTES_AND_FIRST_PAREN_CLOSE _
                       & ".Chart." & PropertyName
End Function


