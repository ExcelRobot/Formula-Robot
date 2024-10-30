Attribute VB_Name = "modNamedRange"

' @Folder "NamedRange.Driver"
' @IgnoreModule SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed
Option Explicit
Option Private Module

Private Function IsStartedWithGivenPrefix(ByVal CellValue As Variant, ByVal IgnorePrefix As String) As Boolean

    If IsError(CellValue) Then
        IsStartedWithGivenPrefix = False
    Else
        IsStartedWithGivenPrefix = Text.IsStartsWith(CStr(CellValue), IgnorePrefix)
    End If

End Function
'
'@TODO: Need to update when creating a global named range and we have a local scoped named range with same name
' then it doesn't create the named range rather update the local one refersTo.
' Ref: https://stackoverflow.com/questions/14902754/trying-to-set-global-named-range-but-local-range-ends-up-getting-set
Public Sub AddNameRange(ByVal SelectionRange As Range, ByVal IgnorePrefix As String _
                                                      , Optional ByVal IsLocal As Boolean _
                                                       , Optional ByVal ScopeSheet As Worksheet _
                                                        , Optional ByVal IsAbsoluteRow As Boolean = True _
                                                         , Optional ByVal IsAbsoluteColumn As Boolean = True _
                                                          , Optional ByVal IsReassign As Boolean = False)

    ' Start of logging
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddNameRange"

    ' Check if the SelectionRange has a dynamic formula
    Dim IsDynamicFormula As Boolean
    IsDynamicFormula = HasDynamicFormula(SelectionRange)

    ' If dynamic formula is present, further validate it
    If IsDynamicFormula Then IsDynamicFormula = (SelectionRange.Address = SelectionRange.Cells(1).SpillParent.SpillingToRange.Address)

    ' Define the DefaultName based on whether the SelectionRange has a dynamic formula
    Dim DefaultName As String
    If IsDynamicFormula Then
        DefaultName = FindDefaultName(SelectionRange.Cells(1, 1).SpillParent, IgnorePrefix)
    Else
        DefaultName = FindDefaultName(SelectionRange, IgnorePrefix)
    End If

    ' Transform the DefaultName to a FinalDefineName
    DefaultName = GetFinalDefineName(DefaultName, False)
    ' Exit if DefaultName is empty
    If DefaultName = vbNullString Then Exit Sub

    ' Adjust the SelectionRange based on whether the rows and columns are absolute
    If Not IsAbsoluteRow Then Set SelectionRange = SelectionRange.Parent.Cells(1, SelectionRange.Column)
    If Not IsAbsoluteColumn Then Set SelectionRange = SelectionRange.Parent.Cells(SelectionRange.Row, 1)

    ' Apply name to the range
    ApplyNameRange SelectionRange, DefaultName _
                                  , IsDynamicFormula, IsLocal _
                                                     , ScopeSheet, True, True, , IsAbsoluteRow _
                                                                                , IsAbsoluteColumn, IsReassign
    ' End of logging
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddNameRange"

End Sub

Private Sub ApplyNameRange(ByVal SelectionRange As Range _
                           , ByVal NameOfNamedRange As String _
                            , ByVal IsDynamicFormula As Boolean _
                             , Optional ByVal IsLocal As Boolean _
                              , Optional ByVal ScopeSheet As Worksheet _
                               , Optional ByVal IsShowPopup As Boolean = False _
                                , Optional ByVal IsCheckForStructuredReference As Boolean = False _
                                 , Optional ByVal IsForceToOverwrite As Boolean = False _
                                  , Optional ByVal IsAbsoluteRow As Boolean = True _
                                   , Optional ByVal IsAbsoluteColumn As Boolean = True _
                                    , Optional ByVal IsReassign As Boolean = False)

    ' Start of logging
    Logger.Log TRACE_LOG, "Enter modNamedRange.ApplyNameRange"

    ' Get the first cell reference if merged and check if it's a valid range to name
    Set SelectionRange = GetFirstCellRefIfMerged(SelectionRange)
    If Not IsValidRangeToName(SelectionRange, IsShowPopup, IsLocal) Then Exit Sub

    ' Make a valid defined name for the named range
    NameOfNamedRange = MakeValidDefinedName(NameOfNamedRange, True)

    ' Find the existing named range
    Dim CurrentName As Name
    Set CurrentName = FindNamedRange(SelectionRange.Worksheet.Parent, NameOfNamedRange)

    ' Check if the name conflicts with an existing named range
    If IsConflictName(SelectionRange, IsLocal, CurrentName, NameOfNamedRange) Then

        ' If reassigning, update the current name's reference to the selection range
        If IsReassign Then
            CurrentName.RefersTo = ConvertToReference(SelectionRange, IsDynamicFormula _
                                                                     , IsCheckForStructuredReference _
                                                                      , IsAbsoluteRow, IsAbsoluteColumn)
            Exit Sub
        End If

        ' If a name conflict arises, show a pop-up message to the user
        If IsShowPopup Then
            MsgBox "Unable to apply name " & NameOfNamedRange & ". A range already exists with this name." _
                   , vbExclamation + vbOKOnly, APP_NAME
        End If

        ' If not forcing to overwrite, exit
        If Not IsForceToOverwrite Then Exit Sub

        ' If reassigning but the named range is not found, show a pop-up message to the user
    ElseIf IsReassign And IsNothing(CurrentName) Then
        MsgBox "Unable to reassign " & NameOfNamedRange & ", named range not found." _
               , vbExclamation + vbOKOnly, APP_NAME
        Exit Sub
    End If

    ' Convert the selection range to a reference
    Dim Reference As String
    Reference = ConvertToReference(SelectionRange, IsDynamicFormula _
                                                  , IsCheckForStructuredReference _
                                                   , IsAbsoluteRow, IsAbsoluteColumn)
    Logger.Log DEBUG_LOG, "Reference : " & Reference
    Dim AddToSheet As Worksheet
    Set AddToSheet = SelectionRange.Parent

    ' Create a named range with the reference
    Set CurrentName = CreateNamedRange(Reference, NameOfNamedRange, IsLocal, ScopeSheet, AddToSheet)
    Logger.Log DEBUG_LOG, NameOfNamedRange & " has been assigned to " & Reference

    ' Apply the named range to the chart and formula
    ApplyNameRangeToChartAndFormula NameOfNamedRange, CurrentName, SelectionRange.Worksheet.Parent

    ' End of logging
    Logger.Log TRACE_LOG, "Exit modNamedRange.ApplyNameRange"

End Sub

Private Sub ApplyNameRangeToChartAndFormula(ByVal NameOfNamedRange As String _
                                            , ByVal CurrentName As Name _
                                             , ByVal ApplyToBook As Workbook)

    ApplyName NameOfNamedRange, ApplyToBook
    ApplyNameToChart CurrentName

End Sub

Private Function CreateNamedRange(ByVal Reference As String, ByVal NameOfNamedRange As String _
                                                            , ByVal IsLocal As Boolean _
                                                             , ByVal ScopeSheet As Worksheet _
                                                              , ByVal AddToSheet As Worksheet) As Name

    ' Set Workbook for the scope based on the provided worksheet
    Dim AddToBook As Workbook
    If IsNotNothing(ScopeSheet) Then Set AddToBook = ScopeSheet.Parent
    If IsNotNothing(AddToSheet) Then Set AddToBook = AddToSheet.Parent

    On Error GoTo ErrorHandler
    Dim CalculationType As XlCalculation

    ' Set application calculation to manual for performance optimization
    CalculationType = Application.Calculation
    Application.Calculation = xlCalculationManual
    SwitchHeaderVisibilityOfTableStartAtA1 AddToBook, False

    Dim CurrentName As Name

    ' Add name to the appropriate scope and assign it to the CurrentName
    If IsLocal And IsNotNothing(ScopeSheet) Then
        ScopeSheet.Names.Add Name:=NameOfNamedRange, RefersTo:=Reference
        Set CurrentName = ScopeSheet.Names(NameOfNamedRange)
    ElseIf IsLocal Then
        AddToSheet.Names.Add Name:=NameOfNamedRange, RefersTo:=Reference
        Set CurrentName = AddToSheet.Names(NameOfNamedRange)
    Else
        AddToSheet.Parent.Names.Add Name:=NameOfNamedRange, RefersTo:=Reference
        Set CurrentName = GetNameFromBook(AddToSheet.Parent, NameOfNamedRange)
    End If

    ' Set function return value to the created name
    Set CreateNamedRange = CurrentName

    ' Revert header visibility and application calculation mode back to original
    SwitchHeaderVisibilityOfTableStartAtA1 AddToBook, True
    Application.Calculation = CalculationType
    Exit Function

ErrorHandler:
    ' In case of error, restore application calculation mode and header visibility, then raise the error again
    Application.Calculation = CalculationType
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description
    SwitchHeaderVisibilityOfTableStartAtA1 AddToBook, True
    Err.Raise ErrorNumber, Err.Source, ErrorDescription

End Function

Public Sub SwitchHeaderVisibilityOfTableStartAtA1(ByVal ForBook As Workbook, ByVal IsShowHeader As Boolean)

    ' Check if workbook is defined
    If IsNothing(ForBook) Then Exit Sub

    ' Create a static variable to store a collection of tables
    Static TableCollection As Collection
    Dim CurrentSheet As Worksheet
    Dim Table As ListObject

    ' Show headers if IsShowHeader is True
    If IsShowHeader Then
        For Each Table In TableCollection
            ' Make headers visible
            Table.ShowHeaders = True
        Next Table

        ' Reset table collection
        Set TableCollection = Nothing
    Else
        ' Instantiate the table collection
        Set TableCollection = New Collection

        ' Loop through each sheet in the workbook
        For Each CurrentSheet In ForBook.Worksheets

            ' If a list object starts from cell A1, turn off its header and add it to the collection
            If IsNotNothing(CurrentSheet.Range("A1").ListObject) Then
                Set Table = CurrentSheet.Range("A1").ListObject
                If Table.ShowHeaders Then
                    Table.ShowHeaders = False
                    TableCollection.Add Table
                End If
            End If
        Next CurrentSheet
    End If

End Sub

Private Function GetFirstCellRefIfMerged(ByVal GivenRange As Range) As Range

    If GivenRange.MergeCells Then
        Set GetFirstCellRefIfMerged = GivenRange.Cells(1)
    Else
        Set GetFirstCellRefIfMerged = GivenRange
    End If

End Function

Private Function IsConflictName(ByVal SelectionRange As Range, ByVal IsLocal As Boolean _
                                                              , ByVal CurrentName As Name _
                                                              , ByVal NameOfTheNamedRange As String) As Boolean

    ' Check if the current name is nothing, return false and exit function
    If IsNothing(CurrentName) Then
        IsConflictName = False
        Exit Function
    End If

    IsConflictName = False

    ' Check for conflict when range is not local and current name isn't local either
    If Not IsLocal And Not IsLocalScopeNamedRange(CurrentName.NameLocal) Then
        IsConflictName = True

        ' Check for conflict when range is local and current name is local
    ElseIf IsLocal And IsLocalScopeNamedRange(CurrentName.NameLocal) Then
        Dim LocalName As String
        LocalName = GetSheetRefForRangeReference(SelectionRange.Worksheet.Name) & NameOfTheNamedRange
        ' Check if the name is present in the workbook
        IsConflictName = IsNamePresent(SelectionRange.Worksheet.Parent, LocalName)
    End If

End Function

Public Sub ApplyName(ByVal NameToApply As String, ByVal ApplyInWorkbook As Workbook)

    ' Logging entry into the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.ApplyName"

    ' Preserve original calculation setting and switch to manual for performance
    Dim CalculationType As XlCalculation
    CalculationType = Application.Calculation
    Application.Calculation = xlCalculationManual

    Dim CurrentSheet As Worksheet
    'Allow program to continue if errors are encountered
    On Error Resume Next

    ' Apply names to all cells in each worksheet in the workbook
    For Each CurrentSheet In ApplyInWorkbook.Worksheets
        CurrentSheet.Cells.ApplyNames Names:=NameToApply, IgnoreRelativeAbsolute:=True, _
                                      UseRowColumnNames:=False, OmitColumn:=True, OmitRow:=True, Order:=1, _
                                      AppendLast:=False
    Next CurrentSheet
    'Reset error handling
    On Error GoTo 0

    ' Restore original calculation setting
    Application.Calculation = CalculationType

    ' Logging exit from the function
    Logger.Log TRACE_LOG, "Exit modNamedRange.ApplyName"

End Sub

Private Sub ApplyNameToChart(ByVal CurrentName As Name)

    ' Exit sub if there's nothing in the CurrentName or if its RefersToRange property is nothing
    If IsNothing(CurrentName) Then Exit Sub
    If IsRefersToRangeIsNothing(CurrentName) Then Exit Sub

    'Set error handling
    On Error GoTo ErrorHandler

    ' Applying the name range to the chart dependencies
    RangeDependencyInChart.ApplyNameRange CurrentName
    Exit Sub

ErrorHandler:                                 'Error handling routine

    ' Store the error number and description
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description

    ' Raise error if ErrorNumber is not 0
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        ' This is only for debugging purpose.
        Resume
    End If

End Sub

Private Function GetFinalDefineName(ByVal DefaultName As String, ByVal JustRemoveInvalidChars As Boolean) As String

    ' Logging the start of function execution
    Logger.Log TRACE_LOG, "Enter modNamedRange.GetFinalDefineName"

    ' Input box for user to enter a name for a range, it is validated and returned
    Dim UserGivenName As String

    ' Ensuring default name is a valid name
    DefaultName = MakeValidDefinedName(DefaultName, JustRemoveInvalidChars)

    ' Loop until a valid user defined name is given or user cancels the operation
    Do While True
        ' Prompting user for input
        UserGivenName = InputBox("Enter range name:", "Named Range Name", DefaultName)

        ' If user gives null string or cancels, function exits
        If UserGivenName = vbNullString Or UserGivenName = "False" Then
            GetFinalDefineName = vbNullString
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.GetFinalDefineName"
            Exit Function
        End If

        ' If user gives a new name
        If UserGivenName <> DefaultName Then
            Dim ValidUserGivenName As String
            ' Making valid defined name
            ValidUserGivenName = MakeValidDefinedName(UserGivenName, True)

            ' If name given by user is not valid
            If LCase$(UserGivenName) <> LCase$(ValidUserGivenName) Then
                ' Prompting user for action, option to use valid name, cancel or reenter
                Dim Answer As VbMsgBoxResult
                Answer = MsgBox("The name specified is not a valid range name." & vbNewLine & vbNewLine _
                                & "Use " & ValidUserGivenName & " instead?", vbYesNoCancel, APP_NAME)

                ' If user chooses to use the valid name
                If Answer = vbYes Then
                    GetFinalDefineName = ValidUserGivenName
                    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.GetFinalDefineName"
                    Exit Function

                    ' If user cancels operation
                ElseIf Answer = vbCancel Then
                    GetFinalDefineName = vbNullString
                    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.GetFinalDefineName"
                    Exit Function
                    ' If user chooses to reenter the name

                Else
                    DefaultName = ValidUserGivenName
                End If

                ' If user given name is valid
            Else
                GetFinalDefineName = ValidUserGivenName
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.GetFinalDefineName"
                Exit Function
            End If

            ' If user gives the same name as default name
        Else
            GetFinalDefineName = DefaultName
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.GetFinalDefineName"
            Exit Function
        End If
    Loop

    ' Logging the end of function execution
    Logger.Log TRACE_LOG, "Exit modNamedRange.GetFinalDefineName"

End Function

Private Function FindDefaultName(ByVal FromRange As Range, ByVal IgnorePrefix As String) As String

    ' Logging the start of the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindDefaultName"

    ' Check if FromRange has more than one cell and the first cell is A1. If so, return the worksheet name as the default name.
    If FromRange.Cells.Count > 1 And FromRange.Cells(1).Address = "$A$1" Then
        FindDefaultName = FromRange.Worksheet.Name
        Exit Function
    End If

    ' Initialize cell references for CurrentCell, CellAbove, CellTwoAbove and CellToLeft based on FromRange
    Dim CellAbove As Range
    Dim CellTwoAbove As Range
    Dim CellToLeft As Range
    Dim CurrentCell As Range
    If IsNotNothing(FromRange) Then Set CurrentCell = FromRange.Cells(1)
    If FromRange.Cells(1).Row > 1 Then Set CellAbove = FromRange.Offset(-1).Cells(1)
    If FromRange.Cells(1).Row > 2 Then Set CellTwoAbove = FromRange.Offset(-2).Cells(1)
    If FromRange.Cells(1).Column > 1 Then Set CellToLeft = FromRange.Offset(0, -1).Cells(1)

    ' Check cells above for a suitable name
    If IsNotNothing(CellAbove) Then
        If GetCellValueIfErrorNullString(CellAbove) = vbNullString Then
            ' If cell above is empty, check two cells above
            If IsNotNothing(CellTwoAbove) Then
                If IsProbableDefineName(CellTwoAbove, IgnorePrefix) Then
                    FindDefaultName = CellTwoAbove.Value
                    Exit Function
                End If
            End If
        ElseIf IsProbableDefineName(CellAbove, IgnorePrefix) Then
            ' If cell above is a probable define name, use it as the default name
            FindDefaultName = CellAbove.Value
            Exit Function
        End If
    End If

    ' Scan left until a non-blank cell is found or reach column A
    If IsNotNothing(CellToLeft) Then
        Do While (GetCellValueIfErrorNullString(CellToLeft) = vbNullString _
                  Or IsStartedWithGivenPrefix(CellToLeft.Value, IgnorePrefix)) _
           And CellToLeft.Column > 1 And Not IsCellHidden(CellToLeft)
            Set CellToLeft = CellToLeft.Offset(0, -1).Cells(1)
        Loop

        ' Check if a suitable name is found in the left scan
        If IsProbableDefineName(CellToLeft, IgnorePrefix) Then
            FindDefaultName = CellToLeft.Value
        End If
    End If

    ' If no suitable label is found, check if current cell has a defined name. If so, use it as the default name.
    If FindDefaultName = vbNullString Then
        If IsRangeHasDefinedName(CurrentCell) Then
            FindDefaultName = CurrentCell.Name.Name
        End If
    End If

    ' Logging the end of the function
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindDefaultName"

End Function

Private Function IsProbableDefineName(ByVal CurrentCell As Range, ByVal IgnorePrefix As String) As Boolean

    ' Logging the start of the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsProbableDefineName"

    ' Creating a cell variable and assigning the first cell of CurrentCell to it
    Dim cell As Range
    If IsNotNothing(CurrentCell) Then Set cell = CurrentCell.Cells(1)

    ' Assessing if the cell can be a probable define name by checking for different conditions
    If IsNothing(cell) Then
        ' If the cell is Nothing, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf Application.WorksheetFunction.Trim(cell.Value) = vbNullString Then
        ' If the cell is empty, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf HasDynamicFormula(cell) Then
        ' If the cell contains a dynamic formula, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf TypeName(cell.Value) <> "String" Then
        ' If the cell value is not a string, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf IsInsideNamedRange(cell) Then
        ' If the cell is inside a named range, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf IsInsideTable(cell) And Not IsInsideTableHeader(cell) Then
        ' If the cell is inside a table but not in the header, then it can't be a probable define name
        IsProbableDefineName = False
    ElseIf Text.IsStartsWith(cell.Value, IgnorePrefix) Then
        ' If the cell value starts with a prefix that should be ignored, then it can't be a probable define name
        IsProbableDefineName = False
    Else
        ' If none of the above conditions are met, then it is a probable define name
        IsProbableDefineName = True
    End If

    ' Logging the end of the function
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsProbableDefineName"

End Function

Private Function ConvertToReference(ByVal DataSource As Range _
                                    , Optional ByVal IsDynamicFormula As Boolean _
                                     , Optional ByVal IsCheckForStructuredReference As Boolean = False _
                                      , Optional ByVal IsAbsoluteRow As Boolean = True _
                                       , Optional ByVal IsAbsoluteColumn As Boolean = True) As String

    ' Start of function, logging
    Logger.Log TRACE_LOG, "Enter modNamedRange.ConvertToReference"

    ' Getting the address of the DataSource Range
    Dim Address As String
    Address = DataSource.Address(IsAbsoluteRow, IsAbsoluteColumn)

    ' Creating a SelectionRange based on the DataSource
    Dim SelectionRange As Range
    Set SelectionRange = DataSource
    Set DataSource = DataSource.Cells(1, 1)

    ' Creating the prefix for the sheet name
    Dim SheetNamePrefix As String
    SheetNamePrefix = GetSheetRefForRangeReference(DataSource.Worksheet.Name, True)

    ' Creating a normal range reference
    Dim NormalRangeRef As String
    NormalRangeRef = EQUAL_SIGN & SheetNamePrefix & Replace( _
                     Address _
                     , LIST_SEPARATOR _
                      , LIST_SEPARATOR & SheetNamePrefix _
                       )                         'Replacing comma for non-contiguous range.

    ' If the data source is a dynamic formula
    If IsDynamicFormula Then
        ConvertToReference = EQUAL_SIGN & SheetNamePrefix _
                             & DataSource.SpillParent.Address(IsAbsoluteRow, IsAbsoluteColumn) & HASH_SIGN
        ' If structured reference check is needed
    ElseIf IsCheckForStructuredReference Then
        ' Check if the data source is inside a table
        If IsInsideTable(DataSource) Then
            ConvertToReference = GetReferenceForTable(DataSource, SelectionRange, NormalRangeRef)
            ' Check if the data source is inside a named range
        ElseIf IsInsideNamedRange(DataSource) Then
            ConvertToReference = GetReferenceForNamedRange(DataSource _
                                                           , SelectionRange _
                                                            , NormalRangeRef, IsAbsoluteRow, IsAbsoluteColumn)
        Else
            ConvertToReference = NormalRangeRef
        End If
    Else
        ConvertToReference = NormalRangeRef
    End If

    ' End of function, logging
    Logger.Log TRACE_LOG, "Exit modNamedRange.ConvertToReference"

End Function

' @Helper For ConvertToReference
Private Function GetReferenceForTable(ByVal DataSource As Range _
                                      , ByVal SelectionRange As Range _
                                       , ByVal NormalRangeRef As String) As String

    ' Find the Excel table (ListObject) that contains the DataSource range
    Dim Table As ListObject
    Set Table = GetTableFromRange(DataSource)

    ' Find the intersection of the table's body with the SelectionRange
    Dim Temp As Range
    Set Temp = FindIntersection(Table.DataBodyRange, SelectionRange)

    ' If there is no intersection
    If IsNothing(Temp) Then
        ' Return the normal range reference
        GetReferenceForTable = NormalRangeRef
        ' If the count of rows in the intersection is equal to the count of rows in the table's body
    ElseIf Temp.Rows.Count = Table.DataBodyRange.Rows.Count Then
        ' Return a reference to the entire data body of the table
        GetReferenceForTable = EQUAL_SIGN & Table.Name & ConvertDataBodyReference(Table, DataSource)
        ' In other cases
    Else
        ' Return the normal range reference
        GetReferenceForTable = NormalRangeRef
    End If

End Function

' @Helper For ConvertToReference
Private Function GetReferenceForNamedRange(ByVal DataSource As Range _
                                           , ByVal SelectionRange As Range _
                                            , ByVal NormalRangeRef As String _
                                             , Optional ByVal IsAbsoluteRow As Boolean = True _
                                              , Optional ByVal IsAbsoluteColumn As Boolean = True) As String
    ' Find the Named Range that contains the DataSource cell
    Dim CurrentName As Name
    Set CurrentName = FindNamedRangeFromSubCell(DataSource)

    ' Check if it is valid to use structured reference for the SelectionRange
    Dim Index As Long
    If IsValidToUseStructuredRef(SelectionRange, CurrentName) Then
        ' If the address of the SelectionRange equals the address of the range referred by the Named Range
        If SelectionRange.Address = CurrentName.RefersToRange.Address Then
            ' Return the normal range reference
            GetReferenceForNamedRange = NormalRangeRef
        Else
            ' Calculate the offset in columns from the first cell of the range referred by the Named Range to the DataSource cell
            Index = DataSource.Column - CurrentName.RefersToRange.Cells(1, 1).Column
            ' Return a string representing a range offset from the DataSource cell, with the same number of rows as the Named Range, and the same number of columns as the SelectionRange
            GetReferenceForNamedRange = EQUAL_SIGN & OFFSET_FX_NAME & FIRST_PARENTHESIS_OPEN _
                                        & CurrentName.RefersToRange.Offset(0, Index).Cells(1, 1).Address(IsAbsoluteRow, IsAbsoluteColumn) _
                                        & LIST_SEPARATOR & "0" & LIST_SEPARATOR _
                                        & "0" & LIST_SEPARATOR & ROWS_FX_NAME & FIRST_PARENTHESIS_OPEN _
                                        & CurrentName.Name & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                                        & SelectionRange.Columns.Count & FIRST_PARENTHESIS_CLOSE
        End If
        ' If it is not valid to use structured reference
    Else
        ' Return the normal range reference
        GetReferenceForNamedRange = NormalRangeRef
    End If

End Function

Private Function IsValidToUseStructuredRef(ByVal SelectionRange As Range, ByVal CurrentName As Name) As Boolean

    IsValidToUseStructuredRef = ((FindIntersection(CurrentName.RefersToRange _
                                                   , SelectionRange).Address = SelectionRange.Address) _
                                 And (SelectionRange.Rows.Count = CurrentName.RefersToRange.Rows.Count))

End Function

Public Function MakeValidDefinedName(ByVal GivenDefinedName As String _
                                     , ByVal JustRemoveInvalidChars As Boolean _
                                      , Optional ByVal IsFinal As Boolean) As String

    ' Logging entry into function
    Logger.Log TRACE_LOG, "Enter modNamedRange.MakeValidDefinedName"

    ' Check if the GivenDefinedName is an empty or blank string
    If Trim$(GivenDefinedName) = vbNullString Then
        ' If IsFinal is True, assign "_Blank" to MakeValidDefinedName, this is used when the input GivenDefinedName is blank and we are finalizing the name
        If IsFinal Then MakeValidDefinedName = "_Blank"
        ' Log the exit due to the Exit Function statement
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.MakeValidDefinedName"
        ' Exit the function when GivenDefinedName is blank
        Exit Function
    End If

    MakeValidDefinedName = MakeValidName(GivenDefinedName, JustRemoveInvalidChars)

    ' Logging exit from function
    Logger.Log TRACE_LOG, "Exit modNamedRange.MakeValidDefinedName"

End Function

Private Function IsValidRangeToName(ByVal GivenRange As Range _
                                    , Optional ByVal IsShowPopup As Boolean = True _
                                     , Optional ByVal IsLocal As Boolean = False) As Boolean

    ' Logging entry into function
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsValidRangeToName"

    ' Assign the result of check if the given range has a defined name to IsValidRangeToName
    IsValidRangeToName = (Not IsRangeHasDefinedName(GivenRange))

    ' If the range already has a defined name
    If Not IsValidRangeToName Then
        ' Obtain the name of the given range
        Dim CurrentName As Name
        Set CurrentName = GivenRange.Name
        ' Check if the scope of the name matches with the expected scope (Local or not)
        IsValidRangeToName = Not (IsLocalScopeNamedRange(CurrentName.NameLocal) = IsLocal)
    End If

    ' If after checking the range is still not valid to be named
    If Not IsValidRangeToName Then
        ' Prepare a message to be displayed
        Dim Message As String
        Message = "Unable to name range " & GivenRange.Address(False, False) & ". The range is already named " _
                  & GivenRange.Name.Name & "."
        ' Show a popup with the message if IsShowPopup is True
        If IsShowPopup Then MsgBox Message, vbExclamation + vbOKOnly, APP_NAME
    End If

    ' Logging exit from function
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsValidRangeToName"

End Function

Private Function IsRangeHasDefinedName(ByVal GivenRange As Range) As Boolean

    ' Logging the start of function
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsRangeHasDefinedName"

    ' Error handling block to identify if GivenRange has a defined name
    On Error GoTo NoNameRangeFound
    If IsNotNothing(GivenRange.Name) Then
        ' If GivenRange has a defined name, return True
        IsRangeHasDefinedName = True
    End If

ExitHandler:
    On Error GoTo 0
    ' Logging function exit
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.IsRangeHasDefinedName"
    Exit Function

NoNameRangeFound:
    Select Case Err.Number
        Case 1004
            ' If error number 1004 (No defined name), return False
            IsRangeHasDefinedName = False
            Resume ExitHandler
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select

    ' Logging function exit
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsRangeHasDefinedName"

End Function

Private Function IsInsideTableHeader(ByVal GivenRange As Range) As Boolean

    ' Logging the start of function
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsInsideTableHeader"

    Dim ActiveTable As ListObject
    Set ActiveTable = GetTableFromRange(GivenRange)
    If IsNothing(ActiveTable) Then
        IsInsideTableHeader = False
    ElseIf IsNoIntersection(ActiveTable.HeaderRowRange, GivenRange) Then
        IsInsideTableHeader = False
    Else
        ' If GivenRange intersects with the HeaderRowRange, return True
        IsInsideTableHeader = True
    End If

    ' Logging function exit
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsInsideTableHeader"

End Function

Public Function FindNamedRange(ByVal FromWorkbook As Workbook, ByVal NameOfTheNamedRange As String) As Name

    ' Logs entering the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindNamedRange"

    ' Iterates over all names in the workbook
    Dim CurrentNamedRange As Name
    For Each CurrentNamedRange In FromWorkbook.Names
        Logger.Log DEBUG_LOG, CurrentNamedRange.Name
        ' If the current name is a local scope name and matches the target name (case insensitive), return it
        If IsLocalScopeNamedRange(CurrentNamedRange.NameLocal) Then
            If VBA.UCase$(NameOfTheNamedRange) = VBA.UCase$(ExtractNameFromLocalNameRange(CurrentNamedRange.NameLocal)) Then
                Set FindNamedRange = CurrentNamedRange
                ' Logs exiting the function due to finding the target name
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.FindNamedRange"
                Exit Function
            End If

            ' If the current name matches the target name (case insensitive), return it
        ElseIf VBA.UCase$(NameOfTheNamedRange) = VBA.UCase$(CurrentNamedRange.Name) Then
            Set FindNamedRange = CurrentNamedRange
            ' Logs exiting the function due to finding the target name
            Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.FindNamedRange"
            Exit Function
        End If
    Next CurrentNamedRange

    ' Logs exiting the function without finding the target name
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindNamedRange"

End Function

Private Function IsNamePresent(ByVal FromWorkbook As Workbook, ByVal NameOfTheNamedRange As String) As Boolean

    ' Iterates over all names in the workbook
    Dim CurrentNamedRange As Name
    For Each CurrentNamedRange In FromWorkbook.Names
        ' If the current name matches the target name, return True
        If CurrentNamedRange.NameLocal = NameOfTheNamedRange Then
            IsNamePresent = True
            Exit Function
        End If
    Next CurrentNamedRange

    ' If no name matches the target name, return False
    IsNamePresent = False

End Function
