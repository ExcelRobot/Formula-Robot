Attribute VB_Name = "modNamedRange"
'@IgnoreModule UndeclaredVariable, AssignmentNotUsed
'@Folder "NamedRange"
' @Folder "NamedRange.Driver"
' @IgnoreModule SuperfluousAnnotationArgument, UnrecognizedAnnotation, ProcedureNotUsed
Option Explicit
Option Private Module

Private Enum LabelSourceOnNameParameterCells
    ONLY_ROW = 1
    ROW_COLUMN = 2
    COLUMN_ROW = 3
End Enum

Private Enum RangeType
    ROW_VECTOR = 1
    COLUMN_VECTOR = 2
    ARRAY_2D = 3
    SINGLE_CELL = 4
End Enum

Public Sub SaveNamedRange(ByVal ForCell As Range, Optional ByVal IgnorePrefix As String = "[")

    ' Exit the subroutine if the cell doesn't have a formula
    If Not ForCell.HasFormula Then Exit Sub
    
    Dim DefaultName As String
    DefaultName = GetOldNameFromComment(ForCell, NAMED_RANGE_NOTE_PREFIX)

    ' If no old name was retrieved, find and process a new default name
    If DefaultName = vbNullString Then
        DefaultName = FindDefaultName(ForCell, IgnorePrefix)
        DefaultName = GetFinalDefineName(DefaultName, True)
    End If
    
    ' Exit the subroutine if the final default name is not valid
    If DefaultName = vbNullString Then Exit Sub

    Dim Reference As String
    Reference = ReplaceAllDependencyWithAbsAddress(ForCell.Formula2, ForCell.Worksheet)

    ' If the formula starts with a lambda function then exit the subroutine
    If IsLambdaFunction(ForCell.Formula2) Then
        Exit Sub
    Else
        Dim CurrentName As Name
        
        ' Create Named Range
        Set CurrentName = CreateNamedRange(Reference, DefaultName, False, Nothing, ForCell.Parent)
        
        ' Apply the Named Range to chart and formulas
        ApplyNameRangeToChartAndFormula DefaultName, CurrentName, ForCell.Worksheet.Parent

        ' Replace the cell's formula with a reference to the Named Range
        ForCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(EQUAL_SIGN & CurrentName.Name)
    End If

    ' Remove the previous comment from the cell
    DeleteComment ForCell
    
End Sub

Private Function ReplaceAllDependencyWithAbsAddress(ByVal Formula As String, FormulaInSheet As Worksheet) As String
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = modCOMWrapper.GetDirectPrecedents(Formula, FormulaInSheet)
    
    Dim Result As String
    If Not IsArray(DirectPrecedents) Then
        Result = Formula
    Else
        Result = Formula
        Dim CurrentPrecedent As Variant
        For Each CurrentPrecedent In DirectPrecedents
            If CurrentPrecedent <> vbNullString Then
                Dim AbsRef As String
                AbsRef = MakeAbsoluteReference(CStr(CurrentPrecedent), FormulaInSheet.Cells(1))
                Result = ReplaceTokenWithNewToken(Result, CStr(CurrentPrecedent), AbsRef)
            End If
        Next CurrentPrecedent
        
    End If
    
    ReplaceAllDependencyWithAbsAddress = Result
    
End Function

Public Sub EditNamedRange(ByVal ForCell As Range)

    ' Exit the subroutine if the cell doesn't have a formula
    If Not ForCell.HasFormula Then Exit Sub

    Dim CurrentName As Name

    ' Try to get the current name from the cell's formula
    
    Dim PossibleNamedRangeName As String
    PossibleNamedRangeName = Text.AfterDelimiter(ForCell.Formula2, EQUAL_SIGN)
    
    ' If local scoped named range is sheet qualified then Workbook will find that.
    If IsNamedRangeExist(ForCell.Worksheet.Parent, PossibleNamedRangeName) Then
        Set CurrentName = ForCell.Worksheet.Parent.Names(PossibleNamedRangeName)
    ElseIf IsLocalScopedNamedRangeExist(ForCell.Worksheet, PossibleNamedRangeName) Then
        Set CurrentName = ForCell.Worksheet.Names(PossibleNamedRangeName)
    Else
        Exit Sub
    End If
    
    If IsLambdaFunction(CurrentName.RefersTo) Then Exit Sub
        
    ' Replace the cell's formula with the formula of the Named Range
    ForCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(CurrentName.RefersTo)
    ' Update or add a comment with the Named Range's name
    modUtility.UpdateOrAddNamedRangeNameNote ForCell, CurrentName.Name, NAMED_RANGE_NOTE_PREFIX

End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Cancel Named Range Edit
' Description:            Cancel edit named range edit mode and use named range name instead of it's refers to and delete the comment.
' Macro Expression:       modNamedRange.CancelNamedRangeEdit([ActiveCell])
' Generated:              09/19/2024 11:51 PM
'----------------------------------------------------------------------------------------------------
Public Sub CancelNamedRangeEdit(ByVal OnCell As Range)
    
    Dim CurrentComment As Comment
    Set CurrentComment = OnCell.Comment
    
    If CurrentComment Is Nothing Then Exit Sub
    
    If Not Text.IsStartsWith(CurrentComment.Text, NAMED_RANGE_NOTE_PREFIX) Then Exit Sub
    
    Dim NamedRangeName As String
    NamedRangeName = Text.AfterDelimiter(CurrentComment.Text, NAMED_RANGE_NOTE_PREFIX)
    
    If IsNamedRangeExist(OnCell.Worksheet.Parent, NamedRangeName) Then
        DeleteComment OnCell
        OnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(EQUAL_SIGN & NamedRangeName)
    ElseIf IsLocalScopedNamedRangeExist(OnCell.Worksheet, NamedRangeName) Then
        DeleteComment OnCell
        OnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(EQUAL_SIGN & NamedRangeName)
    End If
    
End Sub

Public Sub ReAssignLocalNamedRange(ByVal SelectionRange As Range)
    AddNameRange SelectionRange, False, True, IsReassign:=True
End Sub

Public Sub ReAssignGlobalNamedRange(ByVal SelectionRange As Range)
    AddNameRange SelectionRange, False, False, IsReassign:=True
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Name Parameter Column
'  Description:            Name Parameter Column.
'  Macro Expression:       modNamedRange.NameParameterColumn([Selection])
'  Generated:              11/03/2022 12:42 PM
' ----------------------------------------------------------------------------------------------------
Public Sub NameParameterColumn(ByVal SelectionRange As Range)
    
    Dim CurrentColumnRange As Range

    ' Loop through each column in the selected range
    For Each CurrentColumnRange In SelectionRange.Columns
        Dim ProbableNameCell As Range
        
        ' Find the first visible cell in the current column
        Set ProbableNameCell = FindFirstNonHiddenCell(CurrentColumnRange.Cells(1), -1, 0, "[")
        
        ' Loop until we find a valid cell for name or reach the first row
        Do While Not (IsValidCellForName(ProbableNameCell) Or ProbableNameCell.Row = 1)
            Set ProbableNameCell = FindFirstNonHiddenCell(ProbableNameCell, -1, 0, "[")
        Loop
        
        ' If the cell has value, apply it as a name to the current column range
        If GetCellValueIfErrorNullString(ProbableNameCell) <> vbNullString Then
            ApplyNameRange CurrentColumnRange, ProbableNameCell.Value, False, IsCheckForStructuredReference:=True
        End If
    Next CurrentColumnRange

End Sub

Public Sub NameParameterRow(ByVal SelectionRange As Range)
    
    Dim CurrentRowRange As Range

    ' Loop through each row in the selected range
    For Each CurrentRowRange In SelectionRange.Rows
        Dim ProbableNameCell As Range
        
        ' Find the first visible cell in the current row
        Set ProbableNameCell = FindFirstNonHiddenCell(CurrentRowRange.Cells(1), 0, -1, "[")
        
        ' Loop until we find a valid cell for name or reach the first column
        Do While Not (IsValidCellForName(ProbableNameCell) Or ProbableNameCell.Column = 1)
            Set ProbableNameCell = FindFirstNonHiddenCell(ProbableNameCell, 0, -1, "[")
        Loop
        
        ' If the cell has value, apply it as a name to the current row range
        If GetCellValueIfErrorNullString(ProbableNameCell) <> vbNullString Then
            ApplyNameRange CurrentRowRange, ProbableNameCell.Value, False, IsCheckForStructuredReference:=True
        End If
    Next CurrentRowRange

End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Create Relative Column Named Range
'  Description:            Create relative column named range.
'  Macro Expression:       modNamedRange.CreateRelativeColumnNamedRange()
'  Generated:              10/31/2022 09:43 PM
' ----------------------------------------------------------------------------------------------------
Public Sub CreateRelativeColumnNamedRange()
    AddNameRange ActiveCell, False, True, ActiveSheet, False, True
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Create Relative Row Named Range
'  Description:            Create relative row named range.
'  Macro Expression:       modNamedRange.CreateRelativeRowNamedRange()
'  Generated:              10/31/2022 09:44 PM
' ----------------------------------------------------------------------------------------------------
Public Sub CreateRelativeRowNamedRange()
    AddNameRange ActiveCell, False, True, ActiveSheet, True, False
End Sub

' @Description("This will delete Named range but it will not update formula with cell references.")
' @Dependency("No Dependency")
' @ExampleCall : DeleteNamedRangeOnly Selection
' @Date : 05 April 2022 05:49:33 AM
' @PossibleError :
Public Sub DeleteNamedRangeOnly(ByVal SelectionRange As Range)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.DeleteNamedRangeOnly"
    Dim CurrentRange As Range

    ' Loop through each area in the selected range
    For Each CurrentRange In SelectionRange.Areas
        Dim CurrentCell As Range

        ' Loop through each cell in the current area
        For Each CurrentCell In CurrentRange.Cells
            Dim CurrentName As Name

            ' Find if the current cell has a named range associated with it
            Set CurrentName = FindNamedRangeFromSubCell(CurrentCell)

            ' If a named range exists, delete it
            If IsNotNothing(CurrentName) Then CurrentName.Delete
        Next CurrentCell
    Next CurrentRange

    Logger.Log TRACE_LOG, "Exit modNamedRange.DeleteNamedRangeOnly"
    
End Sub

Public Sub AddNameToParameterCells(ByVal ParametersRange As Range, Optional ByVal IsLocal As Boolean _
                                                                  , Optional ByVal IgnorePrefix As String = "[" _
                                                                   , Optional ByVal ScopeSheet As Worksheet)
    
    ' Log the entrance into the subroutine
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddNameToParameterCells"
    
    ' Start the timer for performance tracking
    Dim StartTime As Double
    StartTime = Timer
    
    Dim CurrentRange As Range

    ' Loop through each area in the ParametersRange
    For Each CurrentRange In ParametersRange.Areas
        
        Dim TempCells As Collection
        
        ' Check if the CurrentRange contains merged cells
        If HasMergeCells(CurrentRange) Then
            ' If so, split the cells and store them in TempCells
            Set TempCells = SplitCells(CurrentRange)
        Else
            ' If not, simply add the CurrentRange to TempCells
            Set TempCells = New Collection
            TempCells.Add CurrentRange, CurrentRange.Address
        End If
        
        Dim CurrentItem As Range

        ' Loop through each cell in TempCells
        For Each CurrentItem In TempCells
            ' Apply names to the cells according to the rules specified in the ONLY_ROW LabelSourceOnNameParameterCells case
            NameParameterCellsAllCase CurrentItem, LabelSourceOnNameParameterCells.ONLY_ROW _
                                                  , IsLocal, ScopeSheet, IgnorePrefix
        Next CurrentItem
        
        ' Clean up TempCells for the next iteration
        Set TempCells = Nothing
        
    Next CurrentRange

    ' Print the total execution time for performance tracking
    Logger.Log DEBUG_LOG, "Total Time:" & Timer - StartTime

    ' Log the exit from the subroutine
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddNameToParameterCells"
    
End Sub

Public Sub AddNameToParameterCellsByRowColumn(ByVal ParametersRange As Range _
                                              , Optional ByVal IsLocal As Boolean _
                                               , Optional ByVal IgnorePrefix As String = "[" _
                                                , Optional ByVal ScopeSheet As Worksheet)

    ' Log the entrance into the subroutine
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddNameToParameterCellsByRowColumn"

    Dim CurrentRange As Range

    ' Loop through each area in the ParametersRange
    For Each CurrentRange In ParametersRange.Areas

        ' Apply names to the cells according to the rules specified in the ROW_COLUMN case
        NameParameterCellsAllCase CurrentRange, ROW_COLUMN, IsLocal, ScopeSheet, IgnorePrefix

    Next CurrentRange

    ' Log the exit from the subroutine
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddNameToParameterCellsByRowColumn"
    
End Sub

Public Sub AddNameToParameterCellsByColumnRow(ByVal ParametersRange As Range _
                                              , Optional ByVal IsLocal As Boolean _
                                               , Optional ByVal IgnorePrefix As String = "[" _
                                                , Optional ByVal ScopeSheet As Worksheet)

    ' Log the entrance into the subroutine
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddNameToParameterCellsByColumnRow"

    Dim CurrentRange As Range

    ' Loop through each area in the ParametersRange
    For Each CurrentRange In ParametersRange.Areas

        ' Apply names to the cells according to the rules specified in the COLUMN_ROW case
        NameParameterCellsAllCase CurrentRange, COLUMN_ROW, IsLocal, ScopeSheet, IgnorePrefix

    Next CurrentRange

    ' Log the exit from the subroutine
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddNameToParameterCellsByColumnRow"
    
End Sub

Private Function FindProbableDefaultNameCell(ByVal StartFromCell As Range _
                                             , ByVal RowOffset As Long _
                                              , ByVal ColumnOffset As Long _
                                               , ByVal IgnorePrefix As String) As Range
    ' Begin logging for the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindProbableDefaultNameCell"
    
    ' Set the cell to start from
    Set StartFromCell = StartFromCell.Cells(1, 1)
    
    Dim CurrentCell As Range
    ' Check conditions and set the CurrentCell accordingly
    If RowOffset < 0 And StartFromCell.Row > Abs(RowOffset) Then
        Set CurrentCell = StartFromCell.Offset(RowOffset, 0)
        RowOffset = -1
    End If
    If ColumnOffset < 0 And StartFromCell.Column > Abs(ColumnOffset) Then
        Set CurrentCell = StartFromCell.Offset(0, ColumnOffset)
        ColumnOffset = -1
    End If
    
    ' If no CurrentCell is set, assign StartFromCell
    If IsNothing(CurrentCell) Then Set CurrentCell = StartFromCell
    
    ' Start a loop that will go through cells until specific conditions are met
    Do While True
        ' Check if the cell is hidden
        If IsCellHidden(CurrentCell) Then
            If CurrentCell.Column > Abs(ColumnOffset) And CurrentCell.Row > Abs(RowOffset) Then
                Set CurrentCell = CurrentCell.Offset(RowOffset, ColumnOffset)
            Else
                Set CurrentCell = Nothing
                Exit Do
            End If
            
            ' Logging the current state of the cell
            Logger.Log DEBUG_LOG, CurrentCell.Address & " is hidden. So we will try next one."
            
            ' Checking if the cell value has errors or matches specific conditions
        ElseIf IsError(CurrentCell.Value) Or _
               IsStartedWithGivenPrefix(CurrentCell.Value, IgnorePrefix) Or _
               CurrentCell.Value = vbNullString Then
                                          
            If CurrentCell.Column > Abs(ColumnOffset) And CurrentCell.Row > Abs(RowOffset) Then
                Set CurrentCell = CurrentCell.Offset(RowOffset, ColumnOffset)
            Else
                Set CurrentCell = Nothing
                Exit Do
            End If
            
            ' Logging the current state of the cell
            Logger.Log DEBUG_LOG, CurrentCell.Address & " is blank. So we will try next one."
            
        Else
            Exit Do
        End If
    Loop
    
    ' Set the result to CurrentCell and log the exit of the function
    Set FindProbableDefaultNameCell = CurrentCell
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindProbableDefaultNameCell"
    
End Function

Private Function IsStartedWithGivenPrefix(ByVal CellValue As Variant, ByVal IgnorePrefix As String) As Boolean
    
    If IsError(CellValue) Then
        IsStartedWithGivenPrefix = False
    Else
        IsStartedWithGivenPrefix = Text.IsStartsWith(CStr(CellValue), IgnorePrefix)
    End If
    
End Function

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

Private Function IsValidRangeAndNoConflictingName(ByRef SelectionRange As Range _
                                                  , ByRef NameOfNamedRange As String _
                                                   , ByVal IsLocal As Boolean) As Boolean
    
    ' Get the first cell reference if merged and check if it's a valid range to name
    Set SelectionRange = GetFirstCellRefIfMerged(SelectionRange)
    If Not IsValidRangeToName(SelectionRange, False, IsLocal) Then Exit Function
    
    ' Make a valid defined name for the named range
    NameOfNamedRange = MakeValidDefinedName(NameOfNamedRange, True)
    
    ' Find the existing named range
    Dim CurrentName As Name
    Set CurrentName = FindNamedRange(SelectionRange.Worksheet.Parent, NameOfNamedRange)
    
    ' Check if the name conflicts with an existing named range
    If IsConflictName(SelectionRange, IsLocal, CurrentName, NameOfNamedRange) Then
        ' If a name conflict arises, show a pop-up message to the user
        MsgBox "Unable to apply name " & NameOfNamedRange & ". A range already exists with this name." _
               , vbExclamation + vbOKOnly, APP_NAME
        Exit Function
    End If
    
    ' If no conflicts and range is valid, return True
    IsValidRangeAndNoConflictingName = True
    
End Function

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

Public Function GetNameFromBook(ByVal FromWorkbook As Workbook _
                                 , ByVal NameOfNamedRange As String) As Name
    
    On Error Resume Next
    Dim Result As Name
    Set Result = FromWorkbook.Names(NameOfNamedRange)
    If IsNotNothing(Result) Then
        ' If we have both local and global scoped named range in a workbook with same name
        ' then it may return the local version. In that case Result.Name will have sheet name prefix
        ' as well and that's why it won't be the same.
        If Result.Name <> NameOfNamedRange Then
            Dim CurrentName As Name
            For Each CurrentName In FromWorkbook.Names
                If CurrentName.Name = NameOfNamedRange Then
                    Set Result = CurrentName
                    Exit For
                End If
            Next CurrentName
        End If
        
    End If
    On Error GoTo 0
    
    Set GetNameFromBook = Result
    
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

Public Function ConvertFromA1ToR1C1(ByVal GivenFormula As String) As String
    
    On Error GoTo ErrorHandler
    ConvertFromA1ToR1C1 = Application.ConvertFormula(GivenFormula, xlA1, xlR1C1)
    Exit Function
    
ErrorHandler:
    ConvertFromA1ToR1C1 = GivenFormula
    
End Function

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

Public Function IsRefersToRangeIsNothing(ByVal CurrentName As Name) As Boolean
    
    On Error GoTo ErrorHandler
    IsRefersToRangeIsNothing = IsNothing(CurrentName.RefersToRange)
    Exit Function
    
ErrorHandler:
    IsRefersToRangeIsNothing = True
    
End Function

Private Function IsAlreadyANameExist(ByVal NameOfNamedRange As String, ByVal SelectionRange As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsAlreadyANameExist"
    Dim CurrentName As Name
    Set CurrentName = FindNamedRange(SelectionRange.Worksheet.Parent, NameOfNamedRange)
    If IsNotNothing(CurrentName) Then
        IsAlreadyANameExist = (CurrentName.Name = NameOfNamedRange)
    End If
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsAlreadyANameExist"
    
End Function

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

Public Sub RenameNamedRange(ByVal SelectionRange As Range)
    
    ' Logging the start of the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.RenameNamedRange"
    
    ' If the selection does not have a defined name, log a debug message and exit the sub
    If Not IsRangeHasDefinedName(SelectionRange) Then
        Logger.Log DEBUG_LOG, SelectionRange.Address & " doesn't have any named range"
        Exit Sub
    End If
    
    ' Store the current name of the selection
    Dim OldName As String
    OldName = SelectionRange.Name.Name
    ' Separate the SheetNamePrefix and the OldName if SHEET_NAME_SEPARATOR is found in the OldName
    Dim SheetNamePrefix As String
    SheetNamePrefix = Text.BeforeDelimiter(OldName, SHEET_NAME_SEPARATOR, , FROM_END)
    If Text.Contains(OldName, SHEET_NAME_SEPARATOR) Then
        OldName = Text.AfterDelimiter(OldName, SHEET_NAME_SEPARATOR, , FROM_END)
    End If
    Dim NewName As String
    ' Prompt user for new name
    NewName = InputBox("Enter new name:", "Named Range Name", OldName)
    
    ' If the user inputs nothing or cancels the prompt, log a trace message and exit the sub
    If NewName = vbNullString Or NewName = "False" Then
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.RenameNamedRange"
        Exit Sub
    End If

    ' Make the user's input into a valid defined name
    NewName = MakeValidDefinedName(NewName, True)
    ' Create the full name for the range including the sheet name prefix if any
    Dim QualifiedSheetName As String
    QualifiedSheetName = IIf(SheetNamePrefix = vbNullString _
                             , NewName, SheetNamePrefix & SHEET_NAME_SEPARATOR & NewName)
    
    Dim Message As String
    ' If the new name already exists for the range, show a warning message, log the error and exit the sub
    If IsAlreadyANameExist(QualifiedSheetName, SelectionRange) Then
        Message = "Unable to rename range " & OldName & ".  The name " & NewName & " already exists."
        Logger.Log DEBUG_LOG, Message
        MsgBox Message, vbExclamation + vbOKOnly, APP_NAME
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.RenameNamedRange"
        Exit Sub
    Else
        ' Assign the new name to the selected range
        Dim ChartNamedRangeUpdater As RangeDependencyInChart
        Set ChartNamedRangeUpdater = New RangeDependencyInChart
        ChartNamedRangeUpdater.ExtractAllNamedRangeRef SelectionRange.Name, SelectionRange.Worksheet.Parent
        SelectionRange.Name.Name = NewName
        ChartNamedRangeUpdater.RenameNamedRange NewName
        Set ChartNamedRangeUpdater = Nothing
    End If
    
    ' Logging the end of the function
    Logger.Log TRACE_LOG, "Exit modNamedRange.RenameNamedRange"
    
End Sub

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

Private Function IsWholeColumnDataBodySelected(ByVal GivenRange As Range) As Boolean
    
    ' Logging the start of function
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsWholeColumnDataBodySelected"
    
    ' Check if GivenRange is inside a table
    If IsInsideTable(GivenRange) Then
        Dim ActiveTable As ListObject
        Set ActiveTable = GetTableFromRange(GivenRange)
        Dim CurrentListColumn As ListColumn
        For Each CurrentListColumn In ActiveTable
            If GivenRange.Address = CurrentListColumn.DataBodyRange.Address Then
                ' If GivenRange is a whole column in the table's body, return True
                IsWholeColumnDataBodySelected = True
                ' Logging function exit due to keyword
                Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.IsWholeColumnDataBodySelected"
                Exit Function
            End If
        Next CurrentListColumn
    End If
    ' Logging function exit
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsWholeColumnDataBodySelected"
    
End Function

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Convert Local Named Ranges To Global
'  Description:            Converts all locally scoped named ranges in selection to global.
'  Macro Expression:       modNamedRange.ConvertLocalToGlobal([Selection])
'  Generated:              03/29/2022 07:52
' ----------------------------------------------------------------------------------------------------
Public Sub ConvertLocalToGlobal(ByVal SelectionRange As Range)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.ConvertLocalToGlobal"
    Dim CurrentCell As Range
    For Each CurrentCell In SelectionRange
        If IsValidToConvertFromLocalToGlobal(CurrentCell, SelectionRange) Then
            ConvertLocalNamedRangeToGlobal FindNamedRangeFromSubCell(CurrentCell)
        End If
    Next CurrentCell
    Logger.Log TRACE_LOG, "Exit modNamedRange.ConvertLocalToGlobal"
    
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Convert Global Named Ranges To Local
'  Description:            This will convert all global named range from selction to local named ranges.
'  Macro Expression:       modNamedRange.ConvertGlobalToLocal([Selection])
'  Generated:              03/08/2023 05:52 PM
' ----------------------------------------------------------------------------------------------------
Public Sub ConvertGlobalToLocal(ByVal SelectionRange As Range)
    
    Dim CurrentCell As Range
    For Each CurrentCell In SelectionRange
        If IsValidToConvertFromGlobalToLocal(CurrentCell, SelectionRange) Then
            ConvertGlobalNamedRangeToLocal FindNamedRangeFromSubCell(CurrentCell), CurrentCell
        End If
    Next CurrentCell

End Sub

Private Function IsValidToConvert(ByVal CurrentCell As Range _
                                  , ByVal SelectionRange As Range _
                                   , ByVal FromLocalToGlobal As Boolean) As Boolean
    
    ' Find the named range for the current cell
    Dim CurrentName As Name
    Set CurrentName = FindNamedRangeFromSubCell(CurrentCell)
    
    ' Checking if CurrentName is Nothing, if so, the function returns False
    If IsNothing(CurrentName) Then
        IsValidToConvert = False
        ' Checking if SelectionRange is a subset of the range that CurrentName refers to
    ElseIf IsSubSet(SelectionRange, CurrentName.RefersToRange) Then
        ' Checking if the scope of the named range matches the intended conversion direction
        If IsLocalScopeNamedRange(CurrentName.NameLocal) <> FromLocalToGlobal Then
            IsValidToConvert = False
        Else
            Dim ConvertedName As String
            ' Creating a new name for the conversion process based on whether it is a local to global conversion or vice versa
            If FromLocalToGlobal Then
                ConvertedName = Text.AfterDelimiter(CurrentName.NameLocal, SHEET_NAME_SEPARATOR, , FROM_END)
            Else
                ConvertedName = GetSheetRefForRangeReference(CurrentCell.Worksheet.Name) & CurrentName.Name
            End If
            ' Checking if the converted name already exists in the scope of CurrentCell
            IsValidToConvert = Not (IsAlreadyANameExist(ConvertedName, CurrentCell))
        End If
    Else
        IsValidToConvert = False
    End If
    
End Function

Private Function IsValidToConvertFromGlobalToLocal(ByVal CurrentCell As Range _
                                                   , ByVal SelectionRange As Range) As Boolean
    IsValidToConvertFromGlobalToLocal = IsValidToConvert(CurrentCell, SelectionRange, False)
End Function

Private Function IsValidToConvertFromLocalToGlobal(ByVal CurrentCell As Range _
                                                   , ByVal SelectionRange As Range) As Boolean
    IsValidToConvertFromLocalToGlobal = IsValidToConvert(CurrentCell, SelectionRange, True)
End Function

Private Sub ConvertLocalNamedRangeToGlobal(ByVal GivenName As Name)
    
    ' Logs entering the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.ConvertLocalNamedRangeToGlobal"
    
    ' Extracts name from local name range
    Dim NameOfNamedRange As String
    NameOfNamedRange = ExtractNameFromLocalNameRange(GivenName.NameLocal)
    
    ' Gets the reference range
    Dim ReferToRange As String
    ReferToRange = GivenName.RefersTo
    
    ' Gets the workbook where the name is located
    Dim NameOnWorkbook As Workbook
    Set NameOnWorkbook = GivenName.Parent.Parent
    On Error GoTo ResetBackCalculation
    
    ' Stores the calculation setting and switches to manual calculation to optimize the process
    Dim CalculationType As XlCalculation
    CalculationType = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' Deletes the existing name and adds it to the workbook level (converting it from local to global)
    GivenName.Delete
    NameOnWorkbook.Names.Add NameOfNamedRange, ReferToRange
    
    ' Logs successful conversion
    Logger.Log DEBUG_LOG, NameOfNamedRange & " has been added to " _
                         & NameOnWorkbook.Name & " which refer to " & ReferToRange
    
    ' Restores the original calculation setting
    Application.Calculation = CalculationType
    On Error GoTo 0

Cleanup:
    ' Logs exiting the function
    Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.ConvertLocalNamedRangeToGlobal"
    Exit Sub

ResetBackCalculation:
    
    ' In case of error, restores the original calculation setting and raises an error
    Application.Calculation = CalculationType
    Logger.Log DEBUG_LOG, Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume Cleanup
    Logger.Log TRACE_LOG, "Exit modNamedRange.ConvertLocalNamedRangeToGlobal"
    
End Sub

Private Sub ConvertGlobalNamedRangeToLocal(ByVal GivenName As Name, ByVal FromCell As Range)
    
    ' Starts with the name of the named range and its referred range
    Dim NameOfNamedRange As String
    NameOfNamedRange = GivenName.Name
    Dim ReferToRange As String
    ReferToRange = GivenName.RefersTo
    
    ' The worksheet to which the name will be localized
    Dim NameOnSheet As Worksheet
    Set NameOnSheet = FromCell.Parent
    On Error GoTo ResetBackCalculation
    
    ' Saves the current calculation setting and switches to manual calculation to optimize the process
    Dim CalculationType As XlCalculation
    CalculationType = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    ' Deletes the existing name and adds it to the worksheet level (converting it from global to local)
    GivenName.Delete
    NameOnSheet.Names.Add NameOfNamedRange, ReferToRange
    
    ' Logs the successful conversion
    Logger.Log DEBUG_LOG, NameOfNamedRange & " has been added to " _
                         & NameOnSheet.Name & " which refer to " & ReferToRange
    
    ' Restores the original calculation setting
    Application.Calculation = CalculationType
    On Error GoTo 0
    Set NameOnSheet = Nothing

Cleanup:
    Exit Sub

ResetBackCalculation:
    
    ' In case of error, restores the original calculation setting and raises an error
    Application.Calculation = CalculationType
    Logger.Log DEBUG_LOG, Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume Cleanup
    Resume
    
End Sub

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

Private Function GetRangeType(ByVal GivenRange As Range) As RangeType
    
    ' Logs entering the function
    Logger.Log TRACE_LOG, "Enter modNamedRange.GetRangeType"
    
    ' Checks the type of the given range and returns the appropriate constant
    If GivenRange.Cells(1, 1).MergeArea.Address = GivenRange.Address Then
        GetRangeType = SINGLE_CELL
    ElseIf GivenRange.Rows.Count > 1 And GivenRange.Columns.Count > 1 Then
        GetRangeType = ARRAY_2D
    ElseIf GivenRange.Rows.Count > 1 Then
        GetRangeType = COLUMN_VECTOR
    ElseIf GivenRange.Columns.Count > 1 Then
        GetRangeType = ROW_VECTOR
    Else
        GetRangeType = SINGLE_CELL
    End If
    
    ' Logs exiting the function
    Logger.Log TRACE_LOG, "Exit modNamedRange.GetRangeType"
    
End Function

Public Function FindFirstNonHiddenCell(ByVal StartFromCell As Range _
                                       , ByVal RowOffset As Long _
                                        , ByVal ColumnOffset As Long _
                                         , ByVal IgnorePrefix As String) As Range
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindFirstNonHiddenCell"
    Dim CurrentCell As Range

    ' Find the first non-hidden cell based on the provided offset values
    ' CurrentCell is set based on the row and column offset conditions
    If RowOffset < 0 And StartFromCell.Row > Abs(RowOffset) Then
        Set CurrentCell = StartFromCell.Offset(RowOffset, 0)
        RowOffset = -1
    End If
    If ColumnOffset < 0 And StartFromCell.Column > Abs(ColumnOffset) Then
        Set CurrentCell = StartFromCell.Offset(0, ColumnOffset)
        ColumnOffset = -1
    End If
    
    If IsNothing(CurrentCell) Then Set CurrentCell = StartFromCell
    
    ' Loop to continue until a non-hidden cell that does not start with the provided prefix is found
    Do While True
        If IsCellHidden(CurrentCell) Or IsStartedWithGivenPrefix(CurrentCell.Value, IgnorePrefix) Then
            If CurrentCell.Column > Abs(ColumnOffset) And CurrentCell.Row > Abs(RowOffset) Then
                Set CurrentCell = CurrentCell.Offset(RowOffset, ColumnOffset)
            Else
                Set CurrentCell = Nothing
                Exit Do
            End If
            Logger.Log DEBUG_LOG, CurrentCell.Address & " is hidden. So we will try next one."
        Else
            Exit Do
        End If
    Loop
    
    Set FindFirstNonHiddenCell = CurrentCell
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindFirstNonHiddenCell"

End Function

Private Sub NameParameterCellsAllCase(ByVal ParametersRange As Range _
                                      , ByVal LabelOption As LabelSourceOnNameParameterCells _
                                       , Optional ByVal IsLocal As Boolean _
                                        , Optional ByVal ScopeSheet As Worksheet _
                                         , Optional ByVal IgnorePrefix As String = "[")
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.NameParameterCellsAllCase"
    Dim TypeOfRange  As RangeType
    TypeOfRange = GetRangeType(ParametersRange)
    If LabelOption = ONLY_ROW And TypeOfRange = ARRAY_2D Then LabelOption = ROW_COLUMN

    ' Selecting the right operation based on LabelOption
    Select Case LabelOption
        Case LabelSourceOnNameParameterCells.ONLY_ROW
            HandleNameParameterCells ParametersRange, TypeOfRange, IsLocal, ScopeSheet, IgnorePrefix
        Case LabelSourceOnNameParameterCells.ROW_COLUMN
            HandleNameParameterCellsRow_ColumnOrColumn_Row ParametersRange _
                                                           , ROW_COLUMN, IsLocal, ScopeSheet, IgnorePrefix
        Case LabelSourceOnNameParameterCells.COLUMN_ROW
            HandleNameParameterCellsRow_ColumnOrColumn_Row ParametersRange _
                                                           , COLUMN_ROW, IsLocal, ScopeSheet, IgnorePrefix
        Case Else
            Err.Raise 13, "Case should be only limited to enum member option", "Wrong Case type"
    End Select
    Logger.Log TRACE_LOG, "Exit modNamedRange.NameParameterCellsAllCase"
    
End Sub

Private Sub HandleNameParameterCells(ByVal ParametersRange As Range _
                                     , ByVal TypeOfRange As RangeType _
                                      , Optional ByVal IsLocal As Boolean = False _
                                       , Optional ByVal ScopeSheet As Worksheet _
                                        , Optional ByVal IgnorePrefix As String = "[")
                                                                
    Logger.Log TRACE_LOG, "Enter modNamedRange.HandleNameParameterCells"
    Dim RowOffset As Long
    Dim ColumnOffset As Long

    ' Defining row and column offset based on the type of range
    Select Case TypeOfRange
        Case RangeType.COLUMN_VECTOR, RangeType.SINGLE_CELL
            RowOffset = 0
            ColumnOffset = -1
        Case RangeType.ROW_VECTOR
            ColumnOffset = 0
            RowOffset = -1
    End Select
    
    Dim ProbableCellForName As Range
    Dim CurrentCell As Range

    ' Looping through each cell and finding the probable cell for naming
    For Each CurrentCell In SplitRangeAndConsiderMergeCellsAsOne(ParametersRange)
        If ColumnOffset = -1 Then
            Set ProbableCellForName = FindProbableDefaultNameCell(CurrentCell _
                                                                  , RowOffset, ColumnOffset, IgnorePrefix)
        Else
            Set ProbableCellForName = FindFirstNonHiddenCell(CurrentCell _
                                                             , RowOffset, ColumnOffset, IgnorePrefix)
        End If
        If IsValidCellForName(ProbableCellForName) Then
            ApplyNameRange CurrentCell, ProbableCellForName.Value, False, IsLocal, ScopeSheet
        End If
    Next CurrentCell
    Logger.Log TRACE_LOG, "Exit modNamedRange.HandleNameParameterCells"
    
End Sub

Public Sub TestMergeArea()

    Logger.Log TRACE_LOG, "Enter modNamedRange.TestMergeArea"
    Dim CurrentCell As Range
    Dim SelectionRange As Range
    Set SelectionRange = Selection

    ' Logging the address of each cell
    For Each CurrentCell In SplitRangeAndConsiderMergeCellsAsOne(SelectionRange)
        Logger.Log DEBUG_LOG, CurrentCell.Address
    Next CurrentCell
    Logger.Log TRACE_LOG, "Exit modNamedRange.TestMergeArea"
    
End Sub

Private Function SplitRangeAndConsiderMergeCellsAsOne(ByVal GivenRange As Range) As Collection
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.SplitRangeAndConsiderMergeCellsAsOne"
    
    If IsNothing(GivenRange) Then
        Set SplitRangeAndConsiderMergeCellsAsOne = New Collection
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.SplitRangeAndConsiderMergeCellsAsOne"
        Exit Function
    End If
    
    Dim CellsCollection As Collection
    Set CellsCollection = New Collection
    Dim ResultRange As Range
    
    Do While True
        If IsNothing(ResultRange) Then
            ' Finding intersection of GivenRange and its MergeArea
            Set ResultRange = FindIntersection(GivenRange.Cells(1, 1).MergeArea, GivenRange)
            CellsCollection.Add ResultRange
        ElseIf ResultRange.Address = GivenRange.Address Then
            Exit Do
        Else
            Dim CurrentCell As Range
            For Each CurrentCell In GivenRange.Cells
                Dim Temp As Range
                Set Temp = FindIntersection(ResultRange, CurrentCell)
                If IsNothing(Temp) Then
                    ' Finding intersection of CurrentCell's MergeArea and GivenRange
                    Set Temp = FindIntersection(CurrentCell.MergeArea, GivenRange)
                    Set ResultRange = Union(ResultRange, Temp)
                    CellsCollection.Add Temp
                End If
            Next CurrentCell
        End If
    Loop
    
    Set SplitRangeAndConsiderMergeCellsAsOne = CellsCollection
    Set ResultRange = Nothing
    Logger.Log TRACE_LOG, "Exit modNamedRange.SplitRangeAndConsiderMergeCellsAsOne"
    
End Function

Private Function SplitCells(ByVal GivenRange As Range) As Collection
    
    Dim CellsCollection As Collection
    Set CellsCollection = New Collection
    Dim CurrentRange As Range
    
    For Each CurrentRange In GivenRange.Areas
        
        Dim CurrentCell As Range
        For Each CurrentCell In CurrentRange.Cells
            CellsCollection.Add CurrentCell, CurrentCell.Address
        Next CurrentCell
        
    Next CurrentRange
    
    Dim CurrentItem As Range
    For Each CurrentItem In CellsCollection
        If CurrentItem.MergeCells Then
            ' Checking if merged cell's address is different from address of first cell in merge area
            If CurrentItem.Address <> CurrentItem.MergeArea.Cells(1).Address Then
                CellsCollection.Remove CurrentItem.Address
            End If
        End If
    Next CurrentItem
    
    Set SplitCells = CellsCollection
    
End Function

Public Function HasMergeCells(ByVal GivenRange As Range) As Boolean
    ' Checking if GivenRange has merged cells
    HasMergeCells = (IsNull(GivenRange.MergeCells) Or GivenRange.MergeCells)
End Function

Private Sub HandleNameParameterCellsRow_ColumnOrColumn_Row(ByVal ParametersRange As Range _
                                                           , ByVal LabelOption As LabelSourceOnNameParameterCells _
                                                            , Optional ByVal IsLocal As Boolean _
                                                             , Optional ByVal ScopeSheet As Worksheet _
                                                              , Optional ByVal IgnorePrefix As String = "[")
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.HandleNameParameterCellsRow_ColumnOrColumn_Row"
    
    Dim CurrentColumn As Long
    Dim CellsInCurrentColumn As Range
    Dim ColumnLabel As String
    Dim ProbableColumnLabelCell As Range
    Dim CurrentRow As Long
    Dim CurrentCell As Range
    Dim RowLabelCell As Range
    
    ' Handling cells based on the parameters range's column and row labels
    For CurrentColumn = 1 To ParametersRange.Columns.Count
        Set CellsInCurrentColumn = ParametersRange.Columns(CurrentColumn)
        Set ProbableColumnLabelCell = FindFirstNonHiddenCell(CellsInCurrentColumn.Cells(1, 1), -1, 0, IgnorePrefix)
        If IsValidCellForName(ProbableColumnLabelCell) Then
            ColumnLabel = ProbableColumnLabelCell.Value
            For CurrentRow = 1 To CellsInCurrentColumn.Rows.Count
                Set CurrentCell = CellsInCurrentColumn.Cells(CurrentRow, 1)
                Set RowLabelCell = FindProbableDefaultNameCell(CurrentCell, 0, -1 * CurrentColumn, IgnorePrefix)
                If IsValidCellForName(RowLabelCell) Then
                    Logger.Log DEBUG_LOG, "Cell Ref: " & CurrentCell.Address & " Valid to name it"
                    Dim FinalName As String
                    ' Forming final name based on label option
                    If LabelOption = COLUMN_ROW Then FinalName = ColumnLabel & "_" & RowLabelCell.Value
                    If LabelOption = ROW_COLUMN Then FinalName = RowLabelCell.Value & "_" & ColumnLabel
                    ApplyNameRange CurrentCell, FinalName, False, IsLocal, ScopeSheet
                End If
            Next CurrentRow
        End If
    Next CurrentColumn
    Logger.Log TRACE_LOG, "Exit modNamedRange.HandleNameParameterCellsRow_ColumnOrColumn_Row"
    
End Sub

Private Function IsValidCellForName(ByVal CurrentCell As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsValidCellForName"
    
    If IsNothing(CurrentCell) Then
        IsValidCellForName = False
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.IsValidCellForName"
        Exit Function
    End If
    
    Logger.Log DEBUG_LOG, "Checking validity of name for cell: " & CurrentCell.Address
    ' Checking if cell is valid for naming by checking certain conditions
    IsValidCellForName = Not ((IsNumeric(CurrentCell.Value) And CStr(CurrentCell.Value) <> vbNullString) Or _
                              CurrentCell.HasFormula Or IsInsideNamedRange(CurrentCell) _
                              Or CStr(CurrentCell.Value) = vbNullString)
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsValidCellForName"
    
End Function

'
' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Remove All Named Ranges From Active Workbook
'  Description:            Remove all named ranges From Active Workbook
'  Macro Expression:       modNamedRange_Tests.RemoveAllNamedRanges([ActiveWorkbook])
'  Generated:              03/29/2022 07:18
' ----------------------------------------------------------------------------------------------------

Public Sub RemoveAllNamedRanges(Optional ByVal GivenWorkbook As Workbook = Nothing)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.RemoveAllNameRanges"
    
    ' If GivenWorkbook is not provided, use the active workbook
    If IsNothing(GivenWorkbook) Then Set GivenWorkbook = ActiveWorkbook
    
    ' Remove named ranges from the entire workbook
    RemoveNameRangeFromObject GivenWorkbook

    ' Loop through each sheet in the workbook and remove named ranges from each sheet
    Dim CurrentSheet As Worksheet
    For Each CurrentSheet In GivenWorkbook.Worksheets
        RemoveNameRangeFromObject CurrentSheet
    Next CurrentSheet
    
    Logger.Log TRACE_LOG, "Exit modNamedRange.RemoveAllNameRanges"

End Sub

Private Sub RemoveNameRangeFromObject(ByVal GivenObject As Object)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.RemoveNameRangeFromObject"
    
    ' Loop through each named range in the GivenObject and delete
    Dim CurrentName As Name
    For Each CurrentName In GivenObject.Names
        Logger.Log DEBUG_LOG, CurrentName.Name
        On Error Resume Next
        CurrentName.Delete
    Next CurrentName
    
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modNamedRange.RemoveNameRangeFromObject"
    
End Sub

Public Sub DeleteNamedRangesHavingError(Optional ByVal FromWorkbook As Workbook = Nothing)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.DeleteNamedRangeHavingError"
    
    ' If FromWorkbook is not provided, use the active workbook
    If IsNothing(FromWorkbook) Then Set FromWorkbook = ActiveWorkbook
    
    ' Loop through each named range in the FromWorkbook and delete if it contains error
    Dim CurrentName As Name
    For Each CurrentName In FromWorkbook.Names
        If Text.Contains(CurrentName.RefersTo, REF_ERR_KEYWORD) _
           Or Text.Contains(CurrentName.RefersTo, NAME_ERR_KEYWORD) Then
            On Error Resume Next
            CurrentName.Delete
        End If
    Next CurrentName
    
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modNamedRange.DeleteNamedRangeHavingError"
    
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Name All Table Data Columns
'  Description:            Name all table data columns.
'  Macro Expression:       modNamedRange.NameAllTableDataColumns([ActiveCell])
'  Generated:              07/07/2022 11:11 PM
' ----------------------------------------------------------------------------------------------------
Public Sub NameAllTableDataColumns(ByVal GivenCell As Range)

    Logger.Log TRACE_LOG, "Enter modNamedRange.NameAllTableDataColumns"
    
    Dim NamedRangeCells As Range
    ' Check if the GivenCell is part of a named range or a table and set the range to be named
    If IsCurrentRegionHasNamedRange(GivenCell) Then
        Set NamedRangeCells = GivenCell.CurrentRegion
    ElseIf IsInsideTable(GivenCell) Then
        Set NamedRangeCells = GivenCell.ListObject.DataBodyRange
    Else
        ' Exit subroutine if GivenCell is not part of a named range or a table
        Logger.Log TRACE_LOG, "Exit Due to Exit Keyword modNamedRange.NameAllTableDataColumns"
        Exit Sub
    End If
    
    ' Loop through each column in the range and apply a name to it
    Dim CurrentColumnRange As Range
    For Each CurrentColumnRange In NamedRangeCells.Columns
        NameTableDataColumn CurrentColumnRange, True
    Next CurrentColumnRange
    
    Logger.Log TRACE_LOG, "Exit modNamedRange.NameAllTableDataColumns"
    
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Name Table Data Column
'  Description:            Name table data column.
'  Macro Expression:       modNamedRange.NameTableDataColumn([ActiveCell])
'  Generated:              07/07/2022 11:10 PM
' ----------------------------------------------------------------------------------------------------
Public Sub NameTableDataColumn(ByVal GivenCell As Range, Optional ByVal IsSilentMode As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.NameTableDataColumn"
    
    ' Check if the GivenCell is inside a table, or part of a named range. According to the condition,
    ' different naming functions are called.
    If IsInsideTable(GivenCell) Then
        AddTableColumnNamedRange GivenCell, False, IsSilentMode
    ElseIf IsCurrentRegionHasNamedRange(GivenCell) Then
        AddNamedRangeColumnNamedRange GivenCell, False, IsSilentMode
    End If
    Logger.Log TRACE_LOG, "Exit modNamedRange.NameTableDataColumn"
    
End Sub

Private Function IsCurrentRegionHasNamedRange(ByVal GivenCell As Range) As Boolean
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.IsCurrentRegionHasNamedRange"
    On Error Resume Next
    ' Checking if the current region of GivenCell has a name. If so, returns True
    IsCurrentRegionHasNamedRange = IsNotNothing(GivenCell.CurrentRegion.Name)
    On Error GoTo 0
    Logger.Log TRACE_LOG, "Exit modNamedRange.IsCurrentRegionHasNamedRange"
    
End Function

Private Sub AddNamedRangeColumnNamedRange(ByVal SelectionRange As Range _
                                          , Optional ByVal IsLocal As Boolean = False _
                                           , Optional ByVal IsSilentMode As Boolean = False)
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddNamedRangeColumnNamedRange"
    ' For each area in SelectionRange, and for each column in each area, apply the named range using the first cell of each column
    Dim CurrentRange As Range
    For Each CurrentRange In SelectionRange.Areas
        Dim CurrentCell As Range
        For Each CurrentCell In CurrentRange.Columns
            ApplyNamedRangeColumnNameRangeForCurrentCell CurrentCell.Cells(1, 1), IsLocal, IsSilentMode
        Next CurrentCell
    Next CurrentRange
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddNamedRangeColumnNamedRange"
    
End Sub

Private Sub AddTableColumnNamedRange(ByVal SelectionRange As Range _
                                     , Optional ByVal IsLocal As Boolean = False _
                                      , Optional ByVal IsSilentMode As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.AddTableColumnNamedRange"
    ' For each area in SelectionRange, and for each column in each area, apply the table column named range using the first cell of each column
    Dim CurrentRange As Range
    For Each CurrentRange In SelectionRange.Areas
        Dim CurrentCell As Range
        For Each CurrentCell In CurrentRange.Columns
            ApplyTableColumnNameRangeForCurrentCell CurrentCell.Cells(1, 1), IsLocal, IsSilentMode
        Next CurrentCell
    Next CurrentRange
    Logger.Log TRACE_LOG, "Exit modNamedRange.AddTableColumnNamedRange"

End Sub

Private Sub ApplyTableColumnNameRangeForCurrentCell(ByVal CurrentCell As Range _
                                                    , Optional ByVal IsLocal As Boolean = False _
                                                     , Optional ByVal IsSilentMode As Boolean = False)
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.ApplyTableColumnNameRangeForCurrentCell"
    ' If the current cell is not inside a table, exit the sub
    If Not IsInsideTable(CurrentCell) Then Exit Sub
    
    ' Generate the probable name for the named range based on table name and column header
    Dim Table As ListObject
    Set Table = CurrentCell.ListObject
    Dim ColumnHeading As String
    ColumnHeading = FindTableColumnHeader(Table, CurrentCell)
    Dim ProbableName As String
    ProbableName = Table.Name & "_" & MakeValidDefinedName(ColumnHeading, False)
    
    ' Replace "tbl" with "col" if it is the starting of the probable name
    If Text.IsStartsWith(ProbableName, "tbl") Then
        ProbableName = VBA.Replace(ProbableName, "tbl", "col", 1, 1)
    End If
    
    ' Get the final defined name if not in silent mode
    If Not IsSilentMode Then
        ProbableName = GetFinalDefineName(ProbableName, True)
    End If
    
    ' Apply the named range for the current cell
    ApplyNameRange Table.ListColumns(ColumnHeading).DataBodyRange, ProbableName, False, IsLocal, CurrentCell.Worksheet, False, True, True
    Logger.Log TRACE_LOG, "Exit modNamedRange.ApplyTableColumnNameRangeForCurrentCell"
    
End Sub

Private Function FindTableColumnHeader(ByVal Table As ListObject, ByVal SelectionRange As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindTableColumnHeader"
    ' Find the index of the column based on the difference between the SelectionRange column and the first column of the table header
    Dim Index As Long
    Index = SelectionRange.Column - Table.HeaderRowRange.Cells(1, 1).Column + 1
    
    ' Return the column header based on the index. If multiple columns are selected, a range of column headers is returned
    If SelectionRange.Columns.Count = 1 Then
        FindTableColumnHeader = Table.ListColumns(Index).Name
    Else
        FindTableColumnHeader = LEFT_SQUARE_BRACKET & Table.ListColumns(Index).Name _
                              & RIGHT_SQUARE_BRACKET & ":" & LEFT_SQUARE_BRACKET _
                              & Table.ListColumns(Index + SelectionRange.Columns.Count - 1).Name _
                              & RIGHT_SQUARE_BRACKET
    End If
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindTableColumnHeader"
    
End Function

Private Function FindNamedRangeColumnHeader(ByVal GivenName As Name, ByVal AnyCellInsideName As Range) As String
    
    Logger.Log TRACE_LOG, "Enter modNamedRange.FindNamedRangeColumnHeader"
    ' Calculate index based on the difference between the column number of the cell and the column number of the first cell in the given name range
    Dim Index As Long
    Index = AnyCellInsideName.Column - GivenName.RefersToRange.Cells(1, 1).Column + 1
    
    ' Return the value of the cell in the first row of the column specified by the index
    FindNamedRangeColumnHeader = GivenName.RefersToRange.Columns(Index).Cells(1, 1).Value
    Logger.Log TRACE_LOG, "Exit modNamedRange.FindNamedRangeColumnHeader"
    
End Function

Private Sub ApplyNamedRangeColumnNameRangeForCurrentCell(ByVal CurrentCell As Range _
                                                         , Optional ByVal IsLocal As Boolean = False _
                                                          , Optional ByVal IsSilentMode As Boolean = False)
    Logger.Log TRACE_LOG, "Enter modNamedRange.ApplyNamedRangeColumnNameRangeForCurrentCell"
    ' If the current cell is inside a named range, generate a probable name and apply the named range for the current cell
    If IsInsideNamedRange(CurrentCell) Then
        ' Find the name that includes the current cell
        Dim CurrentName As Name
        Set CurrentName = FindNamedRangeFromSubCell(CurrentCell)
        
        ' Find the column header for the named range
        Dim ColumnHeading As String
        ColumnHeading = FindNamedRangeColumnHeader(CurrentName, CurrentCell)
        
        ' Create a probable name based on the current name and the column heading
        Dim ProbableName As String
        ProbableName = CurrentName.Name & "_" & MakeValidDefinedName(ColumnHeading, False)
        
        ' If the probable name starts with "tbl", replace it with "col"
        If Text.IsStartsWith(ProbableName, "tbl") Then
            ProbableName = VBA.Replace(ProbableName, "tbl", "col", 1, 1)
        End If
        
        ' Get the final name if not in silent mode
        If Not IsSilentMode Then
            ProbableName = GetFinalDefineName(ProbableName, True)
        End If
        
        ' Apply the named range for the current cell with the final or probable name
        ApplyNameRange CurrentCell, ProbableName, False, IsLocal, CurrentCell.Parent, False, True, True
    End If
    Logger.Log TRACE_LOG, "Exit modNamedRange.ApplyNamedRangeColumnNameRangeForCurrentCell"
End Sub

' --------------------------------------------< OA Robot >--------------------------------------------
'  Command Name:           Expand Named Range
'  Description:            Expand named range.
'  Macro Expression:       modNamedRange.ExpandNamedRange([Selection])
'  Generated:              03/08/2023 07:27 PM
' ----------------------------------------------------------------------------------------------------
Public Sub ExpandNamedRange(ByVal SelectionRange As Range)
    
    ' If no range is provided, exit the subroutine
    If IsNothing(SelectionRange) Then Exit Sub
    
    ' Create a new collection to store valid named ranges
    Dim AllValidNamedRange As Collection
    Set AllValidNamedRange = New Collection
    
    ' Loop through all names in the workbook
    Dim CurrentName As Name
    For Each CurrentName In SelectionRange.Worksheet.Parent.Names
        ' If the named range refers to an actual range
        If Not IsRefersToRangeIsNothing(CurrentName) Then
            If CurrentName.RefersToRange.Worksheet.Name = SelectionRange.Worksheet.Name Then
                ' If the first cell in the range referred by the name intersects with the first cell in the selection range, add the name to the collection
                If IsNotNothing(Intersect(CurrentName.RefersToRange.Cells(1, 1), SelectionRange.Cells(1, 1))) Then
                    AllValidNamedRange.Add CurrentName
                End If
            End If
        End If
    Next CurrentName
    
    ' If no named ranges are found, show an informational message
    If AllValidNamedRange.Count = 0 Then
        MsgBox "No named range found to expand.", vbInformation, "Expand Named Range"
        ' If only one named range is found, expand it to the entire selection range
    ElseIf AllValidNamedRange.Count = 1 Then
        Set CurrentName = AllValidNamedRange.Item(1)
        CurrentName.RefersTo = "=" & GetSheetRefForRangeReference(SelectionRange.Worksheet.Name) & SelectionRange.Address
        ' If more than one named range are found, show an informational message
    Else
        MsgBox AllValidNamedRange.Count & " named range found to expand.But this feature is only applicable for one named range." _
               , vbInformation, "Expand Named Range"
    End If
    
End Sub

Public Function RemoveInitialSpaceAndNewLines(ByVal OperationOnText As String) As String

    Dim Index As Long
    Dim CurrentChar As String
    For Index = 1 To VBA.Len(OperationOnText)
        CurrentChar = VBA.Mid$(OperationOnText, Index, 1)
        If Not (CurrentChar = Space(1) Or CurrentChar = vbNewLine Or CurrentChar = VBA.Chr$(10)) Then
            Exit For
        End If
    Next Index
    RemoveInitialSpaceAndNewLines = VBA.Mid$(OperationOnText, Index)

End Function

