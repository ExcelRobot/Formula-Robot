Attribute VB_Name = "modUtility"
Option Explicit

#If VBA7 Then                                    ' Excel 2010 or later
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else                                            ' Excel 2007 or earlier
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

Public Function RemoveStartingEqualSign(ByVal FormulaText As String) As String
    
    Dim Result As String
    Result = Text.RemoveFromStartIfPresent(LTrim$(FormulaText), EQUAL_SIGN)
    
    RemoveStartingEqualSign = Result
    
End Function

Public Function IsInFirstRowOfSpillRange(ByVal CheckOnCell As Range) As Boolean
    
    If CheckOnCell.Cells.Count > 1 Or Not CheckOnCell.Cells(1).HasSpill Then
        IsInFirstRowOfSpillRange = False
    Else
        IsInFirstRowOfSpillRange = (Not Intersect(CheckOnCell.Cells(1).SpillParent.SpillingToRange.Rows(1) _
                                                  , CheckOnCell) Is Nothing)
    End If
    
End Function

Public Function IsInFirstColOfSpillRange(ByVal CheckOnCell As Range) As Boolean
    
    If CheckOnCell.Cells.Count > 1 Or Not CheckOnCell.Cells(1).HasSpill Then
        IsInFirstColOfSpillRange = False
    Else
        IsInFirstColOfSpillRange = (Not Intersect(CheckOnCell.SpillParent.Cells(1).SpillingToRange.Columns(1) _
                                                  , CheckOnCell) Is Nothing)
    End If

End Function

Public Function IsSpillParent(ByVal CheckOnCell As Range) As Boolean
    
    Dim Result As Boolean
    If Not CheckOnCell.HasSpill Then
        Result = False
    ElseIf CheckOnCell.Cells.Count > 1 Then
        Result = False
    ElseIf CheckOnCell.SpillParent.Address = CheckOnCell.Address Then
        Result = True
    End If
    
    IsSpillParent = Result
    
End Function

Public Function GenerateArraySeq(ByVal StartFrom As Long _
                                 , ByVal Count As Long _
                                  , ByVal IsSpillInRows As Boolean) As String
    
    Dim Formula As String
    Formula = LEFT_BRACE
    Dim Counter As Long
    Dim Delimiter As String
    Delimiter = IIf(IsSpillInRows, ARRAY_CONST_ROW_SEPARATOR, ARRAY_CONST_COLUMN_SEPARATOR)
    For Counter = 1 To Count
        Formula = Formula & (StartFrom + Counter - 1) & Delimiter
    Next Counter
    
    Formula = Text.RemoveFromEndIfPresent(Formula, Delimiter)
    If Formula <> LEFT_BRACE Then
        Formula = Formula & RIGHT_BRACE
    End If
    GenerateArraySeq = Formula
    
End Function

Public Function GetParentCellRefIfNoSpill(ByVal FormulaCell As Range _
                                          , ByVal CurrentRange As Range _
                                           , ByVal IsAbsoluteRef As Boolean) As String
    
    Dim CellRef As String
    If FormulaCell.Worksheet.Name <> CurrentRange.Worksheet.Name Then
        CellRef = GetSheetRefForRangeReference(CurrentRange.Worksheet.Name)
    End If
    CellRef = CellRef & CurrentRange.Address(IsAbsoluteRef, IsAbsoluteRef)
    GetParentCellRefIfNoSpill = CellRef
    
End Function

Public Function GetParentCellRef(ByVal FormulaCell As Range _
                                 , ByVal CurrentRange As Range _
                                  , ByVal IsAbsoluteRef As Boolean) As String
    
    Dim CellRef As String
    If FormulaCell.Worksheet.Name <> CurrentRange.Worksheet.Name Then
        CellRef = GetSheetRefForRangeReference(CurrentRange.Worksheet.Name)
    End If
    CellRef = CellRef & CurrentRange.Cells(1).SpillParent.Address(IsAbsoluteRef, IsAbsoluteRef) & HASH_SIGN
    GetParentCellRef = CellRef
    
End Function

Public Function GetParamNameFromCounter(ByVal StartStepName As String _
                                        , ByVal Counter As Long _
                                         , ByVal TotalValidCells As Long) As String
    
    Dim ParamName As String
    If TotalValidCells > 3 Then
        ParamName = StartStepName & "_" & (Counter)
    Else
        ParamName = Application.WorksheetFunction.Rept(Chr$(Asc("x") + Counter - 1), Len(StartStepName))
    End If
    GetParamNameFromCounter = ParamName
    
End Function

Public Function FindLambdas(ByVal FromBook As Workbook) As Collection
    
    ' Finds all lambda functions in the given workbook and returns a collection of their names.
    Dim CurrentName As Name
    Dim AllLambda As Collection
    Set AllLambda = New Collection
    For Each CurrentName In FromBook.Names
        ' Check if the name refers to a lambda function.
        If IsLambdaFunction(CurrentName.RefersTo) Then
            ' Add the name to the collection of lambda functions.
            AllLambda.Add CurrentName, CStr(CurrentName.Name)
        End If
    Next CurrentName
    Set FindLambdas = AllLambda
    
End Function

Public Function IsExistInCollection(ByVal GivenCollection As Collection _
                                    , ByVal Key As String) As Boolean
    
    ' Check if the given Key exists in the Collection.
    On Error GoTo NotExist
    Dim Item  As Variant
    If IsObject(GivenCollection.Item(Key)) Then
        Set Item = GivenCollection.Item(Key)
    Else
        Item = GivenCollection.Item(Key)
    End If
    IsExistInCollection = True
    Exit Function
    
NotExist:
    IsExistInCollection = False
    On Error GoTo 0
    
End Function

Public Function CollectionToArray(ByVal GivenCollection As Collection) As Variant
    
    ' Convert a Collection into a 1D Variant Array.

    If GivenCollection.Count = 0 Then Exit Function
    Dim Result() As Variant
    ReDim Result(1 To GivenCollection.Count, 1 To 1)
    Dim CurrentElement As Variant
    Dim CurrentIndex As Long
    For Each CurrentElement In GivenCollection
        CurrentIndex = CurrentIndex + 1
        Result(CurrentIndex, 1) = CurrentElement
    Next CurrentElement
    CollectionToArray = Result

End Function

Public Function IsEndsWith(ByVal TestOnText As String, ByVal TextToMatch As String) As Boolean
    IsEndsWith = (UCase$(Right$(TestOnText, Len(TextToMatch))) = UCase$(TextToMatch))
End Function

Public Function IsNamedRangeExist(ByVal SearchInBook As Workbook _
                                  , ByVal NameOfTheNamedRange As String) As Boolean
    
    ' Checks if a named range exists in the given workbook.

    Dim IsExist As Boolean
    Dim CurrentName As Name
    For Each CurrentName In SearchInBook.Names
        If CurrentName.Name = NameOfTheNamedRange Then
            IsExist = True
            Exit For
        End If
    Next CurrentName
    IsNamedRangeExist = IsExist
    
End Function

Public Function IsTextPresent(ByVal SearchInText As String, ByVal SearchForText As String) As Boolean
    IsTextPresent = (InStr(1, SearchInText, SearchForText, vbTextCompare) <> 0)
End Function

Public Sub AssingOnUndo(ByVal UndoForMethod As String)
    
    ' Assigns an Undo method for the specified action.
    Const SINGLE_QUOTE As String = "'"
    Const EXCLAMATION_SIGN As String = "!"
    Dim UndoSubName As String
    UndoSubName = SINGLE_QUOTE & ThisWorkbook.Name & SINGLE_QUOTE & EXCLAMATION_SIGN & UndoForMethod & "_Undo"
    Application.OnUndo "Undo " & UndoForMethod & " Action", UndoSubName
    
End Sub

Public Function IsOneColSpillRange(ByVal CellInsideSpillRange As Range) As Boolean
    IsOneColSpillRange = (SpillRangeColCount(CellInsideSpillRange) = 1)
End Function

Public Function IsOneRowSpillRange(ByVal CellInsideSpillRange As Range) As Boolean
    IsOneRowSpillRange = (SpillRangeRowCount(CellInsideSpillRange) = 1)
End Function

Public Function SpillRangeColCount(ByVal CellInsideSpillRange As Range) As Long
    
    SpillRangeColCount = 1
    If CellInsideSpillRange.Cells(1).HasSpill Then
        SpillRangeColCount = CellInsideSpillRange.Cells(1).SpillParent.SpillingToRange.Columns.Count
    End If

End Function

Public Function SpillRangeRowCount(ByVal CellInsideSpillRange As Range) As Long
    
    SpillRangeRowCount = 1
    If CellInsideSpillRange.Cells(1).HasSpill Then
        SpillRangeRowCount = CellInsideSpillRange.Cells(1).SpillParent.SpillingToRange.Rows.Count
    End If
    
End Function

Public Function DropFirstCell(ByVal FromRange As Range) As Range
    
    If FromRange.Cells.Count > 1 Then
        If FromRange.Rows.Count > 1 And FromRange.Columns.Count = 1 Then
            Set DropFirstCell = FromRange.Offset(1).Resize(FromRange.Rows.Count - 1, FromRange.Columns.Count)
        ElseIf FromRange.Columns.Count > 1 And FromRange.Rows.Count = 1 Then
            Set DropFirstCell = FromRange.Offset(, 1).Resize(FromRange.Rows.Count, FromRange.Columns.Count - 1)
        End If
    End If
    
End Function

Public Function IsFormulaTextSame(ByVal FirstCell As Range, ByVal SecondCell As Range) As Boolean
    IsFormulaTextSame = (FormatFormula(FirstCell.Formula2) = FormatFormula(SecondCell.Formula2))
End Function

Private Sub TestIsRowAbsolute()
    
    Debug.Print "No Abs: " & Not (IsRowAbsolute("A1"))
    Debug.Print "Col Abs: " & Not (IsRowAbsolute("$A1"))
    Debug.Print "Row Abs: " & IsRowAbsolute("A$1")
    Debug.Print "Both Row and Col Abs: " & IsRowAbsolute("$A$1")
    Debug.Print "Rectangle Range: " & IsRowAbsolute("$A$1:$A$4")
    Debug.Print "Rectangle Range: " & IsRowAbsolute("A$1:A$4")
    Debug.Print "Rectangle Range: " & Not (IsRowAbsolute("$A1:$A4"))
    Debug.Print "Rectangle Range: " & Not (IsRowAbsolute("A1:A4"))
    Debug.Print "Whole Row: " & Not (IsRowAbsolute("27:27"))
    Debug.Print "Whole Row: " & IsRowAbsolute("$27:$27")
    Debug.Print "Whole Col: " & Not (IsRowAbsolute("A:D"))
    Debug.Print "Whole Col: " & Not (IsRowAbsolute("$A:$D"))
    Debug.Print "Spill Range Ref: " & (IsRowAbsolute("$F$5#"))
    
End Sub

Public Function IsRowAbsolute(ByVal Reference As String) As Boolean
    
    IsRowAbsolute = True
    
    Const DOLLAR_SIGN As String = "$"
    Dim DollarSignPos As Long
    
    Dim SplittedByColon As Variant
    SplittedByColon = Split(Text.RemoveFromEndIfPresent(Reference, "#"), ":")
    
    Dim CurrentA1 As Variant
    For Each CurrentA1 In SplittedByColon
         
        DollarSignPos = InStrRev(CStr(CurrentA1), DOLLAR_SIGN)
        If DollarSignPos = 0 Then
            IsRowAbsolute = False
            Exit For
        End If
        
        Dim TextAfterDollarSign As String
        TextAfterDollarSign = Mid(CStr(CurrentA1), DollarSignPos + 1)
        If Not IsNumeric(TextAfterDollarSign) Then
            IsRowAbsolute = False
            Exit For
        End If
    
    Next CurrentA1
    
End Function

Private Sub TestIsColAbsolute()
    
    Debug.Print "No Abs: " & Not (IsColAbsolute("A1"))
    Debug.Print "Col Abs: " & IsColAbsolute("$A1")
    Debug.Print "Row Abs: " & Not (IsColAbsolute("A$1"))
    Debug.Print "Both Row and Col Abs: " & IsColAbsolute("$A$1")
    Debug.Print "Rectangle Range: " & IsColAbsolute("$A$1:$A$4")
    Debug.Print "Rectangle Range: " & Not (IsColAbsolute("A$1:A$4"))
    Debug.Print "Rectangle Range: " & IsColAbsolute("$A1:$A4")
    Debug.Print "Rectangle Range: " & Not (IsColAbsolute("A1:A4"))
    Debug.Print "Whole Row: " & Not (IsColAbsolute("27:27"))
    Debug.Print "Whole Row: " & Not (IsColAbsolute("$27:$27"))
    Debug.Print "Whole Col: " & Not (IsColAbsolute("A:D"))
    Debug.Print "Whole Col: " & IsColAbsolute("$A:$D")
    Debug.Print "Spill Range Ref: " & (IsColAbsolute("$F$5#"))
    
End Sub

Public Function IsColAbsolute(ByVal Reference As String) As Boolean
    
    IsColAbsolute = True
    
    Const DOLLAR_SIGN As String = "$"
    Dim DollarSignPos As Long
    
    Dim SplittedByColon As Variant
    SplittedByColon = Split(Text.RemoveFromEndIfPresent(Reference, "#"), ":")
    
    Dim CurrentA1 As Variant
    For Each CurrentA1 In SplittedByColon
         
        DollarSignPos = InStr(1, CStr(CurrentA1), DOLLAR_SIGN)
        If DollarSignPos = 0 Then
            IsColAbsolute = False
            Exit For
        End If
        
        ' As Column ref come before row ref($A1) also handle only row ref($27:$27)
        If DollarSignPos <> 1 Or IsNumeric(Mid(CStr(CurrentA1), DollarSignPos + 1)) Then
            IsColAbsolute = False
            Exit For
        End If
        
    Next CurrentA1
    
End Function

Public Function MaxRowCount(ByVal ValidCells As Collection) As Long
    
    Dim Count As Long
    Dim CurrentItem As PrecedencyInfo
    For Each CurrentItem In ValidCells
        If Count < CurrentItem.RowCount Then Count = CurrentItem.RowCount
    Next CurrentItem
    MaxRowCount = Count
    
End Function

Public Function IsSpillRowCountSame(ByVal TestFormula As String _
                                    , ByVal DestinationRange As Range _
                                    , ByVal ExpectedRowCount As Long) As Boolean
    
    On Error GoTo HandleError
    Dim OldFormula As String
    OldFormula = DestinationRange.Cells(1).Formula2
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(TestFormula)
    If DestinationRange.HasSpill Then
        IsSpillRowCountSame = (SpillRangeRowCount(DestinationRange) = ExpectedRowCount)
    End If
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Exit Function
    
HandleError:
    IsSpillRowCountSame = False
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    
End Function

Public Function IsSpillColCountSame(ByVal TestFormula As String _
                                    , ByVal DestinationRange As Range _
                                     , ByVal ExpectedColCount As Long) As Boolean
    
    On Error GoTo HandleError
    Dim OldFormula As String
    OldFormula = DestinationRange.Cells(1).Formula2
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(TestFormula)
    If DestinationRange.HasSpill Then
        IsSpillColCountSame = (SpillRangeColCount(DestinationRange) = ExpectedColCount)
    End If
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    Exit Function
    
HandleError:
    IsSpillColCountSame = False
    DestinationRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(OldFormula)
    
End Function

Public Function IsStartCellSame(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    IsStartCellSame = ((FirstRange.Cells(1).Address = SecondRange.Cells(1).Address) _
                       And (FirstRange.Worksheet.Name = SecondRange.Worksheet.Name))
End Function

Public Function IsAllBlankAndNoFormulas(ByVal CheckCells As Range) As Boolean

    If CheckCells Is Nothing Then
        IsAllBlankAndNoFormulas = False
        Exit Function
    End If
    
    IsAllBlankAndNoFormulas = True
    Dim CurrentCell As Range
    For Each CurrentCell In CheckCells.Cells
        
        If CurrentCell.HasFormula Or Not IsBlankCellNoError(CurrentCell) Then
            IsAllBlankAndNoFormulas = False
            Exit For
        End If
        
    Next CurrentCell

End Function

Public Function IsNothing(ByVal GivenObject As Object) As Boolean
    IsNothing = (GivenObject Is Nothing)
End Function

Public Function IsNotNothing(ByVal GivenObject As Object) As Boolean
    IsNotNothing = (Not GivenObject Is Nothing)
End Function

Public Function HasDynamicFormula(ByVal SelectionRange As Range) As Boolean
    
    ' Check if the selected range contains a dynamic formula (spill range).
    On Error Resume Next
    HasDynamicFormula = SelectionRange.Cells(1).HasSpill
    On Error GoTo 0
    
End Function

Public Function GetOldNameFromComment(ByVal FromCell As Range, ByVal Prefix As String) As String
    
    ' Retrieves the old lambda name from the comment in the FromCell with the specified prefix.
    On Error GoTo NoComment
    Dim CurrentComment As Comment
    Set CurrentComment = FromCell.Comment
    If Text.IsStartsWith(CurrentComment.Text, Prefix) Then
        GetOldNameFromComment = Replace(CurrentComment.Text, Prefix, vbNullString)
    End If
    Exit Function

NoComment:
    GetOldNameFromComment = vbNullString
    
End Function

Public Function ExtractStartFormulaName(ByVal FormulaText As String) As String
    
    ' Extracts the formula name from the given formula text.
    ' If the formula contains parentheses (indicating a function call), it extracts the name before the first parenthesis.
    ' Otherwise, it considers the entire formula text as the name.

    If Text.Contains(FormulaText, FIRST_PARENTHESIS_OPEN) Then
        ExtractStartFormulaName = Text.BeforeDelimiter(FormulaText, FIRST_PARENTHESIS_OPEN)
    Else
        ExtractStartFormulaName = FormulaText
    End If

    ' Remove the equal sign if present at the beginning and trim any leading/trailing spaces.
    ExtractStartFormulaName = Text.RemoveFromStartIfPresent(ExtractStartFormulaName, EQUAL_SIGN)
    ExtractStartFormulaName = Text.Trim(ExtractStartFormulaName)

End Function

Private Function IsFirstCharEqualExceptWhiteSpace(ByVal GivenText As String) As Boolean
    GivenText = RemoveInitialSpaceAndNewLines(GivenText)
    IsFirstCharEqualExceptWhiteSpace = Text.IsStartsWith(GivenText, EQUAL_SIGN)
End Function

Public Sub DeleteComment(ByVal ToCell As Range)
    
    Dim CurrentComment As Comment
    Set CurrentComment = ToCell.Comment
    On Error GoTo ExitSub
    CurrentComment.Delete
    Exit Sub
    
ExitSub:
    
End Sub

Public Sub UpdateOrAddNamedRangeNameNote(ByVal ToCell As Range _
                                     , ByVal LambdaName As String _
                                      , ByVal Prefix As String)
    
    On Error Resume Next
    DeleteComment ToCell
    ToCell.Cells(1).AddComment Prefix & LambdaName
    On Error GoTo 0
    
End Sub

Public Function GetCellValueIfErrorNullString(ByVal GivenCell As Range) As String
    
    ' Get the cell value or return an empty string if an error is encountered.
    If IsError(GivenCell.Value) Then
        GetCellValueIfErrorNullString = vbNullString
    Else
        GetCellValueIfErrorNullString = GivenCell.Value
    End If
    
End Function

Public Function FindNamedRangeFromSubCell(ByVal GivenRange As Range) As Name
    
    ' Find the named range containing the given range.
    Dim CurrentNameRange As Name
    Dim NameOfCurrentNamedRange As String
    Dim ReferredRange As Range
    For Each CurrentNameRange In GivenRange.Worksheet.Parent.Names
        If CurrentNameRange.Visible Then
            NameOfCurrentNamedRange = Replace(CurrentNameRange.Name, EQUAL_SIGN, vbNullString)
            On Error Resume Next
            Set ReferredRange = CurrentNameRange.RefersToRange
            On Error GoTo 0
            If IsNothing(ReferredRange) Then
                ' Logger.Log DEBUG_LOG, NameOfCurrentNamedRange & " not found"
                ' Debug.Assert NameOfCurrentNamedRange <> "_xlpm.side1"
            ElseIf GivenRange.Worksheet.Name = ReferredRange.Worksheet.Name Then
                If HasIntersection(ReferredRange, GivenRange) Then
                    Set FindNamedRangeFromSubCell = CurrentNameRange
                    Exit Function
                End If
            End If
        End If
    Next CurrentNameRange

    Set FindNamedRangeFromSubCell = Nothing

End Function

Public Function HasIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    HasIntersection = IsNotNothing(FindIntersection(FirstRange, SecondRange))
End Function

Public Function FindIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Range
    
    On Error Resume Next
    Set FindIntersection = Intersect(FirstRange, SecondRange)
    On Error GoTo 0
    
End Function

Public Function IsCellHidden(ByVal CurrentCell As Range) As Boolean
    ' Check if the CurrentCell or its entire row/column is hidden.
    IsCellHidden = (CurrentCell.EntireColumn.Hidden Or CurrentCell.EntireRow.Hidden)
End Function

Public Function IsLocalScopeNamedRange(ByVal LocalName As String) As Boolean
    
    Dim FoundAt As Long
    FoundAt = InStr(1, LocalName, SHEET_NAME_SEPARATOR)
    IsLocalScopeNamedRange = (FoundAt <> 0)
    
End Function

Public Function IsInsideNamedRange(ByVal GivenRange As Range) As Boolean
    
    ' Check if the given range is inside a named range.
    Dim CurrentName As Name
    Set CurrentName = FindNamedRangeFromSubCell(GivenRange)
    IsInsideNamedRange = IsNotNothing(CurrentName)
    
End Function

Public Function IsInsideTableOrNamedRange(ByVal CurrentCell As Range) As Boolean
    
    ' Checks if the given cell is inside a table or a named range.
    ' CurrentCell: The cell to check.

    If modUtility.IsInsideNamedRange(CurrentCell) Then
        IsInsideTableOrNamedRange = True
    ElseIf modUtility.IsInsideTable(CurrentCell) Then
        IsInsideTableOrNamedRange = True
    End If
    
End Function

Public Function IsInsideTable(ByVal GivenRange As Range) As Boolean
    
    ' Check if the given range is inside a table.
    Dim ActiveTable As ListObject
    Set ActiveTable = GetTableFromRange(GivenRange)
    IsInsideTable = IsNotNothing(ActiveTable)

End Function

Public Function GetTableFromRange(ByVal GivenRange As Range) As ListObject
    ' Get the table object from the given range.
    Set GetTableFromRange = GivenRange.ListObject
End Function

Private Function ConvertToProperColumnName(ByVal GivenColumnName As String) As String
    
    
    ' We convert the GivenColumnName to a correct structured reference format
    Dim SpecialCharsToPutEscapeChar As Variant
    ' Sequence is important here as escape character is single quote
    
    SpecialCharsToPutEscapeChar = Array(SINGLE_QUOTE, HASH_SIGN, "[", "]")
    ' Ref : https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e
    Dim CurrentChar As Variant
    For Each CurrentChar In SpecialCharsToPutEscapeChar
        GivenColumnName = VBA.Replace(GivenColumnName, CurrentChar, SINGLE_QUOTE & CurrentChar)
    Next CurrentChar
    
    ConvertToProperColumnName = LEFT_SQUARE_BRACKET & GivenColumnName & RIGHT_SQUARE_BRACKET
    
    
End Function

Private Sub Test()
    
    Dim ValidName As String
    ValidName = MakeValidName("colChartData_Growth1000_Fund", True)
    Debug.Print ValidName
    Debug.Print "Test Pass: " & (ValidName = "colChartData_Growth1000_Fund")
    
End Sub

Public Function MakeValidName(ByVal GivenInvalidName As String _
                              , ByVal JustRemoveInvalidChars As Boolean) As String
    
    Dim ValidName As String
    ' Replace Newline with space.
    ValidName = ReplaceLineBreak(Trim$(GivenInvalidName), ONE_SPACE)
    
    ValidName = ReplacePlaceHolders(ValidName)
    
    ' Remove Invalid char but keep space.
    ValidName = RemoveInvalidCharcters(ValidName, True)
    
    If Not JustRemoveInvalidChars Then
        
        ' Replace dots with underscores in the name.
        ValidName = VBA.Replace(ValidName, DOT, UNDER_SCORE)
    
        ' Convert To proper sentence form.
        ValidName = Text.Trim(ConvertVarNameToSentence(ValidName))
        ValidName = ConvertToPascalCase(ValidName)
        
    Else
        ValidName = VBA.Replace(ValidName, ONE_SPACE, vbNullString)
    End If
    
    ' If the name is a range reference, split it and add underscores.
    If IsRangeReference(ValidName) Then
        Dim ColRefAndRowRef As Collection
        Set ColRefAndRowRef = Text.SplitDigitAndNonDigit(ValidName)
        ValidName = ColRefAndRowRef.Item(1) & UNDER_SCORE & ColRefAndRowRef.Item(2)
    End If
    
    ' Limit the length of the name to MAX_ALLOWED_LENGTH.
    If Len(ValidName) > modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH Then
        ValidName = Left$(ValidName, modSharedConstant.MAX_ALLOWED_LET_STEP_NAME_LENGTH)
    End If
    
    MakeValidName = ValidName
    
End Function

Private Function ReplaceLineBreak(ByVal Text As String, ReplaceWith As String) As String
    
    Dim ReplacedText As String
    ReplacedText = Replace(Text, vbNewLine, ReplaceWith)
    ReplacedText = Replace(ReplacedText, Chr$(10), ReplaceWith)
    ReplacedText = Replace(ReplacedText, Chr$(13), ReplaceWith)
    ReplaceLineBreak = ReplacedText
    
End Function

' Replace specific placeholders with their corresponding values.
Public Function ReplacePlaceHolders(ByVal GivenName As String) As String

    Dim PlaceHolders As Variant
    PlaceHolders = Array("%", HASH_SIGN, "&", "<", ">", EQUAL_SIGN)

    Dim ReplaceWiths As Variant
    ReplaceWiths = Array("Percent", "Number", "And", "LessThan", "GreaterThan", "Equals")

    Dim CurrentIndex As Long

    ' Loop through each placeholder and replace it with the corresponding value.
    For CurrentIndex = LBound(PlaceHolders) To UBound(PlaceHolders)
        GivenName = Replace(GivenName, PlaceHolders(CurrentIndex), ReplaceWiths(CurrentIndex))
    Next CurrentIndex

    ' Return the modified name.
    ReplacePlaceHolders = GivenName

End Function

' Remove invalid characters from the given name.
Public Function RemoveInvalidCharcters(ByVal GivenName As String, KeepSpace As Boolean) As String

    Dim Output As String
    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)
        ' Check if the current character is a valid first character for the name.
        If IsValidFirstChar(CurrentChar) Then
            Output = CurrentChar
            Exit For
        End If
    Next CurrentCharIndex

    ' If the given name is not empty and there are characters after the first valid character,
    ' update the name accordingly. Otherwise, set the name to an empty string.
    If Len(GivenName) <> CurrentCharIndex And Len(GivenName) > CurrentCharIndex Then
        GivenName = Right$(GivenName, Len(GivenName) - CurrentCharIndex)
    Else
        GivenName = vbNullString
    End If

    ' Return the updated name with valid characters.
    RemoveInvalidCharcters = Output & GetValidCharForSecondToOnward(GivenName, KeepSpace)

End Function

Public Function ConvertVarNameToSentence(VarName As String) As String
    
    Dim Sentence As String
    Sentence = Replace(VarName, DOT, ONE_SPACE)
    Sentence = Replace(Sentence, UNDER_SCORE, ONE_SPACE)
    Sentence = ReplaceLineBreak(Sentence, ONE_SPACE)
    Sentence = ConcatenateCollection(Text.SplitDigitAndNonDigit(Sentence), ONE_SPACE)
    Dim Words As Variant
    Words = Split(Trim$(Sentence), ONE_SPACE)
    Sentence = vbNullString
    
    Dim Word As Variant
    For Each Word In Words
        Sentence = Sentence & ONE_SPACE & PutSpaceOnLowerCaseToUpperCaseTransition(Word)
    Next Word
    
    Words = Split(Trim$(Sentence), ONE_SPACE)
    Sentence = vbNullString
    
    For Each Word In Words
        Sentence = Sentence & ONE_SPACE & PutSpaceBeforeLastCapsFromStart(Word)
    Next Word
    ConvertVarNameToSentence = Trim$(Sentence)
    
End Function

'  This just replace space with VBNullstring and convert first char of each word to upper case except first one.
Private Function ConvertToCamelCase(ByVal VarName As String) As String
    
    Dim ValidName As String
    ValidName = Text.Trim(CapitalizeFirstCharOfEachWord(VarName))
    If Text.Contains(ValidName, ONE_SPACE) Then
        
        If Not IsAllCaps(Text.BeforeDelimiter(ValidName, ONE_SPACE)) Then
            ValidName = LCase(Text.BeforeDelimiter(ValidName, ONE_SPACE)) & ONE_SPACE _
                        & ConvertToProperCaseOfEachWord( _
                        Text.AfterDelimiter(ValidName, ONE_SPACE))
        End If
        
    Else
        If Not IsAllCaps(ValidName) Then
            ValidName = LCase(ValidName)
        End If
    End If
    ValidName = Replace(ValidName, ONE_SPACE, vbNullString)
    
    ConvertToCamelCase = ValidName
    
End Function

'  This just replace space with VBNullstring and convert first char of each word to upper case
Private Function ConvertToPascalCase(ByVal VarName As String) As String
    
    Dim ValidName As String
    ValidName = Text.Trim(CapitalizeFirstCharOfEachWord(VarName))
    ValidName = ConvertToProperCaseOfEachWord(ValidName)
    ValidName = Replace(ValidName, ONE_SPACE, vbNullString)
    ConvertToPascalCase = ValidName
    
End Function

' Check if the given reference is a valid range reference.
Public Function IsRangeReference(ByVal GivenRef As String) As Boolean

    ' Use ConvertFormula to try converting the reference to R1C1 notation.
    If IsError(Application.ConvertFormula("=" & GivenRef, xlA1, xlR1C1, , Range("A1"))) Then
        IsRangeReference = False
    Else
        ' Check if the converted R1C1 notation is different from the original reference.
        IsRangeReference = (UCase$(Application.ConvertFormula("=" & GivenRef _
                                                              , xlA1, xlR1C1 _
                                                                     , , Range("A1"))) <> UCase$("=" & GivenRef))
    End If

End Function

' Check if the given character is a valid first character for the name.
Public Function IsValidFirstChar(ByVal GivenChar As String) As Boolean
    
    Static InvalidFirstChars As Collection
    If InvalidFirstChars Is Nothing Then
        Set InvalidFirstChars = New Collection
        With InvalidFirstChars
            AddCharsToColl InvalidFirstChars, 1, 64
            .Add 91, CStr(91)
            AddCharsToColl InvalidFirstChars, 93, 94
            .Add 96, CStr(96)
            AddCharsToColl InvalidFirstChars, 123, 130
            .Add 132, CStr(132)
            .Add 136, CStr(136)
            .Add 139, CStr(139)
            .Add 141, CStr(141)
            AddCharsToColl InvalidFirstChars, 143, 144
            .Add 149, CStr(149)
            .Add 152, CStr(152)
            .Add 155, CStr(155)
            .Add 157, CStr(157)
            .Add 160, CStr(160)
            AddCharsToColl InvalidFirstChars, 162, 163
            AddCharsToColl InvalidFirstChars, 165, 166
            .Add 169, CStr(169)
            AddCharsToColl InvalidFirstChars, 171, 172
            .Add 174, CStr(174)
            .Add 187, CStr(187)
        End With
    End If
    
    IsValidFirstChar = (Not IsExistInCollection(InvalidFirstChars, CStr(Asc(GivenChar))))

End Function

' Get the valid characters from the given name starting from the second character.
Public Function GetValidCharForSecondToOnward(ByVal GivenName As String, KeepSpace As Boolean) As String

    Dim Result As String
    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' Check if the current character is a valid second character for the name.
        If IsValidSecondChar(CurrentChar) Or (KeepSpace And CurrentChar = ONE_SPACE) Then
            Result = Result & CurrentChar
        End If
    Next CurrentCharIndex

    ' Return the result containing valid characters.
    GetValidCharForSecondToOnward = Result

End Function

Public Function IsValidSecondChar(ByVal GivenChar As String) As Boolean
    
    Static InvalidChars As Collection
    If InvalidChars Is Nothing Then
        Set InvalidChars = New Collection
        AddCharsToColl InvalidChars, 1, 45
        InvalidChars.Add 47, CStr(47)
        AddCharsToColl InvalidChars, 58, 62
        With InvalidChars
            .Add 64, CStr(64)
            .Add 91, CStr(91)
            AddCharsToColl InvalidChars, 93, 94
            .Add 96, CStr(96)
            AddCharsToColl InvalidChars, 123, 127
            AddCharsToColl InvalidChars, 129, 130
            .Add 132, CStr(132)
            .Add 139, CStr(139)
            .Add 141, CStr(141)
            AddCharsToColl InvalidChars, 143, 144
            .Add 149, CStr(149)
            .Add 155, CStr(155)
            .Add 157, CStr(157)
            .Add 160, CStr(160)
            AddCharsToColl InvalidChars, 162, 163
            AddCharsToColl InvalidChars, 165, 166
            .Add 169, CStr(169)
            AddCharsToColl InvalidChars, 171, 172
            .Add 174, CStr(174)
            .Add 187, CStr(187)
        End With
    End If
    
    IsValidSecondChar = (Not IsExistInCollection(InvalidChars, CStr(Asc(GivenChar))))

End Function

Private Sub AddCharsToColl(ByRef ToColl As Collection, ByVal StartCodeIndex As Long, ByVal EndCodeIndex As Long)
    
    Dim CodeIndex As Long
    For CodeIndex = StartCodeIndex To EndCodeIndex
        ToColl.Add CodeIndex, CStr(CodeIndex)
    Next CodeIndex
    
End Sub

Public Function ConcatenateCollection(ByVal GivenCollection As Collection _
                                      , Optional ByVal Delimiter As String = ",") As String
    
    Dim Result As String
    Dim CurrentItem As Variant
    For Each CurrentItem In GivenCollection
        Result = Result & CStr(CurrentItem) & Delimiter
    Next CurrentItem
    
    If Result = vbNullString Then
        ConcatenateCollection = vbNullString
    Else
        ConcatenateCollection = Left$(Result, Len(Result) - Len(Delimiter))
    End If
    
End Function

Public Function PutSpaceOnLowerCaseToUpperCaseTransition(ByVal CurrentWord As String) As String
    
    Dim Result As String
    Dim Index As Long
    Dim CurrentCharacter As String
    Dim NextCharacter As String
    For Index = 1 To Len(CurrentWord) - 1
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        NextCharacter = Mid$(CurrentWord, Index + 1, 1)
        Result = Result & CurrentCharacter
        If Not IsCapitalLetter(CurrentCharacter) And IsAlphabet(CurrentCharacter) _
           And IsCapitalLetter(NextCharacter) Then
            Result = Result & Space(1)
        End If
    Next Index
    If CurrentWord <> vbNullString Then Result = Result & Right$(CurrentWord, 1)
    PutSpaceOnLowerCaseToUpperCaseTransition = Result
    
End Function

'PutSpaceBeforeLastCapsFromStart("CASE%$Rules") >> "CASE%$ Rules"
Public Function PutSpaceBeforeLastCapsFromStart(ByVal CurrentWord As String) As String
    
    If CurrentWord = vbNullString Then Exit Function
    If IsAllCaps(CurrentWord) Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If
    
    
    Dim Index As Long
    Dim CurrentCharacter As String
    If Not IsCapitalLetter(Left$(CurrentWord, 1)) Then
        PutSpaceBeforeLastCapsFromStart = CurrentWord
        Exit Function
    End If
    
    Dim LowerCaseCharIndex As Long
    For Index = 2 To Len(CurrentWord)
        CurrentCharacter = Mid$(CurrentWord, Index, 1)
        If Not IsCapitalLetter(CurrentCharacter) And IsAlphabet(CurrentCharacter) Then
            LowerCaseCharIndex = Index
            Exit For
        End If
    Next Index
    
    Dim Result As String
    If LowerCaseCharIndex < 3 Then
        Result = CurrentWord
    Else
        Result = Left(CurrentWord, LowerCaseCharIndex - 2) _
                 & ONE_SPACE & Mid(CurrentWord, LowerCaseCharIndex - 1)
    End If
    
    PutSpaceBeforeLastCapsFromStart = Result
    
End Function

Public Function IsCapitalLetter(ByVal GivenLetter As String) As Boolean
    If Len(GivenLetter) > 1 Then
        Err.Raise 13, "IsCapitalLetter Function", "Given Letter need to be one character String"
    End If
    If GivenLetter = vbNullString Then
        Err.Raise 5, "IsCapitalLetter Function", "Given Letter can't be nullstring"
    End If

    Const ASCII_CODE_FOR_A As Integer = 65
    Const ASCII_CODE_FOR_Z As Integer = 90
    Dim ASCIICodeForGivenLetter As Integer
    ASCIICodeForGivenLetter = Asc(GivenLetter)
    IsCapitalLetter = (ASCIICodeForGivenLetter >= ASCII_CODE_FOR_A _
                       And ASCIICodeForGivenLetter <= ASCII_CODE_FOR_Z)

End Function

Public Function IsAlphabet(Char As String) As Boolean
    
    Dim CharCode As Long
    CharCode = Asc(LCase(Char))
    IsAlphabet = (CharCode >= Asc("a") And CharCode <= Asc("z"))
    
End Function

Public Function CapitalizeFirstCharOfEachWord(ByVal GivenName As String) As String

    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' If the current character is a space (ASCII code 32),
        ' capitalize the first char follows it.
        Const SPACE_ASCII_VALUE As Long = 32
        If Asc(CurrentChar) = SPACE_ASCII_VALUE Then
            If CurrentCharIndex < Len(GivenName) Then GivenName = CapitalizeNthCharacter(GivenName _
                                                                                         , CurrentCharIndex + 1)
        End If
    Next CurrentCharIndex
    CapitalizeFirstCharOfEachWord = CapitalizeNthCharacter(GivenName, 1)
    
End Function

' Check if Upper case text and input text is equal or not.
Public Function IsAllCaps(Text As String) As Boolean
    IsAllCaps = (UCase$(Text) = Text)
End Function

' Convert To proper case only if the entire word is not Upper Case
' Example ConvertToProperCaseOfEachWord("USA is a deveLoped Coutry") >> USA Is A Developed Coutry
Public Function ConvertToProperCaseOfEachWord(ByVal Sentence As String) As String
    
    Dim Words As Variant
    Words = Split(Sentence, ONE_SPACE)
    Dim CurrentIndex As Long
    For CurrentIndex = LBound(Words) To UBound(Words)
        Dim CurrentWord As String
        CurrentWord = Words(CurrentIndex)
        If IsAllCaps(CurrentWord) Then
            Words(CurrentIndex) = CurrentWord
        Else
            Words(CurrentIndex) = Text.Proper(CurrentWord)
        End If
    Next CurrentIndex
    ConvertToProperCaseOfEachWord = Join(Words, ONE_SPACE)
    
End Function

' Capitalize first character of each word in the given name that follows a line break.
Public Function CapitalizeFirstCharOfEachWordAfterLineBreak(ByVal GivenName As String) As String

    Dim CurrentCharIndex As Long
    Dim CurrentChar As String

    ' Loop through each character in the given name.
    For CurrentCharIndex = 1 To Len(GivenName)
        CurrentChar = Mid$(GivenName, CurrentCharIndex, 1)

        ' If the current character is a line break (ASCII code 10),
        ' capitalize the first char that follows it.
        If Asc(CurrentChar) = 10 Then
            If CurrentCharIndex < Len(GivenName) Then GivenName = CapitalizeNthCharacter(GivenName _
                                                                                         , CurrentCharIndex + 1)
        End If
    Next CurrentCharIndex

    ' Return the modified name.
    CapitalizeFirstCharOfEachWordAfterLineBreak = GivenName

End Function

' Capitalize the Nth character in the given text.
Public Function CapitalizeNthCharacter(ByRef GivenText As String, ByVal NthIndex As Long) As String

    Dim TextLength As Long
    TextLength = Len(GivenText)

    ' Check if the text is empty.
    If TextLength = 0 Then
        CapitalizeNthCharacter = GivenText
        Exit Function
    End If

    ' Check if the NthIndex is valid.
    If NthIndex > TextLength Then
        Err.Raise 13, "Type Mismatch", "NthIndex needs to be less than text length"
    End If

    ' Capitalize the Nth character based on its position.
    If NthIndex = TextLength Then
        CapitalizeNthCharacter = Left$(GivenText, TextLength - 1) & UCase$(Right$(GivenText, 1))
    ElseIf NthIndex = 1 Then
        CapitalizeNthCharacter = UCase$(Left$(GivenText, 1)) & Right$(GivenText, TextLength - 1)
    Else
        CapitalizeNthCharacter = Left$(GivenText, NthIndex - 1) & UCase$(Mid$(GivenText _
                                                                              , NthIndex, 1)) _
                                 & Text.SubString(GivenText, NthIndex + 1)
    End If

End Function

Public Function IsNoIntersection(ByVal FirstRange As Range, ByVal SecondRange As Range) As Boolean
    IsNoIntersection = IsNothing(FindIntersection(FirstRange, SecondRange))
End Function

Public Function IsSubSet(ByVal ParentSet As Range, ByVal ChildSet As Range) As Boolean
    
    ' Check if ChildSet is a subset of ParentSet.
    Dim CommonSection As Range
    Set CommonSection = FindIntersection(ParentSet, ChildSet)
    If IsNothing(CommonSection) Then
        IsSubSet = False
    ElseIf IsNotNothing(ChildSet) Then
        IsSubSet = (CommonSection.Address = ChildSet.Address)
    End If
    
End Function

Public Function ExtractNameFromLocalNameRange(ByVal LocalName As String) As String
    
    ' Extracts the name from a local named range.
    Dim Result As String
    If Text.Contains(LocalName, SHEET_NAME_SEPARATOR) Then
        Result = Text.AfterDelimiter(LocalName, SHEET_NAME_SEPARATOR, , FROM_END)
    Else
        Result = LocalName
    End If
    
    ExtractNameFromLocalNameRange = Result
    
End Function

Public Function IsAllCellBlank(ByVal NeededRange As Range) As Boolean
    ' Checks if all cells in the specified range are blank.
    IsAllCellBlank = (Application.WorksheetFunction.CountA(NeededRange) = 0)
End Function

Public Sub ScrollToDependencyDataRange(ByVal Table As ListObject)
    
    ' Scrolls to the dependency data range in the specified table.
    Application.GoTo Table.Range, True
    Table.Range(1, 1).Select
    
End Sub

Public Function RemoveTopRowHeader(ByVal InputArray As Variant) As Variant

    ' Check if the input is an array
    If Not IsArrayAllocated(InputArray) Then
        RemoveTopRowHeader = InputArray
        Exit Function
    End If

    ' Declare variable for number of rows
    Dim NumRows As Long
    Dim CurrentRow As Long
    Dim CurrentCol As Long

    ' Get the number of rows in the input array
    NumRows = UBound(InputArray, 1) - LBound(InputArray, 1) + 1

    ' Check if the input array has more than one row
    If NumRows <= 1 Then
        RemoveTopRowHeader = Empty
        Exit Function
    End If

    ' Declare a result array without the top row, using the same lower bounds as the input array
    Dim ResultArray() As Variant
    ReDim ResultArray(LBound(InputArray, 1) To UBound(InputArray, 1) - 1 _
                      , LBound(InputArray, 2) To UBound(InputArray, 2))

    ' Copy the content without the top row
    For CurrentRow = LBound(InputArray, 1) + 1 To UBound(InputArray, 1)
        For CurrentCol = LBound(InputArray, 2) To UBound(InputArray, 2)
            ResultArray(CurrentRow - 1, CurrentCol) = InputArray(CurrentRow, CurrentCol)
        Next CurrentCol
    Next CurrentRow

    ' Assign the result to the function's return value
    RemoveTopRowHeader = ResultArray

End Function

Public Function IsArrayAllocated(ByVal Arr As Variant) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsArrayAllocated
    ' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
    ' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
    ' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
    ' allocated.
    '
    ' The VBA IsArray function indicates whether a variable is an array, but it does not
    ' distinguish between allocated and unallocated arrays. It will return TRUE for both
    ' allocated and unallocated arrays. This function tests whether the array has actually
    ' been allocated.
    '
    ' This function is just the reverse of IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim N As Long
    On Error Resume Next

    ' if Arr is not an array, return FALSE and get out.
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occurred.
    N = UBound(Arr, 1)
    If (Err.Number = 0) Then
        ''''''''''''''''''''''''''''''''''''''
        ' Under some circumstances, if an array
        ' is not allocated, Err.Number will be
        ' 0. To acccomodate this case, we test
        ' whether LBound <= Ubound. If this
        ' is True, the array is allocated. Otherwise,
        ' the array is not allocated.
        '''''''''''''''''''''''''''''''''''''''
        If LBound(Arr) <= UBound(Arr) Then
            ' no error. array has been allocated.
            IsArrayAllocated = True
        Else
            IsArrayAllocated = False
        End If
    Else
        ' error. unallocated array
        IsArrayAllocated = False
    End If

End Function

Public Sub AssignFormulaIfErrorPrintIntoDebugWindow(ByVal PutFormulaOnCell As Range _
                                                    , ByVal FormulaText As String _
                                                     , Optional ByVal Message As String = vbNullString)
    
    ' Assigns a formula to the specified cell and prints the formula into the debug window if an error occurs.
    On Error GoTo PrintFormulaToDebugWindow
    PutFormulaOnCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(FormulaText)
    Exit Sub

PrintFormulaToDebugWindow:
    Debug.Print Message & FormulaText
    
End Sub

Public Function RemoveDollarSign(ByVal RangeAddress As String) As String
    RemoveDollarSign = VBA.Replace(RangeAddress, DOLLAR_SIGN, vbNullString)
End Function

Public Function IsTableExist(ByVal InBook As Workbook, ByVal TableName As String) As Boolean
    
    Dim Sheet As Worksheet
    For Each Sheet In InBook.Worksheets
        Dim Table As ListObject
        For Each Table In Sheet.ListObjects
            If Table.Name = TableName Then
                IsTableExist = True
                Exit Function
            End If
        Next Table
    Next Sheet
    
End Function

Public Function IsBlankCellNoError(ByVal CheckCell As Range) As Boolean
    
    If CheckCell.Cells.Count = 1 Then
        If IsError(CheckCell.Value) Then
            IsBlankCellNoError = False
        ElseIf CheckCell.Value = vbNullString Then
            IsBlankCellNoError = True
        End If
    End If
    
End Function

Public Function IsArrayOfNullString(ByVal InputArr As Variant) As Boolean
    
    IsArrayOfNullString = True
    Dim CurrentElement As Variant
    For Each CurrentElement In InputArr
        If CurrentElement <> vbNullString Then
            IsArrayOfNullString = False
            Exit Function
        End If
    Next CurrentElement
    
End Function

Public Function GetSpillParent(ByVal AnyCellInsideSpill As Range) As Range
    If AnyCellInsideSpill.HasSpill Then
        Set GetSpillParent = AnyCellInsideSpill.Cells(1).SpillParent
    End If
End Function

Public Function GetSpillRange(ByVal AnyCellInsideSpill As Range) As Range
    
    If AnyCellInsideSpill.HasSpill Then
        Set GetSpillRange = AnyCellInsideSpill.Cells(1).SpillParent.SpillingToRange
    End If

End Function

Public Function IsValidDefinedName(ByVal NameToCheck As String) As Boolean
    IsValidDefinedName = (NameToCheck = RemoveInvalidCharcters(NameToCheck, False))
End Function

Public Function FilterUsingSpecialCells(ByVal FromRange As Range _
                                        , ByVal CellType As XlCellType) As Range
    
    Set FilterUsingSpecialCells = Intersect(FromRange, FromRange.SpecialCells(CellType))
    
End Function

Public Function ConvertToFullyQualifiedCellRef(ByVal ForCell As Range) As String
    
    ' Converts a cell reference to a fully qualified cell reference with book name and sheet names.
    ' Example output: '[Different Locale Functions Map.xlsm]Keywords Locale Map'!$H$6
    
    ConvertToFullyQualifiedCellRef = SINGLE_QUOTE & LEFT_SQUARE_BRACKET & WorkbookNameFromRange(ForCell) _
                                     & RIGHT_SQUARE_BRACKET & Replace(ForCell.Worksheet.Name, SINGLE_QUOTE, SINGLE_QUOTE & SINGLE_QUOTE) _
                                     & SINGLE_QUOTE & SHEET_NAME_SEPARATOR & ForCell.Address
                                     
End Function

Public Function GetRangeReference(ByVal GivenCells As Range _
                                  , Optional ByVal IsAbsolute As Boolean = True) As String
    
    ' Retrieves the reference of the given range as a string.

    GetRangeReference = GivenCells.Address(IsAbsolute, IsAbsolute)

    ' Check if the given range is part of a dynamic array formula.
    If GivenCells.Cells.Count > 1 And GivenCells.Cells(1, 1).HasSpill Then
        Dim TempRange As Range
        Set TempRange = GivenCells.Cells(1, 1)

        ' If it is a spill range, append the dynamic cell reference sign to the range reference.
        If TempRange.SpillParent.SpillingToRange.Address = GivenCells.Address Then
            GetRangeReference = TempRange.SpillParent.Address(IsAbsolute, IsAbsolute) & DYNAMIC_CELL_REFERENCE_SIGN
        End If
    End If
    
End Function

Public Function GetSheetRefForRangeReference(ByVal SheetName As String _
                                             , Optional ByVal IsSingleQuoteMandatory As Boolean = False) As String
    
    ' Returns the sheet reference for the range reference.
    Dim IsSingleQuoteNeeded As Boolean
    If IsSingleQuoteMandatory Then
        IsSingleQuoteNeeded = True
    Else
        IsSingleQuoteNeeded = IsAnyNonAlphanumeric(SheetName)
    End If
    
    Dim Result As String
    If IsSingleQuoteNeeded Then
        ' for single quote we need to escape with double single quote
        Result = SINGLE_QUOTE _
               & Replace(SheetName, SINGLE_QUOTE, SINGLE_QUOTE & SINGLE_QUOTE) _
               & SINGLE_QUOTE & SHEET_NAME_SEPARATOR
    Else
        Result = SheetName & SHEET_NAME_SEPARATOR
    End If
    
    GetSheetRefForRangeReference = Result
    
End Function

Public Function IsAnyNonAlphanumeric(ByVal Text As String) As String
    
    Dim Result As Boolean
    Dim Index As Long
    Dim CurrentCharacter As String
    For Index = 1 To Len(Text)
        CurrentCharacter = Mid(Text, Index, 1)
        If Not CurrentCharacter Like "[A-Za-z0-9]" Then
            Result = True
            Exit For
        End If
    Next Index
    
    IsAnyNonAlphanumeric = Result
    
End Function

Public Function RemoveSheetQualifierFromSheetQualifiedRangeRef(ByVal SheetQualifiedRef As String) As String
    
    ' e.g:  RemoveSheetQualifierFromSheetQualifiedRangeRef("'All Functions Name'!$C$5") >> $C$5
    
    If Not Text.Contains(SheetQualifiedRef, SHEET_NAME_SEPARATOR) Then
        Err.Raise 13, "Invalid Ref", "Input should have sheet name as well."
    End If
    
    RemoveSheetQualifierFromSheetQualifiedRangeRef = Text.AfterDelimiter(SheetQualifiedRef, SHEET_NAME_SEPARATOR, , FROM_END)
    
End Function

Public Function GetRangeRefWithSheetName(ByVal GivenRange As Range _
                                         , Optional ByVal IsSingleQuoteMandatory As Boolean = False _
                                          , Optional ByVal IsAbsolute As Boolean = True) As String
    
    ' Returns the reference of the given range with the sheet name.
    ' If IsAbsolute is True, the reference is absolute; otherwise, it's relative.
    Dim SheetRef As Worksheet
    Set SheetRef = GivenRange.Parent
    GetRangeRefWithSheetName = GetSheetRefForRangeReference(SheetRef.Name, IsSingleQuoteMandatory) _
                               & GetRangeReference(GivenRange, IsAbsolute)
                               
End Function

Public Function WorkbookNameFromRange(ByVal FromRange As Range) As String
    WorkbookNameFromRange = FromRange.Worksheet.Parent.Name
End Function

Public Function ConvertSpillRangeDependencyToAbsRef(ByVal FormulaCell As Range) As String
    
    ' This will convert spill range reference to absolute form.
    ' For example if the formula is =SUM(FILTER(Q183#,P183#=S183)) then it will convert it to
    ' =SUM(FILTER($Q$183#,$P$183#=S183))
    
    Dim DirectPrecedents As Variant
    DirectPrecedents = modCOMWrapper.GetDirectPrecedents(FormulaCell.Formula2, FormulaCell.Worksheet)
    
    If Not IsArray(DirectPrecedents) Then
        ConvertSpillRangeDependencyToAbsRef = FormulaCell.Formula2
        Exit Function
    End If
    
    Dim Result As String
    Result = FormulaCell.Formula2
    
    Dim CurrentPrecedency As Variant
    For Each CurrentPrecedency In DirectPrecedents
        
        Dim PrecedentCellAsText As String
        PrecedentCellAsText = CStr(CurrentPrecedency)
        
        Dim CurrentRange As Range
        If Text.IsEndsWith(PrecedentCellAsText, HASH_SIGN) Then
            
            Set CurrentRange = RangeResolver.GetRangeForDependency(PrecedentCellAsText, FormulaCell)
            Dim ReplaceWith As String
            ReplaceWith = vbNullString
            If CurrentRange.Worksheet.Name <> FormulaCell.Worksheet.Name Then
                ReplaceWith = GetSheetRefForRangeReference(CurrentRange.Worksheet.Name, True)
            End If
            
            ReplaceWith = ReplaceWith & CurrentRange.Cells(1).Address & HASH_SIGN
            
            Result = ReplaceTokenWithNewToken(Result, PrecedentCellAsText, ReplaceWith)
            
        End If
        
    Next CurrentPrecedency
    
    ConvertSpillRangeDependencyToAbsRef = Result
    
End Function

Public Function IsSpilledRangeRef(ByVal RangeReference As String) As Boolean
    IsSpilledRangeRef = Text.IsEndsWith(RangeReference, DYNAMIC_CELL_REFERENCE_SIGN)
End Function

Public Function IsClosedWorkbookRef(ByVal RangeRef As String) As String
    
    Dim Result As Boolean
    ' One drive or share point location.
    ' example: 'https://d.docs.live.net/6edd704b8f8c537b/TextOffset lambda testing.xlsm'!TestName
    If Text.IsStartsWith(RangeRef, "'https://") Then
        Result = True
    ElseIf Text.Contains(RangeRef, ":\") Then
        ' local drive location:
        ' example: 'D:\Downloads\Email Manager V10.xlsm'!TemplateEmailFilePath
        Result = True
    Else
        Result = False
    End If
    
    IsClosedWorkbookRef = Result
    
End Function

Public Function IsSheetExist(ByVal SheetTabName As String _
                             , Optional ByVal GivenWorkbook As Workbook) As Boolean

    '@Description("This function will determine if a sheet is exist or not by using tab name")
    '@Dependency("No Dependency")
    '@ExampleCall : IsSheetExist("SheetTabName")
    '@Date : 14 October 2021 07:03:05 PM

    If GivenWorkbook Is Nothing Then Set GivenWorkbook = ThisWorkbook

    Dim TemporarySheet As Worksheet
    On Error Resume Next
    Set TemporarySheet = GivenWorkbook.Worksheets(SheetTabName)

    IsSheetExist = (Not TemporarySheet Is Nothing)
    On Error GoTo 0

End Function

Public Function IsOpenWorkbookExists(ByVal BookName As String) As Boolean
    
    
    Dim Result As Boolean
    Dim CurrentBook As Workbook
    For Each CurrentBook In Application.Workbooks
        If CurrentBook.Name = BookName Then
            Result = True
            Exit For
        End If
    Next CurrentBook
    
    If Result Then
        IsOpenWorkbookExists = Result
        Exit Function
    End If
    
    Dim CurrentAddIn As AddIn
    For Each CurrentAddIn In Application.AddIns
        If CurrentAddIn.IsOpen And CurrentAddIn.Name = BookName Then
            Result = True
            Exit For
        End If
    Next CurrentAddIn
    
    IsOpenWorkbookExists = Result
    
End Function

Public Function IsLocalScopedNamedRangeExist(ScopeSheet As Worksheet _
                                             , NamedRangeName As String) As Boolean
    
    Dim SheetQualifiedName As String
    SheetQualifiedName = NamedRangeName
    If Not Text.Contains(NamedRangeName, SHEET_NAME_SEPARATOR) Then
        SheetQualifiedName = GetSheetRefForRangeReference(ScopeSheet.Name, False) & NamedRangeName
    End If
    
    Dim CurrentName As Name
    For Each CurrentName In ScopeSheet.Names
        If CurrentName.Name = SheetQualifiedName Then
            IsLocalScopedNamedRangeExist = True
            Exit Function
        End If
    Next CurrentName
    
    IsLocalScopedNamedRangeExist = False
    
End Function

Public Function RemoveSheetQualifierIfPresent(ByVal RangeRef As String) As String
    
    ' e.g:  RemoveSheetQualifierFromSheetQualifiedRangeRef("'All Functions Name'!$C$5") >> $C$5
    
    Dim Result As String
    
    If Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        Result = Text.AfterDelimiter(RangeRef, SHEET_NAME_SEPARATOR, , FROM_END)
    Else
        Result = RangeRef
    End If
    
    RemoveSheetQualifierIfPresent = Result
    
End Function

Public Function Max(ByVal FirstNumber As Variant, ByVal SecondNumber As Variant) As Variant
    Max = Application.WorksheetFunction.Max(FirstNumber, SecondNumber)
End Function

Public Function IsSubRange(ByVal ParentRange As Range _
                           , ByVal ChildRange As Range) As Boolean

    If ChildRange Is Nothing Then Exit Function
    If ParentRange Is Nothing Then Exit Function

    Dim InterSectionRange As Range
    Set InterSectionRange = Intersect(ParentRange, ChildRange)

    If InterSectionRange Is Nothing Then Exit Function
    IsSubRange = (ChildRange.Address = InterSectionRange.Address)

End Function

Public Function MakeAbsoluteReference(ByVal RangeAddress As String _
                                        , ByVal HelperCell As Range) As String
    
    ' Convert the given range address to an absolute reference.

    Dim CurrentRange As Range
    Set CurrentRange = RangeResolver.GetRangeForDependency(RangeAddress, HelperCell)
    
    ' In case of cell address, It can be just simple cell address, spill range cell address or sheet qualified.
    ' In all case it must have the first cell address in it's range ref. If not then bail out.
    
    Dim Result As String
    If IsNothing(CurrentRange) Then
        Result = RangeAddress
    ElseIf Not IsReferredByRangeAddress(RangeAddress, CurrentRange) Then
        Result = RangeAddress
    ' Check if the parent worksheet of the current range is the same as the parent worksheet of the helper cell.
    ElseIf CurrentRange.Worksheet.Name = HelperCell.Worksheet.Name Then
        ' If they are the same, return the range reference with absolute references.
        Result = GetRangeReference(CurrentRange, True)
    Else
        ' If they are different, return the range reference with absolute references and sheet name.
        Result = GetRangeRefWithSheetName(CurrentRange, , True)
    End If
    
    MakeAbsoluteReference = Result
    
End Function

Private Function IsReferredByRangeAddress(ByVal RangeRef As String, ByVal ResolvedRange As Range) As Boolean
    
    If Text.IsEndsWith(RangeRef, HASH_SIGN) Then
        RangeRef = Text.RemoveFromEnd(RangeRef, Len(HASH_SIGN))
        Set ResolvedRange = ResolvedRange.Cells(1)
    End If
    
    If Text.Contains(RangeRef, ResolvedRange.Worksheet.Name) And Text.Contains(RangeRef, SHEET_NAME_SEPARATOR) Then
        RangeRef = Text.AfterDelimiter(RangeRef, SHEET_NAME_SEPARATOR)
    End If
    
    RangeRef = RemoveDollarSign(RangeRef)
    
    IsReferredByRangeAddress = (RangeRef = ResolvedRange.Address(False, False))
    
End Function

Public Function ReplaceInvalidCharFromFormulaWithValid(ByVal Formula As String) As String
    
    Dim Result As String
    Result = Replace(Formula, vbCrLf, vbLf)
    Result = Replace(Result, Chr(160), Chr(32))
    
    ReplaceInvalidCharFromFormulaWithValid = Result
    
End Function
