Attribute VB_Name = "modRemoveOuterFunction"
'@IgnoreModule UndeclaredVariable
'@Folder "RemoveOuterFunction"
Option Explicit
Option Private Module

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Remove Outer Function
' Description:            Remove outer function.
' Macro Expression:       modRemoveOuterFunction.RemoveOuterFunction([ActiveCell],[ActiveCell.Offset(0,1)])
' Generated:              03/30/2023 07:44 AM
'----------------------------------------------------------------------------------------------------
Public Sub RemoveOuterFunction(ByVal GivenCell As Range)
    
    ' Check if the GivenCell is valid and contains a formula, exit otherwise
    If IsNothing(GivenCell) Then Exit Sub
    
    Dim FormulaCells As Range
    On Error Resume Next
    Set FormulaCells = FilterUsingSpecialCells(GivenCell, xlCellTypeFormulas)
    On Error GoTo 0
    If IsNothing(FormulaCells) Then Exit Sub
    
    On Error GoTo ErrorHandler                   ' Initiate error handling block
    Dim CurrentCell As Range
    For Each CurrentCell In FormulaCells.Cells
        RemoveOuterFXFromCurrentCell CurrentCell
    Next CurrentCell
    Exit Sub

ErrorHandler:
    ' Error handling block
    Dim ErrorNumber As Long
    ErrorNumber = Err.Number
    Dim ErrorDescription As String
    ErrorDescription = Err.Description

    ' Raise an error if the ErrorNumber is not zero
    If ErrorNumber <> 0 Then
        Err.Raise ErrorNumber, Err.Source, ErrorDescription
        ' Resume the execution for debugging purposes.
        Resume
    End If

End Sub

Private Sub RemoveOuterFXFromCurrentCell(ByVal CurrentCell As Range)
    
    If IsNothing(CurrentCell) Then Exit Sub
    If Not CurrentCell.HasFormula Then Exit Sub

    ' Compute and store the result of the outer function of the formula in CurrentCell
    Dim ResultFormula As String
    ResultFormula = RemoveOuterFunctionFromFormula(CurrentCell.Formula2)
    ' Print the ResultFormula for debugging purposes
    Logger.Log DEBUG_LOG, ResultFormula
    ' If the ResultFormula is different from the original formula in CurrentCell, update the formula in OutputCell and activate it if it was active
    If ResultFormula <> CurrentCell.Formula2 Then
        CurrentCell.Formula2 = ReplaceInvalidCharFromFormulaWithValid(ResultFormula)
    End If
    
End Sub

Private Sub Test()

    Dim StartFormula As String
    StartFormula = ActiveCell.Formula2
'    StartFormula = FormatFormula(StartFormula)
    Do While True
        Dim AfterRemoved As String
        AfterRemoved = RemoveOuterFunctionFromFormula(StartFormula)
        If StartFormula = AfterRemoved Then Exit Do
        Logger.Log DEBUG_LOG, "Before Removing : " & StartFormula
        Logger.Log DEBUG_LOG, "After Removed : " & AfterRemoved
        Logger.Log DEBUG_LOG, vbNewLine
        StartFormula = AfterRemoved
    Loop

End Sub

'@MethodDescription(Remove outer function once and generate the result formula)
Public Function RemoveOuterFunctionFromFormula(ByVal FormulaText As String) As String

    Dim AfterRemoved As String
    If IsLambdaFunction(FormulaText) Then
        AfterRemoved = RemoveOuterFromLambda(FormulaText)
    ElseIf IsLetFunction(FormulaText) Then
        AfterRemoved = RemoveOuterFromLet(FormulaText)
    Else
        ' Base Case
        Dim OuterRemoved As String
        OuterRemoved = GetDependencyFunctionResult(FormulaText _
                                                   , FIRST_ARGUMENT_OF_OUTER_FUNCTION, False)
        ' If there is no function then it will return empty string like A2+A3
        If OuterRemoved = vbNullString Then OuterRemoved = FormulaText
        AfterRemoved = OuterRemoved
    End If
    AfterRemoved = FormatFormula(AfterRemoved)
    RemoveOuterFunctionFromFormula = AfterRemoved

End Function

Private Sub TestRemoveOuterFromLet()

    Dim TestSetupRange As Range
    Set TestSetupRange = Range("RemoveOuterFromLet")
    Dim Counter As Long
    For Counter = 1 To TestSetupRange.Rows.Count
        RunRemoveOuterFromLetOrLambdaTest TestSetupRange.Cells(Counter, 1).Formula2 _
                                          , TestSetupRange.Cells(Counter, 2).Formula2, True
    Next Counter

End Sub

Private Sub RunRemoveOuterFromLetOrLambdaTest(ByVal TestFormula As String _
                                              , ByVal ExpectedFormula As String _
                                              , ByVal IsRemoveLet As Boolean)

    Dim ActualFormula As String
    If IsRemoveLet Then
        ActualFormula = RemoveOuterFromLet(TestFormula)
    Else
        ActualFormula = RemoveOuterFromLambda(TestFormula)
    End If

    ActualFormula = FormatFormula(ActualFormula)
    ExpectedFormula = FormatFormula(ExpectedFormula)
    If ActualFormula <> ExpectedFormula Then
        Logger.Log DEBUG_LOG, "Test Formula : " & TestFormula
        Logger.Log DEBUG_LOG, "Actual Formula : " & ActualFormula
        Logger.Log DEBUG_LOG, "Expected Formula : " & ExpectedFormula
        Logger.Log DEBUG_LOG, "Test Result : " & IIf(ActualFormula = ExpectedFormula, "Pass", "Fail")
        Logger.Log DEBUG_LOG, vbNewLine
    End If

End Sub

'@Recursive
Private Function RemoveOuterFromLet(ByVal LetFormula As String _
                                    , Optional ByVal Prefix As String = EQUAL_SIGN _
                                     , Optional ByVal Suffix As String = vbNullString) As String

    Dim RemovedFormula As String
    If Not IsLetFunction(LetFormula) Then
        RemovedFormula = RemoveOuterOnNullStringReturnBack(LetFormula)
        RemoveOuterFromLet = Prefix & Text.RemoveFromStartIfPresent(RemovedFormula, EQUAL_SIGN) & Suffix
        Exit Function
    End If
    Suffix = GetLetFormulaInvocation(LetFormula) & Suffix
    Dim LetParts As Variant
    LetParts = GetDependencyFunctionResult(LetFormula, LET_PARTS)
    Dim LastRowIndex As Long
    LastRowIndex = UBound(LetParts, 1)
    Dim FirstColIndex As Long
    FirstColIndex = LBound(LetParts, 2)
    ' If it just refer to the last step
    Dim ResultStepCalc As String
    ResultStepCalc = LetParts(LastRowIndex, FirstColIndex)
    Dim SecondLastStepCalc As String
    SecondLastStepCalc = LetParts(LastRowIndex - 1, FirstColIndex)
    Dim FirstRowIndex As Long
    FirstRowIndex = LBound(LetParts, 1)
    If ResultStepCalc = SecondLastStepCalc Then

        ' We have a bug here. We need to remove considering that there was just the calculation of that step.
        If LastRowIndex - 1 = FirstRowIndex Then
            RemovedFormula = LetParts(LastRowIndex - 1, FirstColIndex + LET_PARTS_VALUE_COL_INDEX - 1)
        Else
            RemovedFormula = ConcatenateLetParts(LetParts, FirstRowIndex, LastRowIndex - 2, vbNullString)
            RemovedFormula = RemovedFormula _
                             & LetParts(LastRowIndex - 1, FirstColIndex + LET_PARTS_VALUE_COL_INDEX - 1) _
                             & FIRST_PARENTHESIS_CLOSE
        End If
        RemoveOuterFromLet = RemoveOuterFromLet(EQUAL_SIGN & RemovedFormula, Prefix, Suffix)

    Else
        ResultStepCalc = EQUAL_SIGN & ResultStepCalc
        If IsLambdaFunction(ResultStepCalc) Then
            
            Prefix = Prefix & ConcatenateLetParts(LetParts, FirstRowIndex, LastRowIndex - 1, vbNullString)
            Suffix = FIRST_PARENTHESIS_CLOSE & Suffix
            RemoveOuterFromLet = RemoveOuterFromLambda(ResultStepCalc, Prefix, Suffix)
            
        ElseIf IsLetFunction(ResultStepCalc) Then
            
            Prefix = Prefix & ConcatenateLetParts(LetParts, FirstRowIndex, LastRowIndex - 1, vbNullString)
            Suffix = FIRST_PARENTHESIS_CLOSE & Suffix
            RemoveOuterFromLet = RemoveOuterFromLet(ResultStepCalc, Prefix, Suffix)
            
        Else
            
            RemovedFormula = GetDependencyFunctionResult(ResultStepCalc, FIRST_ARGUMENT_OF_OUTER_FUNCTION, False)
            ' Meaning last step is just normal calc. So we have to replace that calculating with last step name.
            If RemovedFormula = vbNullString Then
                RemovedFormula = ConcatenateLetParts(LetParts, FirstRowIndex, LastRowIndex - 1, SecondLastStepCalc)
            Else
                ' But if we can replace then add that at the end of the let calc.
                RemovedFormula = Text.RemoveFromStartIfPresent(RemovedFormula, EQUAL_SIGN)
                RemovedFormula = ConcatenateLetParts(LetParts, FirstRowIndex, LastRowIndex - 1, RemovedFormula)
            End If

            RemoveOuterFromLet = Prefix & RemovedFormula & Suffix
            
        End If
    End If

End Function

Private Function RemoveOuterOnNullStringReturnBack(ByVal Formula As String) As String

    Dim AfterRemoved As String
    AfterRemoved = GetDependencyFunctionResult(Formula, FIRST_ARGUMENT_OF_OUTER_FUNCTION, False)
    If AfterRemoved = vbNullString Then
        RemoveOuterOnNullStringReturnBack = Formula
    Else
        RemoveOuterOnNullStringReturnBack = AfterRemoved
    End If

End Function

Private Sub TestRemoveOuterFromLambda()

    Dim TestSetupRange As Range
    Set TestSetupRange = Range("RemoveOuterFromLambda")
    Dim Counter As Long
    For Counter = 1 To TestSetupRange.Rows.Count
        RunRemoveOuterFromLetOrLambdaTest TestSetupRange.Cells(Counter, 1).Formula2 _
                                          , TestSetupRange.Cells(Counter, 2).Formula2, False
    Next Counter

End Sub

'@Recursive
Private Function RemoveOuterFromLambda(ByVal LambdaFormula As String _
                                    , Optional ByVal Prefix As String = EQUAL_SIGN _
                                     , Optional ByVal Suffix As String = vbNullString) As String

    Dim RemovedFormula As String
    If Not IsLambdaFunction(LambdaFormula) Then
        RemovedFormula = RemoveOuterOnNullStringReturnBack(LambdaFormula)
        RemoveOuterFromLambda = Prefix & Text.RemoveFromStartIfPresent(RemovedFormula, EQUAL_SIGN) & Suffix
        Exit Function
    End If

    Dim LambdaParts As Variant
    LambdaParts = GetDependencyFunctionResult(LambdaFormula, LAMBDA_PARTS)
    Dim LastRowIndex As Long
    LastRowIndex = UBound(LambdaParts, 1)
    Dim FirstColIndex As Long
    FirstColIndex = LBound(LambdaParts, 2)
    
    ' If it just refer to the last step
    Dim ResultStepCalc As String
    ResultStepCalc = EQUAL_SIGN & LambdaParts(LastRowIndex, FirstColIndex)
    Prefix = Prefix & Text.RemoveFromStart(GetUptoLambdaParamDefPart(LambdaFormula), 1)
    Suffix = FIRST_PARENTHESIS_CLOSE & GetLambdaInvocationPart(LambdaFormula) & Suffix
    
    If IsLambdaFunction(ResultStepCalc) Then
        RemoveOuterFromLambda = RemoveOuterFromLambda(ResultStepCalc, Prefix, Suffix)
    ElseIf IsLetFunction(ResultStepCalc) Then
        RemoveOuterFromLambda = RemoveOuterFromLet(ResultStepCalc, Prefix, Suffix)
    Else
        RemovedFormula = GetDependencyFunctionResult(ResultStepCalc, FIRST_ARGUMENT_OF_OUTER_FUNCTION, False)
        If RemovedFormula = vbNullString Then RemovedFormula = ResultStepCalc
        RemovedFormula = Text.RemoveFromStartIfPresent(RemovedFormula, EQUAL_SIGN)
        RemoveOuterFromLambda = Prefix & RemovedFormula & Suffix
    End If

End Function


Private Function ConcatenateLetParts(ByVal LetParts As Variant _
                                     , ByVal StartRowIndex As Long _
                                      , ByVal EndRowIndex As Long _
                                       , Optional ByVal ResultStepName As String = vbNullString _
                                        , Optional ByVal IsLETNeeded As Boolean = True) As String
    Dim RowIndex As Long
    Dim Formula As String
    Dim StepName As String
    Dim StepCalc As String
    Dim FirstColIndex As Long
    FirstColIndex = LBound(LetParts, 2)
    If StartRowIndex >= EndRowIndex Then

        StepCalc = LetParts(StartRowIndex, FirstColIndex + LET_PARTS_VALUE_COL_INDEX - 1)
        If IsLETNeeded Then
            Formula = LET_AND_OPEN_PAREN & LetParts(StartRowIndex, FirstColIndex) _
                      & LIST_SEPARATOR & StepCalc & LIST_SEPARATOR
        Else
            Formula = StepCalc
        End If

        If ResultStepName <> vbNullString Then
            Formula = Formula & ResultStepName & FIRST_PARENTHESIS_CLOSE
        End If
        ConcatenateLetParts = Formula
        Exit Function

    End If

    Formula = LET_AND_OPEN_PAREN

    For RowIndex = StartRowIndex To EndRowIndex
        StepName = LetParts(RowIndex, FirstColIndex)
        StepCalc = LetParts(RowIndex, FirstColIndex + LET_PARTS_VALUE_COL_INDEX - 1)
        Formula = Formula & StepName & LIST_SEPARATOR & StepCalc & LIST_SEPARATOR
    Next RowIndex

    If ResultStepName <> vbNullString Then
        Formula = Formula & ResultStepName & FIRST_PARENTHESIS_CLOSE
    End If

    ConcatenateLetParts = Formula

End Function
