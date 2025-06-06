VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableLookupLambdaGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SIX_SPACE As String = THREE_SPACE & THREE_SPACE
Private Const NINE_SPACE As String = SIX_SPACE & THREE_SPACE
Private Const TWELVE_SPACE As String = SIX_SPACE & SIX_SPACE

Private Type TTableLookupLambdaGenerator
    ColumnListStepName As String
    TableStepName As String
    CriteriaLambdaStepName As String
    DatasStepName As String
    HeadersStepName As String
    ReturnStepName As String
    CriteriaStepName As String
    FilteredDataStepName As String
    SortOrderLambdaStepName As String
    SortLambdaStepName As String
    ReturnColIndexStepName As String
    ResultStepName As String
    DefPart As String
    ParamPart As String
    LookupColumnPart As String
    ColumnIndexPart As String
    CriteriaPart As String
    FinalCriteriaPart As String
    SortPart As String
    FilterLambda As String
End Type
Private This As TTableLookupLambdaGenerator

Public Property Get DefPart() As String
    DefPart = This.DefPart
End Property

Public Property Get FilterLambda() As String
    FilterLambda = Replace(This.FilterLambda, vbNewLine, Chr$(10))
End Property

Private Sub UpdateStepsName(ByVal UniqueStepsName As Collection)

    This.ColumnListStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "column_list")
    This.TableStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Table")
    This.CriteriaLambdaStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "fnCriteria")
    This.DatasStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Data")
    This.HeadersStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Header")
    This.ReturnStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_ReturnColumns")
    This.CriteriaStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Criteria")
    This.FilteredDataStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Filtered")
    This.SortOrderLambdaStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "fnSortOrder")
    This.SortLambdaStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "fnSort")
    This.ReturnColIndexStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_ReturnColumnIndex")
    This.ResultStepName = CreateUniqueNameByAppendingNumber(UniqueStepsName, "_Result")
    
End Sub

Public Sub GenerateTemplateLambda(ByVal IsTable As Boolean _
                                  , ByVal TableOrNamedRangeName As String _
                                   , ByVal ValidHeaders As Range)
    
    Dim UniqueStepsName As Collection
    Set UniqueStepsName = CreateUniqueLetStepNames(ValidHeaders)
    UpdateStepsName UniqueStepsName
    
    Dim TableStepCalc As String
    TableStepCalc = TableOrNamedRangeName & IIf(IsTable, TABLE_ALL_MARKER, vbNullString)
    
    This.SortPart = LAMBDA_AND_OPEN_PAREN & SORTBY_FN_NAME & FIRST_PARENTHESIS_OPEN _
                    & vbNewLine & SIX_SPACE & This.FilteredDataStepName & LIST_SEPARATOR & NEW_LINE
    
    This.FinalCriteriaPart = This.CriteriaStepName & LIST_SEPARATOR _
                             & ONE_SPACE & EXPAND_FN_NAME & FIRST_PARENTHESIS_OPEN
    
    Dim Counter As Long
    Dim CurrentCell As Range
    For Each CurrentCell In ValidHeaders.Cells
        Counter = Counter + 1
        Dim CurrentParamName As String
        CurrentParamName = UniqueStepsName.Item(Counter) & "_Criteria"
        This.DefPart = This.DefPart & "[" & CurrentParamName & "]" & LIST_SEPARATOR
        This.ParamPart = This.ParamPart _
                         & THREE_SPACE & "_Param" & Counter & LIST_SEPARATOR _
                         & ONE_SPACE & CurrentParamName & LIST_SEPARATOR & vbNewLine
        
        Dim EncodedHeader As String
        EncodedHeader = DOUBLE_QUOTE & EscapeDoubleQuote(CurrentCell.Value) & DOUBLE_QUOTE
            
        This.LookupColumnPart = This.LookupColumnPart _
                                & THREE_SPACE & "_LookupColumn" & Counter & LIST_SEPARATOR _
                                & ONE_SPACE & EncodedHeader & LIST_SEPARATOR & vbNewLine
        
        This.ColumnIndexPart = This.ColumnIndexPart _
                               & THREE_SPACE & "_ColumnIndex" & Counter & LIST_SEPARATOR _
                               & ONE_SPACE & XMATCH_FN_NAME & FIRST_PARENTHESIS_OPEN _
                               & "_LookupColumn" & Counter & LIST_SEPARATOR & ONE_SPACE _
                               & This.HeadersStepName & LIST_SEPARATOR & ONE_SPACE & "0" _
                               & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & vbNewLine
        
        This.CriteriaPart = This.CriteriaPart _
                            & THREE_SPACE & "_Criteria" & Counter & LIST_SEPARATOR _
                            & ONE_SPACE & This.CriteriaLambdaStepName & FIRST_PARENTHESIS_OPEN _
                            & CHOOSECOLS_FN_NAME & FIRST_PARENTHESIS_OPEN & This.DatasStepName _
                            & LIST_SEPARATOR & ONE_SPACE & "_ColumnIndex" & Counter _
                            & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & ONE_SPACE _
                            & "_Param" & Counter & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & vbNewLine
        
        This.FinalCriteriaPart = This.FinalCriteriaPart & "_Criteria" & Counter & ONE_SPACE & "*" & ONE_SPACE
        
        This.SortPart = This.SortPart & SIX_SPACE & This.SortOrderLambdaStepName _
                        & FIRST_PARENTHESIS_OPEN & "_ColumnIndex" & Counter _
                        & LIST_SEPARATOR & ONE_SPACE & "_Param" & Counter _
                        & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR _
                        & vbNewLine & SIX_SPACE & "1" & LIST_SEPARATOR & NEW_LINE
        
    Next CurrentCell
    
    This.FinalCriteriaPart = Text.RemoveFromEndIfPresent(This.FinalCriteriaPart, "* ")
    This.SortPart = Text.RemoveFromEndIfPresent(This.SortPart, LIST_SEPARATOR & NEW_LINE)
    
    Dim Concatenator As TextConcatenator
    Set Concatenator = New TextConcatenator
    With Concatenator
        
        .Concatenate EQUAL_SIGN & LAMBDA_AND_OPEN_PAREN & "[" & This.ColumnListStepName & "]"
        .Concatenate LIST_SEPARATOR & This.DefPart & ONE_SPACE & LET_AND_OPEN_PAREN & NEW_LINE
        
        .Concatenate This.ParamPart
        
        .Concatenate THREE_SPACE & This.TableStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate TableStepCalc & LIST_SEPARATOR & NEW_LINE
        
        .Concatenate This.LookupColumnPart
        .Concatenate THREE_SPACE & "LAMBDA_TYPE_VALUE" & LIST_SEPARATOR & ONE_SPACE & "128"
        .Concatenate LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & "ARRAY_TYPE_VALUE" & LIST_SEPARATOR & ONE_SPACE & "64"
        .Concatenate LIST_SEPARATOR & NEW_LINE
        
        .Concatenate THREE_SPACE & This.CriteriaLambdaStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & GetCriteriaLambda() & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.DatasStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate DROP_FN_NAME & FIRST_PARENTHESIS_OPEN & This.TableStepName
        .Concatenate LIST_SEPARATOR & ONE_SPACE & "1" & FIRST_PARENTHESIS_CLOSE
        .Concatenate LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.HeadersStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate CHOOSEROWS_FN_NAME & FIRST_PARENTHESIS_OPEN & This.TableStepName
        .Concatenate LIST_SEPARATOR & ONE_SPACE & "1" & FIRST_PARENTHESIS_CLOSE
        .Concatenate LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.ReturnStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & TOROW_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
                     
        .Concatenate SIX_SPACE & IF_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
        
        .Concatenate NINE_SPACE & ISOMITTED_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.ColumnListStepName & FIRST_PARENTHESIS_CLOSE
        .Concatenate LIST_SEPARATOR & ONE_SPACE & NEW_LINE
                     
        .Concatenate NINE_SPACE & This.HeadersStepName & LIST_SEPARATOR & ONE_SPACE & NEW_LINE
        
        .Concatenate NINE_SPACE & IF_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
        .Concatenate TWELVE_SPACE & AND_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate ISNUMBER_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.ColumnListStepName & FIRST_PARENTHESIS_CLOSE
        .Concatenate FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        
        .Concatenate TWELVE_SPACE & CHOOSECOLS_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.HeadersStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate This.ColumnListStepName & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate TWELVE_SPACE & This.ColumnListStepName & NEW_LINE
        .Concatenate NINE_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate SIX_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate This.ColumnIndexPart
        .Concatenate This.CriteriaPart
        
        .Concatenate THREE_SPACE & This.FinalCriteriaPart & LIST_SEPARATOR & ONE_SPACE
        .Concatenate ROWS_FN_NAME & FIRST_PARENTHESIS_OPEN & This.DatasStepName
        .Concatenate FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
        .Concatenate ONE_SPACE & "1" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate "1" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.FilteredDataStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & FILTER_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.DatasStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate This.CriteriaStepName & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.SortOrderLambdaStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & GetSortOrderLambda() & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & "\\Note" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate """Using thunking for lazy evaluation.""" & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.SortLambdaStepName & LIST_SEPARATOR & ONE_SPACE & This.SortPart & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate THREE_SPACE & This.ReturnColIndexStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate XMATCH_FN_NAME & FIRST_PARENTHESIS_OPEN & This.ReturnStepName
        .Concatenate LIST_SEPARATOR & ONE_SPACE & This.HeadersStepName
        .Concatenate LIST_SEPARATOR & ONE_SPACE & "0" & FIRST_PARENTHESIS_CLOSE
        .Concatenate LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate THREE_SPACE & This.ResultStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & IFERROR_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
                     
        .Concatenate SIX_SPACE & IF_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
        .Concatenate NINE_SPACE & ROWS_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.FilteredDataStepName & FIRST_PARENTHESIS_CLOSE
        .Concatenate ONE_SPACE & "= 1" & LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate NINE_SPACE & VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.ReturnStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate INDEX_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.FilteredDataStepName & LIST_SEPARATOR
        .Concatenate ONE_SPACE & "1" & LIST_SEPARATOR
        .Concatenate ONE_SPACE & This.ReturnColIndexStepName
        .Concatenate FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
        .Concatenate LIST_SEPARATOR & NEW_LINE
                     
        .Concatenate NINE_SPACE & VSTACK_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.ReturnStepName & LIST_SEPARATOR & ONE_SPACE
        .Concatenate CHOOSECOLS_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate This.SortLambdaStepName & FIRST_PARENTHESIS_OPEN
        .Concatenate FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
        .Concatenate ONE_SPACE & This.ReturnColIndexStepName & FIRST_PARENTHESIS_CLOSE
        .Concatenate FIRST_PARENTHESIS_CLOSE & NEW_LINE
                     
        .Concatenate SIX_SPACE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate SIX_SPACE & This.ReturnStepName & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate THREE_SPACE & This.ResultStepName & NEW_LINE
        .Concatenate FIRST_PARENTHESIS_CLOSE & FIRST_PARENTHESIS_CLOSE
    End With
    
    This.FilterLambda = Concatenator.Text
    
End Sub

Private Function GetCriteriaLambda() As String
    
    Dim Concatenator As New TextConcatenator
    With Concatenator
        .Concatenate LAMBDA_AND_OPEN_PAREN & "col_data" & LIST_SEPARATOR
        .Concatenate "[lookup_value]" & LIST_SEPARATOR & NEW_LINE
        .Concatenate SIX_SPACE & IFERROR_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate IF_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
        .Concatenate NINE_SPACE & ISOMITTED_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate "lookup_value" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        
        .Concatenate NINE_SPACE & TRUE_KEYWORD & LIST_SEPARATOR & NEW_LINE
        
        .Concatenate NINE_SPACE & SWITCH_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate TYPE_FN_NAME & FIRST_PARENTHESIS_OPEN & "lookup_value"
        .Concatenate FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate TWELVE_SPACE & "LAMBDA_TYPE_VALUE" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate MAP_FN_NAME & FIRST_PARENTHESIS_OPEN & "col_data" & LIST_SEPARATOR
        .Concatenate ONE_SPACE & LAMBDA_AND_OPEN_PAREN & "val"
        .Concatenate LIST_SEPARATOR & ONE_SPACE & "lookup_value(val)))" & LIST_SEPARATOR & NEW_LINE
        .Concatenate TWELVE_SPACE & "ARRAY_TYPE_VALUE" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate BYROW_FN_NAME & FIRST_PARENTHESIS_OPEN & "col_data = " & TOROW_FN_NAME
        .Concatenate "(lookup_value)" & LIST_SEPARATOR & ONE_SPACE & LAMBDA_AND_OPEN_PAREN
        .Concatenate "row" & LIST_SEPARATOR & ONE_SPACE & OR_FN_NAME & "(row)))" & LIST_SEPARATOR & NEW_LINE
        .Concatenate TWELVE_SPACE & "col_data = lookup_value" & NEW_LINE
        .Concatenate NINE_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate SIX_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate SIX_SPACE & LIST_SEPARATOR & ONE_SPACE
        .Concatenate FALSE_KEYWORD & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE
    End With
    
    GetCriteriaLambda = Concatenator.Text
    
End Function

Private Function GetSortOrderLambda() As String
    
    Dim Concatenator As New TextConcatenator
    With Concatenator
        .Concatenate LAMBDA_AND_OPEN_PAREN & "column_index" & LIST_SEPARATOR
        .Concatenate "[lookup_value]" & LIST_SEPARATOR & NEW_LINE
        .Concatenate SIX_SPACE & IF_FN_NAME & FIRST_PARENTHESIS_OPEN & NEW_LINE
        .Concatenate NINE_SPACE & TYPE_FN_NAME
        .Concatenate FIRST_PARENTHESIS_OPEN & "lookup_value) = ARRAY_TYPE_VALUE" & LIST_SEPARATOR & NEW_LINE
        .Concatenate NINE_SPACE & XMATCH_FN_NAME & FIRST_PARENTHESIS_OPEN
        .Concatenate CHOOSECOLS_FN_NAME & FIRST_PARENTHESIS_OPEN & "_Filtered" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate "column_index" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
        .Concatenate ONE_SPACE & "lookup_value" & LIST_SEPARATOR & ONE_SPACE & "0"
        .Concatenate FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR & NEW_LINE
        .Concatenate NINE_SPACE & EXPAND_FN_NAME & FIRST_PARENTHESIS_OPEN & "1" & LIST_SEPARATOR & ONE_SPACE
        .Concatenate ROWS_FN_NAME & FIRST_PARENTHESIS_OPEN & "_Filtered" & FIRST_PARENTHESIS_CLOSE & LIST_SEPARATOR
        .Concatenate ONE_SPACE & "1" & LIST_SEPARATOR & ONE_SPACE & "1" & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate SIX_SPACE & FIRST_PARENTHESIS_CLOSE & NEW_LINE
        .Concatenate THREE_SPACE & FIRST_PARENTHESIS_CLOSE
    End With
    
    GetSortOrderLambda = Concatenator.Text
    
    '    Debug.Print GetSortOrderLambda
    
End Function


Private Function CreateUniqueLetStepNames(ByVal HeaderCells As Range) As Collection
    
    Dim UniqueStepsName As Collection
    Set UniqueStepsName = New Collection
    
    Dim CurrentCell  As Range
    For Each CurrentCell In HeaderCells.Cells
        
        Dim CurrentParamName As String
        CurrentParamName = MakeValidName(CurrentCell.Value, False)
        CurrentParamName = CreateUniqueNameByAppendingNumber(UniqueStepsName, CurrentParamName)
        UniqueStepsName.Add CurrentParamName, CurrentParamName
        
    Next CurrentCell
    
    Set CreateUniqueLetStepNames = UniqueStepsName
    
End Function

Private Function CreateUniqueNameByAppendingNumber(ByVal AlreadyTakenNames As Collection, ByVal NewName As String) As String
    
    Dim Counter As Long
    Counter = 1
    Do While IsExistInCollection(AlreadyTakenNames, NewName)
                
        Counter = Counter + 1
        If Counter = 2 Then
            NewName = NewName & Counter
        Else
            NewName = Left$(NewName, Len(CStr(Counter - 1))) & Counter
        End If
                
    Loop
    
    CreateUniqueNameByAppendingNumber = NewName
    
End Function

