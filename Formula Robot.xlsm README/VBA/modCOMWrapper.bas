Attribute VB_Name = "modCOMWrapper"
Option Explicit
Option Private Module
#Const DEVELOPMENT_MODE = True
 
Public Enum DependencyFunctions
    LET_PARTS = 1
    LAMBDA_PARTS = 2
    FIRST_ARGUMENT_OF_OUTER_FUNCTION = 3
End Enum

' Find all the dependency workbook is necessary for removing Lambda which is in the name manager
Public Function GetDirectPrecedents(ByVal Formula As String _
                                    , ByVal FormulaInSheet As Worksheet) As Variant
    
    Dim Dependencies As Collection
    Set Dependencies = GetDirectPrecedentsFromExpr(Formula, FormulaInSheet)
    
    Dim Result As Variant
    If Dependencies.Count = 0 Then
        ReDim Result(1 To 1, 1 To 1) As String
        Result(1, 1) = vbNullString
        GetDirectPrecedents = Result
        Exit Function
    End If
    
    Dim Lambdas As Collection
    Set Lambdas = FindLambdas(FormulaInSheet.Parent)
    Dim ValidDependencies As Collection
    Set ValidDependencies = New Collection
    Dim CurrentDependency As Variant
    Dim QualifiedSheetName As String
    QualifiedSheetName = GetSheetRefForRangeReference(FormulaInSheet.Name, False)
    
    For Each CurrentDependency In Dependencies
        ' Check if local or global lambdas present or not
        If Not (IsExistInCollection(Lambdas, CStr(CurrentDependency)) _
                Or IsExistInCollection(Lambdas, QualifiedSheetName & CurrentDependency)) Then
            ValidDependencies.Add CurrentDependency
        End If
    Next CurrentDependency
    
    If ValidDependencies.Count = 0 Then
        ReDim Result(1 To 1, 1 To 1) As String
        Result(1, 1) = vbNullString
        GetDirectPrecedents = Result
    Else
        GetDirectPrecedents = CollectionToArray(ValidDependencies)
    End If
    
    Set ValidDependencies = Nothing
    
End Function

Private Function GetDirectPrecedentsFromExpr(ByVal Formula As String _
                                             , ByVal FormulaInSheet As Worksheet) As Collection
    
    If Formula = vbNullString Then
        Set GetDirectPrecedentsFromExpr = New Collection
        Exit Function
    End If
    
    #If DEVELOPMENT_MODE Then
        Dim ParseResult As OARobot.FormulaParseResult
        Dim CurrentExpr As OARobot.Expr
    #Else
        Dim ParseResult As Object
        Dim CurrentExpr As Object
    #End If
    
    Dim FormulaInBook As Workbook
    Set FormulaInBook = FormulaInSheet.Parent
    
    Set ParseResult = ParseFormula(Formula, FormulaInBook)
    
    If Not ParseResult.ParseSuccess Then Err.Raise 13, "DirectPrecedents", "Formula parsing failed."
    
    Dim Precedents As Collection
    Set Precedents = New Collection
    
    Dim Counter As Long
    For Counter = 0 To ParseResult.Expr.DirectPrecedents.Count - 1
        Set CurrentExpr = ParseResult.Expr.DirectPrecedents.Item(Counter)
        Precedents.Add CurrentExpr.Formula
    Next Counter
    
    Set GetDirectPrecedentsFromExpr = Precedents
           
End Function

' Check if the outer function is LAMBDA or not and it is the entire function.
Public Function IsLambdaFunction(ByVal Formula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        IsLambdaFunction = False
        Exit Function
    End If
    
    If ParsedFormulaResult.Expr.IsLambda Then
        IsLambdaFunction = True
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        IsLambdaFunction = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLambda
    Else
        IsLambdaFunction = False
    End If
    
End Function

Private Sub Test()
    
    Dim TestFormula As String
    TestFormula = "=AND(Category=""Financial"",Price<>0)"
    
    Dim ExpectedFormula As String
    
    Dim ActualFormula As String
    ActualFormula = ReplaceTokenWithNewToken(TestFormula, "Price", "y")
    Debug.Print ActualFormula
    
End Sub

Public Function ReplaceTokenWithNewToken(ByVal OnFormula As String _
                                         , ByVal OldToken As String _
                                          , ByVal NewToken As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ExprReplacer As OARobot.ExpressionReplacer
        Dim InputExpr As OARobot.Expr
    #Else
        Dim ExprReplacer As Object
        Dim InputExpr As Object
    #End If
    
    Set ExprReplacer = GetExpressionReplacer()
    
    With ExprReplacer
        .FindWhat = OldToken
        .ReplaceWith = NewToken
    End With

    Set InputExpr = GetExpr(OnFormula)
    Set InputExpr = InputExpr.Rewrite(ExprReplacer)
    ReplaceTokenWithNewToken = InputExpr.Formula(True)
    
End Function

Public Function IsFormulaParsedSucessfully(ByVal FormulaText As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(FormulaText)
    
    IsFormulaParsedSucessfully = ParsedFormulaResult.ParseSuccess
   
End Function

Public Function FormatFormula(ByVal FormulaText As String, Optional ByVal CompactConfig As Boolean = False) As String
    
    Dim Formatter As Object
    Set Formatter = CreateObject("OARobot.FormulaFormatter")
    
    With Formatter
        If CompactConfig Then .CompactConfig
        FormatFormula = .Format(FormulaText)
    End With
    Set Formatter = Nothing
    
End Function

' Check if outer function is LET or not and it is the entire function.
Public Function IsLetFunction(ByVal Formula As String) As Boolean
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedFormulaResult As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        IsLetFunction = False
        Exit Function
    End If
    
    If ParsedFormulaResult.Expr.IsLet Then
        IsLetFunction = True
    ElseIf ParsedFormulaResult.Expr.IsFunction Then
        IsLetFunction = ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLet
    Else
        IsLetFunction = False
    End If
    
    Set ParsedFormulaResult = Nothing
    
End Function

Public Function GetDependencyFunctionResult(ByVal Formula As String _
                                            , ByVal DependencyFunctionName As DependencyFunctions _
                                             , Optional ByVal RemoveHeaderRow As Boolean = True) As Variant
    
    #If DEVELOPMENT_MODE Then
        
        Dim ParseResult As OARobot.FormulaParseResult
        Dim Processor As OARobot.ExprProcessing
        Set Processor = New OARobot.ExprProcessing
    #Else
        
        Dim ParseResult As Object
        Dim Processor As Object
        Set Processor = CreateObject("OARobot.ExprProcessing")
    #End If
    
    Set ParseResult = ParseFormula(Formula)
    
    Dim Result As Variant
    Select Case DependencyFunctionName
        
        Case DependencyFunctions.FIRST_ARGUMENT_OF_OUTER_FUNCTION
            Dim AfterRemoved As Object
            Set AfterRemoved = Processor.FirstArgumentOfOuterFunction(ParseResult.Expr)
            ' It will be nothing in case of formula like: =J106 or =SUM(F5#)+1
            If AfterRemoved Is Nothing Then
                Result = Formula
            Else
                Result = EQUAL_SIGN & AfterRemoved.Formula
            End If
        
        Case DependencyFunctions.LET_PARTS
            Result = Processor.LetParts(ParseResult.Expr)

        Case DependencyFunctions.LAMBDA_PARTS
            Result = Processor.LambdaParts(ParseResult.Expr)
            
        Case Else
            Err.Raise 13, "Wrong Input Argument"

    End Select
    
    If RemoveHeaderRow Then Result = RemoveTopRowHeader(Result)
    GetDependencyFunctionResult = Result

End Function

' If a let function return a lambda as an output then we can invoke that. This will return that invocation part
' Example formula =LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1))))(5,8)
' Output will be (5,8)
'?GetLetFormulaInvocation(ActiveCell.Formula2)
Public Function GetLetFormulaInvocation(ByVal Formula As String) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim Args As OARobot.ExprCollection
    #Else
        Dim ParsedFormulaResult As Object
        Dim Args As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(Formula)
    
    Dim Invocation As String
    
    If Not ParsedFormulaResult.ParseSuccess Then
        ' If parse failed
        Invocation = vbNullString
    ElseIf ParsedFormulaResult.Expr.IsLet Then
        ' Parsing successful but Let statement
        Invocation = vbNullString
    ElseIf Not ParsedFormulaResult.Expr.IsFunction Then
        ' If not a function then
        Invocation = vbNullString
    ElseIf ParsedFormulaResult.Expr.AsFunction.FunctionName.IsLet _
           And ParsedFormulaResult.Expr.AsFunction.Args.Count > 0 Then
        
        Invocation = FIRST_PARENTHESIS_OPEN
        Set Args = ParsedFormulaResult.Expr.AsFunction.Args
        Dim Counter As Long
        For Counter = 0 To Args.Count - 1
            Invocation = Invocation & Args.Item(Counter).Formula & LIST_SEPARATOR
        Next Counter
        Invocation = Text.RemoveFromEndIfPresent(Invocation, LIST_SEPARATOR) & FIRST_PARENTHESIS_CLOSE
    
    End If
    
    GetLetFormulaInvocation = Invocation
    
    Set ParsedFormulaResult = Nothing
    
End Function

' This will extract upto the function definition. Meaning if the lambda is =LAMBDA(a,b,a*2)(10,5) >> =LAMBDA(a,b,
' It handled optional arguments, no param lambda as well.
Public Function GetUptoLambdaParamDefPart(ByVal LambdaFormula As String) As String
    
    Dim DefPart As String
    DefPart = GetLambdaDefPart(LambdaFormula)
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim Parameters As OARobot.TokenCollection
    #Else
        Dim ParsedFormulaResult As Object
        Dim Parameters As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(DefPart)
    If ParsedFormulaResult.Expr.IsLambda Then
        
        Set Parameters = ParsedFormulaResult.Expr.AsLambda.Parameters
        Dim ParamDefPart As String
        ParamDefPart = EQUAL_SIGN & LAMBDA_FN_NAME & FIRST_PARENTHESIS_OPEN
        Dim CurrentParam As Object
        
        Dim Counter As Long
        For Counter = 0 To Parameters.Count - 1
            Set CurrentParam = Parameters.Item(Counter)
            ParamDefPart = ParamDefPart & CurrentParam.String & LIST_SEPARATOR
        Next Counter
        
        GetUptoLambdaParamDefPart = ParamDefPart
        
    Else
        Err.Raise 13, "This function was expecting a lambda function."
    End If
    
End Function

Public Function GetLambdaDefPart(ByVal LambdaFormula As String) As String
    
    Dim SplittedPart As Variant
    SplittedPart = SplitLambdaDef(LambdaFormula)
    GetLambdaDefPart = SplittedPart(LBound(SplittedPart))
    
End Function

Private Function SplitLambdaDef(ByVal LambdaFormula As String) As String()
    
    Dim SplittedFormula(0 To 1) As String
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim InputExpr As OARobot.Expr
        Dim InputFunctions As OARobot.ExprFunction
    #Else
        Dim ParsedFormulaResult As Object
        Dim InputExpr As Object
        Dim InputFunctions As Object
    #End If
    
    Set ParsedFormulaResult = ParseFormula(LambdaFormula)
    
    If Not ParsedFormulaResult.ParseSuccess Then
        SplitLambdaDef = SplittedFormula
        Logger.Log DEBUG_LOG, "Failed to parse: " & ParsedFormulaResult.Formula
        Exit Function
    End If
    
    Set InputExpr = ParsedFormulaResult.Expr
    
    If InputExpr.IsFunction Then
        ' If lambda with invocation and only if the formula doesn't have any operands.
        ' For example =Lambda(a,a*2)+1 so here + is the operator and it will not be IsFunction
        
        Set InputFunctions = InputExpr.AsFunction
        If Not InputFunctions.FunctionName.IsLambda Then
            SplitLambdaDef = SplittedFormula
            Exit Function
        End If
                
        SplittedFormula(0) = InputFunctions.FunctionName.Formula(True)
        SplittedFormula(1) = GenerateInvocationPart(InputFunctions)
            
    ElseIf InputExpr.IsLambda Then
        ' If only lambda is present and no invocation.
        SplittedFormula(0) = InputExpr.Formula(True)
    End If
    
    SplitLambdaDef = SplittedFormula
    
End Function

Public Function GetLambdaInvocationPart(ByVal LambdaFormula As String) As String
    
    Dim SplittedPart As Variant
    SplittedPart = SplitLambdaDef(LambdaFormula)
    GetLambdaInvocationPart = SplittedPart(LBound(SplittedPart) + 1)
    
End Function

Public Function GetAllParamAndStepName(ByVal FormulaText As String) As Collection
    
    'Dim Processor As Object
    
    #If DEVELOPMENT_MODE Then
        Dim ParsedFormulaResult As OARobot.FormulaParseResult
        Dim Processor As OARobot.ExprProcessing
        Set Processor = New OARobot.ExprProcessing
    #Else
        Dim ParsedFormulaResult As Object
        Dim Processor As Object
        Set Processor = CreateObject("OARobot.ExprProcessing")
    #End If
    
    Set ParsedFormulaResult = ParseFormula(FormulaText)
    
    Dim Names As Collection
    Set Names = New Collection
    
    Dim Counter As Long
    Counter = 1
    
    '@TODO: Need to update this.
    Dim Temp As Variant
    Temp = Processor.LambdaParts(ParsedFormulaResult.Expr, Counter)
    Do While IsArrayAllocated(Temp)
        
        Dim FirstColumnIndex  As Long
        FirstColumnIndex = LBound(Temp, 2)
        Dim RowIndex As Long
        For RowIndex = LBound(Temp, 1) + 1 To UBound(Temp, 1) - 1
            
            Dim StepName As String
            StepName = Temp(RowIndex, FirstColumnIndex)
            StepName = Text.RemoveFromStartIfPresent(StepName, "[")
            StepName = Text.RemoveFromEndIfPresent(StepName, "]")
            If Not IsExistInCollection(Names, StepName) Then
                Names.Add StepName, StepName
            End If
            
        Next RowIndex
        
        Counter = Counter + 1
        Temp = Processor.LambdaParts(ParsedFormulaResult.Expr, Counter)
    Loop
    
    '@TODO: Need to update this.
    Counter = 1
    Temp = Processor.LetParts(ParsedFormulaResult.Expr, Counter)
    
    Do While IsArrayAllocated(Temp)
        
        FirstColumnIndex = LBound(Temp, 2)
        For RowIndex = LBound(Temp, 1) + 1 To UBound(Temp, 1) - 1
            
            StepName = Temp(RowIndex, FirstColumnIndex)
            If Not IsExistInCollection(Names, StepName) Then
                Names.Add StepName, StepName
            End If
            
        Next RowIndex
        
        Counter = Counter + 1
        Temp = Processor.LetParts(ParsedFormulaResult.Expr, Counter)
    Loop
    
    Set GetAllParamAndStepName = Names
    
End Function

#If DEVELOPMENT_MODE Then
Private Function GenerateInvocationPart(ByVal InputFunctions As OARobot.ExprFunction) As String
#Else
Private Function GenerateInvocationPart(ByVal InputFunctions As Object) As String
#End If
    
    #If DEVELOPMENT_MODE Then
        Dim Args As OARobot.ExprCollection
        Dim ArgSeps As OARobot.TokenCollection
    #Else
        Dim Args As Object
        Dim ArgSeps As Object
    #End If
    
    Dim InvocationPart As String
    InvocationPart = InputFunctions.LeftParen.String
            
    Set Args = InputFunctions.Args
    Set ArgSeps = InputFunctions.ArgSeparators
    Dim Counter As Long
                            
    For Counter = 0 To Args.Count - 1
        InvocationPart = InvocationPart & Args.Item(Counter).Formula(False)
        If Counter <= ArgSeps.Count - 1 Then
            InvocationPart = InvocationPart & ArgSeps.Item(Counter).String
        End If
    Next Counter
    
    InvocationPart = InvocationPart & InputFunctions.RightParen.String
    GenerateInvocationPart = InvocationPart
    
End Function

#If DEVELOPMENT_MODE Then
Public Function GetScope(Optional ByVal ForBook As Workbook) As OARobot.FormulaScopeInfo
#Else
Public Function GetScope(Optional ByVal ForBook As Workbook) As Object
#End If
    
    If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
    #If DEVELOPMENT_MODE Then
        Dim ScopeFactory As New OARobot.FormulaScopeFactory
        Set GetScope = ScopeFactory.CreateWorkbook(ForBook.Name)
    #Else
        Set GetScope = CreateObject("OARobot.FormulaScopeFactory").CreateWorkbook(ForBook.Name)
    #End If
    
End Function

#If DEVELOPMENT_MODE Then
Public Function GetNames(Optional ByVal ForBook As Workbook) As OARobot.XLDefinedNames
#Else
Public Function GetNames(Optional ByVal ForBook As Workbook) As Object
#End If
    
    If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
    #If DEVELOPMENT_MODE Then
        Dim NamesFactory As New OARobot.DefinedNamesFactory
        Set GetNames = NamesFactory.Create(ForBook)
    #Else
        Set GetNames = CreateObject("OARobot.DefinedNamesFactory").Create(ForBook)
    #End If
    
End Function

#If DEVELOPMENT_MODE Then
Public Function ParseFormula(ByVal Formula As String _
                             , Optional ByVal ForBook As Workbook _
                              , Optional ByVal IsR1C1 As Boolean = False) As OARobot.FormulaParseResult
#Else
Public Function ParseFormula(ByVal Formula As String _
                             , Optional ByVal ForBook As Workbook _
                              , Optional ByVal IsR1C1 As Boolean = False) As Object
#End If
    
Static ScopeBookName As String
If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
#If DEVELOPMENT_MODE Then
    Static Scope As OARobot.FormulaScopeInfo
    Static DefinedNames As OARobot.XLDefinedNames
    Dim Parser As OARobot.FormulaParser
    Dim LocaleFactory As FormulaLocaleInfoFactory
    Set LocaleFactory = New FormulaLocaleInfoFactory
#Else
    Static Scope As Object
    Static DefinedNames As Object
    Dim Parser As Object
    Dim LocaleFactory As Object
    Set LocaleFactory = CreateObject("FormulaLocaleInfoFactory")
#End If
    
If ScopeBookName <> ForBook.Name Or ScopeBookName = vbNullString Then
    Set Scope = GetScope(ForBook)
    Set DefinedNames = GetNames(ForBook)
    ScopeBookName = ForBook.Name
End If
    
Set Parser = GetFormulaParser()
    
If ForBook Is Nothing Then Set ForBook = ActiveWorkbook
    
Set ParseFormula = Parser.Parse(Formula, IsR1C1, LocaleFactory.EN_US, Scope, DefinedNames)
    
End Function

#If DEVELOPMENT_MODE Then
Public Function GetFormulaParser() As OARobot.FormulaParser
#Else
Public Function GetFormulaParser() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetFormulaParser = New OARobot.FormulaParser
    #Else
        Set GetFormulaParser = CreateObject("OARobot.FormulaParser")
    #End If

End Function

' Return Expr object by parsing the formula. If parsing fail then it will return nothing
#If DEVELOPMENT_MODE Then
Public Function GetExpr(ByVal Formula As String) As OARobot.Expr
#Else
Public Function GetExpr(ByVal Formula As String) As Object
#End If
        
    #If DEVELOPMENT_MODE Then
        Dim ParsedResult As OARobot.FormulaParseResult
    #Else
        Dim ParsedResult As Object
    #End If
    
    Set ParsedResult = ParseFormula(Text.PadIfNotPresent(Formula, EQUAL_SIGN, FROM_START))
    If ParsedResult.ParseSuccess Then
        Set GetExpr = ParsedResult.Expr
    Else
        Set GetExpr = Nothing
    End If
    
End Function

#If DEVELOPMENT_MODE Then
Public Function GetExpressionReplacer() As OARobot.ExpressionReplacer
#Else
Public Function GetExpressionReplacer() As Object
#End If

    #If DEVELOPMENT_MODE Then
        Set GetExpressionReplacer = New OARobot.ExpressionReplacer
    #Else
        Set GetExpressionReplacer = CreateObject("OARobot.ExpressionReplacer")
    #End If

End Function

Public Function GetUsedFunctions(ByVal Formula As String, Optional ByVal IsR1C1 As Boolean = False) As Variant
    
    If Formula = vbNullString Then
        GetUsedFunctions = vbEmpty
        Exit Function
    End If
    
    #If DEVELOPMENT_MODE Then
        Dim ParseResult As OARobot.FormulaParseResult
    #Else
        Dim ParseResult As Object
    #End If
    
    Set ParseResult = ParseFormula(Formula, , IsR1C1)
    
    If Not ParseResult.ParseSuccess Then Err.Raise 13, "UsedFunctions", "Formula parsing failed."
    
    GetUsedFunctions = ParseResult.Expr.UsedFunctions
           
End Function
