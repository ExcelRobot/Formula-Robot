Attribute VB_Name = "RemoveOuterFunctionTests"
'@TestModule
'@Folder("Tests.RemoveOuterFunction")
Option Explicit
Option Private Module
Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("LAMBDA No LET Calc")
Private Sub LAMBDANoLETCalc()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,TRANSPOSE(SEQUENCE(x,y)))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,SEQUENCE(x,y))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA No LET Calc2")
Private Sub LAMBDANoLETCalc2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,SEQUENCE(x,y))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,x)(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA Just Return Param")
Private Sub LAMBDAJustReturnParam()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,x)(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,x)(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Last Step Calc")
Private Sub LETLastStepCalc()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,TRANSPOSE(SEQUENCE(x,y)))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,SEQUENCE(x,y))"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Last Step Calc2")
Private Sub LETLastStepCalc2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,SEQUENCE(x,y))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,x)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Does not Refer Last Step")
Private Sub LETResultDoesnotReferLastStep()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,x)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,y)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Refer Last Step")
Private Sub LETResultReferLastStep()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,y)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,x)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET One Single Step")
Private Sub LETOneSingleStep()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,x)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=5"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("Simple Formula")
Private Sub SimpleFormula()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=5"
    Dim ExpectedFormula As String
    ExpectedFormula = "=5"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Refer Last Step2")
Private Sub LETResultReferLastStep2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,z,SEQUENCE(x,y),result,TRANSPOSE(z),result)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,z,SEQUENCE(x,y),z)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Refer Last Step3")
Private Sub LETResultReferLastStep3()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,z,SEQUENCE(x,y),z)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,x)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Refer Last Step4")
Private Sub LETResultReferLastStep4()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,z,SEQUENCE(y,x),result,TRANSPOSE(z),result)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,z,SEQUENCE(y,x),z)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Refer Last Step5")
Private Sub LETResultReferLastStep5()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,z,SEQUENCE(y,x),z)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,y)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Is Calc But Not Formula2")
Private Sub LETResultIsCalcButNotFormula2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(x,5,y,8,z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(x,5,y,8,z,SEQUENCE(y,x),result,TRANSPOSE(z),result)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA Result Is Calc But Not Formula")
Private Sub LAMBDAResultIsCalcButNotFormula()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1)))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA Result Refer Last Step")
Private Sub LAMBDAResultReferLastStep()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,LET(z,SEQUENCE(y,x),z))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA Result Refer Last Step2")
Private Sub LAMBDAResultReferLastStep2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,LET(z,SEQUENCE(y,x),z))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,y)(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LAMBDA Just Return Param2")
Private Sub LAMBDAJustReturnParam2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LAMBDA(x,y,y)(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LAMBDA(x,y,y)(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET Result Is Calc But Not Formula")
Private Sub LETResultIsCalcButNotFormula()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1))))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result)))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Last Step Is Result2")
Private Sub LETAndLAMBDALastStepIsResult2()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result)))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),z)))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Last Step Is Result3")
Private Sub LETAndLAMBDALastStepIsResult3()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),z)))(5,8)"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,y))(5,8)"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Result Step Is Calc")
Private Sub LETAndLAMBDAResultStepIsCalc()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),2*(result+1)))(5,8))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result))(5,8))"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Last Step Is Result4")
Private Sub LETAndLAMBDALastStepIsResult4()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),result,TRANSPOSE(z),result))(5,8))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),z))(5,8))"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Last Step Is Result5")
Private Sub LETAndLAMBDALastStepIsResult5()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,LET(z,SEQUENCE(y,x),z))(5,8))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,y)(5,8))"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("LET And LAMBDA Nothing To Remove")
Private Sub LETAndLAMBDANothingToRemove()

    On Error GoTo TestFail

    'Arrange:
    Dim Formula As String
    Formula = "=LET(a,1,LAMBDA(x,y,y)(5,8))"
    Dim ExpectedFormula As String
    ExpectedFormula = "=LET(a,1,LAMBDA(x,y,y)(5,8))"
    ExpectedFormula = FormatFormula(ExpectedFormula)

    'Act:
    Dim ActualFormula As String
    ActualFormula = RemoveOuterFunctionFromFormula(Formula)
    ActualFormula = FormatFormula(ActualFormula)

    'Assert:
    Assert.AreEqual ExpectedFormula, ActualFormula

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

