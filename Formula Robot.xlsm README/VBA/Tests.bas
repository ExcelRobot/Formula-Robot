Attribute VB_Name = "Tests"
Option Explicit
#Const DEVELOPMENT_MODE = False

Public Sub RunFillDownTests()
    
    With ActiveSheet
        Const TOTAL_TILE_TEST_COUNT As Long = 28
        Dim Counter As Long
        For Counter = 1 To TOTAL_TILE_TEST_COUNT
            Dim FormulaCell As Range
            Set FormulaCell = Range("FillDownBefore" & Counter)
            Dim AfterCell As Range
            Set AfterCell = Range("FillDownAfter" & Counter)
            
            FormulaCell.Copy AfterCell
            If Counter = 18 Then
                Set FormulaCell = FormulaCell.Rows(1)
                FormulaCell.Copy AfterCell
            End If
'            Debug.Assert Counter <> 20
            PasteFillDown FormulaCell, AfterCell
            ' Native Fill Down is bit slower.
            If Counter = 13 Then
                Application.Wait (Now + TimeValue("00:00:05"))
            End If
            
        Next Counter
    End With
    ShowMessageForFailedTests
    
End Sub

Private Sub ShowMessageForFailedTests()
    
    Dim AllRefs As Variant
    AllRefs = Text.SplitText(Text.RemoveFromEndIfPresent(Text.RemoveFromStartIfPresent(ActiveSheet.Range("B2").Formula2, "=AND("), ")"), ",")
    
    Dim FailedTestRefs As String
    Dim CurrentRef As Variant
    For Each CurrentRef In AllRefs
        If Not Range(CStr(CurrentRef)) Then
            FailedTestRefs = FailedTestRefs & "," & CurrentRef
        End If
    Next CurrentRef
    
    If FailedTestRefs <> vbNullString Then
        MsgBox "Tests failed for:" & Text.RemoveFromStartIfPresent(FailedTestRefs, ",")
    End If
    
End Sub

Private Sub ClearPreviousResult(ByVal AfterCell As Range, ByVal IsForFillDown As Boolean)
    
    Dim Temp As Range
    If AfterCell.HasSpill Then
        AfterCell.Formula = vbNullString
    Else
        Set Temp = AfterCell
        Do While Not IsAllBlankAndNoFormulas(Temp)
            Temp.ClearContents
            If IsForFillDown Then
                Set Temp = Temp.Offset(1, 0)
            Else
                Set Temp = Temp.Offset(0, 1)
            End If
        Loop
    End If
    
End Sub

Public Sub RunFillToRightTests()
    
    With ActiveSheet
        Const TOTAL_TILE_TEST_COUNT As Long = 27
        Dim Counter As Long
        For Counter = 1 To TOTAL_TILE_TEST_COUNT
            Dim FormulaCell As Range
            Set FormulaCell = Range("FillToRightBefore" & Counter)
            Dim AfterCell As Range
            Set AfterCell = Range("FillToRightAfter" & Counter)
            
            FormulaCell.Copy AfterCell
            If Counter = 18 Then
                Set FormulaCell = FormulaCell.Columns(1)
                FormulaCell.Copy AfterCell
            End If
            
            PasteFillToRight FormulaCell, AfterCell
            
        Next Counter
    End With
    ShowMessageForFailedTests
    
End Sub

Private Sub PrintWaitFinish()
    Debug.Print "Waiting is done."
End Sub

Public Sub RunMapTests()
    
    With ActiveSheet
        Const TOTAL_TILE_TEST_COUNT As Long = 5
        Dim Counter As Long
        For Counter = 1 To TOTAL_TILE_TEST_COUNT
            Dim FormulaCell As Range
            Set FormulaCell = Range("MapBefore" & Counter)
            Dim AfterCell As Range
            Set AfterCell = Range("MapAfter" & Counter)
            ClearPreviousResult AfterCell, True
            MapToArray FormulaCell, AfterCell
        Next Counter
    End With
    MsgBox "All test is successfully executed.", vbOKOnly, "Map To Array Tests"
    
End Sub

Public Sub RunFilterTests()
    
    With ActiveSheet
        Const TOTAL_TILE_TEST_COUNT As Long = 8
        Dim Counter As Long
        For Counter = 1 To TOTAL_TILE_TEST_COUNT
            Dim FormulaCell As Range
            Set FormulaCell = Range("FilterBefore" & Counter)
            Dim AfterCell As Range
            Set AfterCell = Range("FilterAfter" & Counter)
            ClearPreviousResult AfterCell, True
            ApplyFilterToArray FormulaCell, AfterCell
        Next Counter
    End With
    MsgBox "All test is successfully executed.", vbOKOnly, "Apply Filter To Array Tests"
    
End Sub

Public Sub TestExtractAllNamedRangeRef()
    
    Dim DepExtractor As RangeDependencyInChart
    Set DepExtractor = New RangeDependencyInChart
    Const TEST_NAME As String = "'Local Scoped Named Range'!HPercentAverage"
    DepExtractor.ExtractAllNamedRangeRef ActiveWorkbook.Names(TEST_NAME), ActiveWorkbook
    Debug.Print DepExtractor.AllDependency.Count
    
End Sub

Private Sub Test()
    
    
    Const Formula As String = "=AND(Category=""Financial"",Price<>0)"
    Dim Dependencies As Variant
    Dim FormulaProcessor As Object
    Set FormulaProcessor = CreateObject("OARobot.FormulaProcessing")
        
    Dependencies = FormulaProcessor.DirectPrecedents(Formula)
    
    Dim IsTestPass As Boolean
    IsTestPass = ((UBound(Dependencies, 1) - LBound(Dependencies, 1) + 1) = 2)
    
    If IsTestPass Then
        MsgBox "Test passed successfully."
    Else
        MsgBox "Test failed."
    End If
    
End Sub

