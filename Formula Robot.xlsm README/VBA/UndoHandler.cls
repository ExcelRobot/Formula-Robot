VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UndoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Public Enum OptType
    VBA_VERSION = 1
    DYNAMIC_ARRAY_VERSION = 2
End Enum

Private Type TUndoHandler
    TypeOfOpt As OptType
    ClearRange As Range
    FirstCellOldFormula As String
    SecondCellFormulaR1C1 As String
End Type

Private this As TUndoHandler

Public Property Get SecondCellFormulaR1C1() As String
    SecondCellFormulaR1C1 = this.SecondCellFormulaR1C1
End Property

Public Property Let SecondCellFormulaR1C1(ByVal RHS As String)
    this.SecondCellFormulaR1C1 = RHS
End Property

Public Property Get TypeOfOpt() As OptType
    TypeOfOpt = this.TypeOfOpt
End Property

Public Property Let TypeOfOpt(ByVal RHS As OptType)
    this.TypeOfOpt = RHS
End Property

Public Property Get ClearRange() As Range
    Set ClearRange = this.ClearRange
End Property

Public Property Set ClearRange(ByVal RHS As Range)
    Set this.ClearRange = RHS
End Property

Public Property Get FirstCellOldFormula() As String
    FirstCellOldFormula = this.FirstCellOldFormula
End Property

Public Property Let FirstCellOldFormula(ByVal RHS As String)
    this.FirstCellOldFormula = RHS
End Property

Public Property Get Self() As UndoHandler
    Set Self = Me
End Property

Public Function Create(ByVal TypeOfOpt As OptType _
                       , ByVal ClearRange As Range _
                        , ByVal FirstCellOldFormula As String) As UndoHandler

    With New UndoHandler
        .TypeOfOpt = TypeOfOpt
        Set .ClearRange = ClearRange
        .FirstCellOldFormula = FirstCellOldFormula
        Set Create = .Self
    End With

End Function

Public Sub Undo()
    
    If this.TypeOfOpt = VBA_VERSION Then
        this.ClearRange.ClearContents
        If this.SecondCellFormulaR1C1 <> vbNullString Then
            ClearRange.Formula2R1C1 = ReplaceInvalidCharFromFormulaWithValid(this.SecondCellFormulaR1C1)
            ClearRange.Calculate
        End If
    ElseIf this.TypeOfOpt = DYNAMIC_ARRAY_VERSION Then
        ClearRange.Cells(1).Formula2 = ReplaceInvalidCharFromFormulaWithValid(this.FirstCellOldFormula)
    End If
    
End Sub

