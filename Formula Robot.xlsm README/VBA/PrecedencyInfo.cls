VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrecedencyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TPrecedencyInfo
    NameInFormula As String
    ParentCellAddress As String
    ChoosePartFormula As String
    ColOrRowIndex As Long
    RowCount As Long
    ColCount As Long
    TopLeftCellRowNo As Long
    TopLeftCellColNo As Long
    HasSpill As Boolean
    NameInFormulaRange As Range
    AbsRangeRef As String
    AbsChoosePartFormula As String
End Type

Private This As TPrecedencyInfo

Public Property Get AbsRangeRef() As String
    AbsRangeRef = This.AbsRangeRef
End Property

Public Property Let AbsRangeRef(ByVal RHS As String)
    This.AbsRangeRef = RHS
End Property

Public Property Get AbsChoosePartFormula() As String
    AbsChoosePartFormula = This.AbsChoosePartFormula
End Property

Public Property Let AbsChoosePartFormula(ByVal RHS As String)
    This.AbsChoosePartFormula = RHS
End Property

Public Property Get NameInFormulaRange() As Range
    Set NameInFormulaRange = This.NameInFormulaRange
End Property

Public Property Set NameInFormulaRange(ByVal RHS As Range)
    Set This.NameInFormulaRange = RHS
End Property

Public Property Get HasSpill() As Boolean
    HasSpill = This.HasSpill
End Property

Public Property Let HasSpill(ByVal RHS As Boolean)
    This.HasSpill = RHS
End Property

Public Property Get NameInFormula() As String
    NameInFormula = This.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    This.NameInFormula = RHS
End Property

Public Property Get RangeRef() As String
    RangeRef = This.ParentCellAddress
End Property

Public Property Let RangeRef(ByVal RHS As String)
    This.ParentCellAddress = RHS
End Property

Public Property Get ChoosePartFormula() As String
    ChoosePartFormula = This.ChoosePartFormula
End Property

Public Property Let ChoosePartFormula(ByVal RHS As String)
    This.ChoosePartFormula = RHS
End Property

Public Property Get ColOrRowIndex() As Long
    ColOrRowIndex = This.ColOrRowIndex
End Property

Public Property Let ColOrRowIndex(ByVal RHS As Long)
    This.ColOrRowIndex = RHS
End Property

Public Property Get RowCount() As Long
    RowCount = This.RowCount
End Property

Public Property Let RowCount(ByVal RHS As Long)
    This.RowCount = RHS
End Property

Public Property Get ColCount() As Long
    ColCount = This.ColCount
End Property

Public Property Let ColCount(ByVal RHS As Long)
    This.ColCount = RHS
End Property

Public Property Get TopLeftCellRowNo() As Long
    TopLeftCellRowNo = This.TopLeftCellRowNo
End Property

Public Property Let TopLeftCellRowNo(ByVal RHS As Long)
    This.TopLeftCellRowNo = RHS
End Property

Public Property Get TopLeftCellColNo() As Long
    TopLeftCellColNo = This.TopLeftCellColNo
End Property

Public Property Let TopLeftCellColNo(ByVal RHS As Long)
    This.TopLeftCellColNo = RHS
End Property

