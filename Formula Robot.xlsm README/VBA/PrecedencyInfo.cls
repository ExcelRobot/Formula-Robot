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
    FillLagWithCell As Range
    IsOverlapWithFormulaCellIfResized As Boolean
End Type
Private this As TPrecedencyInfo

Public Property Get IsOverlapWithFormulaCellIfResized() As Boolean
    IsOverlapWithFormulaCellIfResized = this.IsOverlapWithFormulaCellIfResized
End Property

Public Property Let IsOverlapWithFormulaCellIfResized(ByVal RHS As Boolean)
    this.IsOverlapWithFormulaCellIfResized = RHS
End Property

Public Property Get FillLagWithCell() As Range
    Set FillLagWithCell = this.FillLagWithCell
End Property

Public Property Set FillLagWithCell(ByVal RHS As Range)
    Set this.FillLagWithCell = RHS
End Property

Public Property Get AbsRangeRef() As String
    AbsRangeRef = this.AbsRangeRef
End Property

Public Property Let AbsRangeRef(ByVal RHS As String)
    this.AbsRangeRef = RHS
End Property

Public Property Get AbsChoosePartFormula() As String
    AbsChoosePartFormula = this.AbsChoosePartFormula
End Property

Public Property Let AbsChoosePartFormula(ByVal RHS As String)
    this.AbsChoosePartFormula = RHS
End Property

Public Property Get NameInFormulaRange() As Range
    Set NameInFormulaRange = this.NameInFormulaRange
End Property

Public Property Set NameInFormulaRange(ByVal RHS As Range)
    Set this.NameInFormulaRange = RHS
End Property

Public Property Get HasSpill() As Boolean
    HasSpill = this.HasSpill
End Property

Public Property Let HasSpill(ByVal RHS As Boolean)
    this.HasSpill = RHS
End Property

Public Property Get NameInFormula() As String
    NameInFormula = this.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    this.NameInFormula = RHS
End Property

Public Property Get RangeRef() As String
    RangeRef = this.ParentCellAddress
End Property

Public Property Let RangeRef(ByVal RHS As String)
    this.ParentCellAddress = RHS
End Property

Public Property Get ChoosePartFormula() As String
    ChoosePartFormula = this.ChoosePartFormula
End Property

Public Property Let ChoosePartFormula(ByVal RHS As String)
    this.ChoosePartFormula = RHS
End Property

Public Property Get ColOrRowIndex() As Long
    ColOrRowIndex = this.ColOrRowIndex
End Property

Public Property Let ColOrRowIndex(ByVal RHS As Long)
    this.ColOrRowIndex = RHS
End Property

Public Property Get RowCount() As Long
    RowCount = this.RowCount
End Property

Public Property Let RowCount(ByVal RHS As Long)
    this.RowCount = RHS
End Property

Public Property Get ColCount() As Long
    ColCount = this.ColCount
End Property

Public Property Let ColCount(ByVal RHS As Long)
    this.ColCount = RHS
End Property

Public Property Get TopLeftCellRowNo() As Long
    TopLeftCellRowNo = this.TopLeftCellRowNo
End Property

Public Property Let TopLeftCellRowNo(ByVal RHS As Long)
    this.TopLeftCellRowNo = RHS
End Property

Public Property Get TopLeftCellColNo() As Long
    TopLeftCellColNo = this.TopLeftCellColNo
End Property

Public Property Let TopLeftCellColNo(ByVal RHS As Long)
    this.TopLeftCellColNo = RHS
End Property

