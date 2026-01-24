VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializerCellRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type TInitializerCellRef
    ParentCellRef As String
    HasSpill As Boolean
    Index As Long
    IsWholeSpillUsed As Boolean
    NameInFormula As String
End Type
Private this As TInitializerCellRef

Public Property Get NameInFormula() As String
    NameInFormula = this.NameInFormula
End Property

Public Property Let NameInFormula(ByVal RHS As String)
    this.NameInFormula = RHS
End Property

Public Property Get ParentCellRef() As String
    ParentCellRef = this.ParentCellRef
End Property

Public Property Let ParentCellRef(ByVal RHS As String)
    this.ParentCellRef = RHS
End Property

Public Property Get HasSpill() As Boolean
    HasSpill = this.HasSpill
End Property

Public Property Let HasSpill(ByVal RHS As Boolean)
    this.HasSpill = RHS
End Property

Public Property Get Index() As Long
    Index = this.Index
End Property

Public Property Let Index(ByVal RHS As Long)
    this.Index = RHS
End Property

Public Property Get IsWholeInitializerRangeUsed() As Boolean
    IsWholeInitializerRangeUsed = this.IsWholeSpillUsed
End Property

Public Property Let IsWholeInitializerRangeUsed(ByVal RHS As Boolean)
    this.IsWholeSpillUsed = RHS
End Property
