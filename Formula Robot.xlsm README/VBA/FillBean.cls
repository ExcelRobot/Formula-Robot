VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FillBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TFillBean
    OnRange As Range
    TypeOfFill As FillType
    ProbableFillDimension As Long
    DAFormula As String
    Filler As IFiller
    UndoOperator As UndoHandler
End Type

Private this As TFillBean

Public Property Get UndoOperator() As UndoHandler
    Set UndoOperator = this.UndoOperator
End Property

Public Property Set UndoOperator(ByVal RHS As UndoHandler)
    Set this.UndoOperator = RHS
End Property

Public Property Get OnRange() As Range
    Set OnRange = this.OnRange
End Property

Public Property Set OnRange(ByVal RHS As Range)
    Set this.OnRange = RHS
End Property

Public Property Get TypeOfFill() As FillType
    TypeOfFill = this.TypeOfFill
End Property

Public Property Let TypeOfFill(ByVal RHS As FillType)
    this.TypeOfFill = RHS
End Property

Public Property Get ProbableFillDimension() As Long
    ProbableFillDimension = this.ProbableFillDimension
End Property

Public Property Let ProbableFillDimension(ByVal RHS As Long)
    this.ProbableFillDimension = RHS
End Property

Public Property Get DAFormula() As String
    DAFormula = this.DAFormula
End Property

Public Property Let DAFormula(ByVal RHS As String)
    this.DAFormula = RHS
End Property

Public Property Get Filler() As IFiller
    Set Filler = this.Filler
End Property

Public Property Set Filler(ByVal RHS As IFiller)
    Set this.Filler = RHS
End Property

Public Sub DoFill()
    
    If this.DAFormula = vbNullString Then
        this.Filler.Fill this.OnRange, this.TypeOfFill
    Else
        this.Filler.Fill this.OnRange, this.TypeOfFill, this.DAFormula
    End If
    
End Sub
