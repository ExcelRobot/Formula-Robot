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

Private This As TFillBean

Public Property Get UndoOperator() As UndoHandler
    Set UndoOperator = This.UndoOperator
End Property

Public Property Set UndoOperator(ByVal RHS As UndoHandler)
    Set This.UndoOperator = RHS
End Property

Public Property Get OnRange() As Range
    Set OnRange = This.OnRange
End Property

Public Property Set OnRange(ByVal RHS As Range)
    Set This.OnRange = RHS
End Property

Public Property Get TypeOfFill() As FillType
    TypeOfFill = This.TypeOfFill
End Property

Public Property Let TypeOfFill(ByVal RHS As FillType)
    This.TypeOfFill = RHS
End Property

Public Property Get ProbableFillDimension() As Long
    ProbableFillDimension = This.ProbableFillDimension
End Property

Public Property Let ProbableFillDimension(ByVal RHS As Long)
    This.ProbableFillDimension = RHS
End Property

Public Property Get DAFormula() As String
    DAFormula = This.DAFormula
End Property

Public Property Let DAFormula(ByVal RHS As String)
    This.DAFormula = RHS
End Property

Public Property Get Filler() As IFiller
    Set Filler = This.Filler
End Property

Public Property Set Filler(ByVal RHS As IFiller)
    Set This.Filler = RHS
End Property

Public Sub DoFill()
    
    If This.DAFormula = vbNullString Then
        This.Filler.Fill This.OnRange, This.TypeOfFill
    Else
        This.Filler.Fill This.OnRange, This.TypeOfFill, This.DAFormula
    End If
    
End Sub
