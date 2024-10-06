VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DAFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Filler.Impl")

Option Explicit

Implements IFiller

Private Sub IFiller_Fill(ByVal FillRange As Range, ByVal TypeOfFill As FillType, Optional ByVal ArrayFormula As String)
    
    If IsNotNothing(FillRange) Then
        FillRange.Cells(1).Formula2 = ArrayFormula
        
        Dim PrevSelection As Variant
        Set PrevSelection = Selection
        Dim CutCopyMode As XlCutCopyMode
        CutCopyMode = Application.CutCopyMode
        FillRange.Cells(1).Copy
        FillRange.Cells(1).SpillingToRange.PasteSpecial xlPasteFormats
        Application.CutCopyMode = CutCopyMode
        If Not PrevSelection Is Nothing Then PrevSelection.Select
        
    End If
    
End Sub

