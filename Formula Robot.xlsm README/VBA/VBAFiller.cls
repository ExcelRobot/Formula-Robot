VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBAFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Filler.Impl")

Option Explicit

Implements IFiller

Private Sub IFiller_Fill(ByVal FillRange As Range, ByVal TypeOfFill As FillType, Optional ByVal ArrayFormula As String)
    
    If IsNotNothing(FillRange) Then
        If TypeOfFill = Fill_DOWN Then
            FillRange.Resize(1).AutoFill Destination:=FillRange, Type:=xlFillDefault
        ElseIf TypeOfFill = FILL_TO_RIGHT Then
            FillRange.Resize(, 1).AutoFill Destination:=FillRange, Type:=xlFillDefault
        End If
    End If
    
End Sub
