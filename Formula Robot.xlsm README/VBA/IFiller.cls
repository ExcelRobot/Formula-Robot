VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder("Filler.Interface")


Option Explicit

Public Enum FillType
    Fill_DOWN = 1
    FILL_TO_RIGHT = 2
End Enum

Public Sub Fill(ByVal FillRange As Range, ByVal TypeOfFill As FillType, Optional ByVal ArrayFormula As String)

End Sub
