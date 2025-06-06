VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchFunctionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

Private Type TSearchFunctionData
    Location As String
    Functions As String
    SearchSection As String
    Formula As String
End Type

Private This As TSearchFunctionData

Public Property Get Location() As String
    Location = This.Location
End Property

Public Property Let Location(ByVal RHS As String)
    This.Location = RHS
End Property

Public Property Get Functions() As String
    Functions = This.Functions
End Property

Public Property Let Functions(ByVal RHS As String)
    This.Functions = RHS
End Property

Public Property Get SearchSection() As String
    SearchSection = This.SearchSection
End Property

Public Property Let SearchSection(ByVal RHS As String)
    This.SearchSection = RHS
End Property

Public Property Get Formula() As String
    Formula = Text.PadIfNotPresent(This.Formula, "'", FROM_START)
End Property

Public Property Let Formula(ByVal RHS As String)
    This.Formula = RHS
End Property

Public Property Get Self() As SearchFunctionData
    Set Self = Me
End Property

Public Function Create(ByVal Location As String _
                       , ByVal Functions As String _
                        , ByVal SearchSection As String _
                         , ByVal Formula As String) As SearchFunctionData

    With New SearchFunctionData
        .Location = Location
        .Functions = Functions
        .SearchSection = SearchSection
        .Formula = Formula
        Set Create = .Self
    End With

End Function


