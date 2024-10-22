Attribute VB_Name = "modDynamicFormula"
'@IgnoreModule UndeclaredVariable
'@Folder "DynamicFormula"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Select Spill Parent
' Description:            This will select spill parent cell. If no spill in ActiveCell then do nothing.
' Macro Expression:       modDynamicFormula.SelectSpillParent()
' Generated:              06/04/2023 03:05 PM
'----------------------------------------------------------------------------------------------------
Public Sub SelectSpillParent()

    ' Check if the currently active cell is part of a spill range
    If ActiveCell.HasSpill Then
        ' If it is, select the parent of the spill range
        ActiveCell.SpillParent.Select
        ' Check if the selected cell is visible in the current window
        If Intersect(ActiveWindow.VisibleRange, ActiveCell.SpillParent) Is Nothing Then
            ' If it's not visible, scroll to make the active cell visible
            Application.Goto ActiveCell, True
        End If
    End If

End Sub

