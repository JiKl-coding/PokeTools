VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Movedex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    Dim watchRange As Range
    On Error GoTo CleanFail

    Set watchRange = Union(Me.Range("GAMEVERSION"), Me.Range("MVLIST"))
    
    If (Me.Range("GAMEVERSION") = vbNullString Or Me.Range("GAMEVERSION") = "0") Then
        Me.Range("GAMEVERSION") = "All"
    End If
    
    If target Is Nothing Then Exit Sub
    If watchRange Is Nothing Then Exit Sub
    If Intersect(target, watchRange) Is Nothing Then Exit Sub
    If target.CountLarge > 1 Then Exit Sub

    Application.EnableEvents = False
    Functions.UnprotectForMacros Movedex

    MoveDexLogic.HandleWorksheetChange Me, target

CleanExit:
    Functions.ProtectForMacros Movedex
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

