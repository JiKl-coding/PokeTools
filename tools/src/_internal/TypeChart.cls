VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo CleanFail

    Dim Type1 As Range
    Dim Type2 As Range

    Set Type1 = Me.Range("PKMN_TYPE_1")
    Set Type2 = Me.Range("PKMN_TYPE_2")

    ' Only react when one of the type cells changes
    If Intersect(target, Union(Type1, Type2)) Is Nothing Then Exit Sub

    ' Prevent infinite event loop
    Application.EnableEvents = False

    ' If both types are the same and not empty, clear TYPE_2
    If Len(Type1.Value2) > 0 And Type1.Value2 = Type2.Value2 Then
        Type2.Value2 = vbNullString
    End If

CleanExit:
    ' Always re-enable events
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    ' Ensure events are restored even if an error occurs
    Resume CleanExit
End Sub

