VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' === ThisWorkbook code ===
Option Explicit


Private Sub Workbook_Open()
    Functions.CheckPaths
    Functions.EnsurePokedataOpen
    Application.EnableEvents = True
    Functions.ProtectForMacros Pokedex
    Functions.ProtectForMacros Lists
    Functions.ProtectForMacros TypeChart
End Sub

Private Sub Workbook_Activate()
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    
    If Not Functions.OWNS_POKEDATA Then
        Exit Sub
    End If

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.name) = "pokedata.xlsx" Then
            Application.DisplayAlerts = False
            wb.Saved = True
            wb.Close SaveChanges:=False
            Application.DisplayAlerts = True
            Exit For
        End If
    Next wb

    On Error GoTo 0
End Sub



