VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
    ' Sheets protection
    If Intersect(target, Me.Range("PROTECT_SHEETS")) Is Nothing Then Exit Sub

    On Error GoTo CleanFail
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If Me.Range("PROTECT_SHEETS").value = False Then
        Functions.UnprotectAllSheets
    Else
        Functions.ProtectAllSheets
    End If
    

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    ' Fail silently (or MsgBox for debugging)
    Resume CleanExit
End Sub
