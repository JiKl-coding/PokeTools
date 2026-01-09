VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGridComboEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Combo As MSForms.ComboBox
Attribute Combo.VB_VarHelpID = -1

Private mParent As Object

Public Sub Init(ByVal parentForm As Object, ByVal cbo As MSForms.ComboBox)
    Set mParent = parentForm
    Set Combo = cbo
End Sub

' Trigger filter updates only on committed changes
Private Sub Combo_AfterUpdate()
  On Error Resume Next
  CallByName mParent, "FilterControlChanged", VbMethod, Combo.name
  If Err.Number <> 0 Then
    Err.Clear
    CallByName mParent, "FiltersChanged", VbMethod
  End If
  On Error GoTo 0
End Sub

Private Sub Combo_Click()
  On Error Resume Next
  CallByName mParent, "ComboClicked", VbMethod, Combo.name
  On Error GoTo 0
End Sub

Private Sub Combo_DropButtonClick()
  On Error Resume Next
  CallByName mParent, "ComboClicked", VbMethod, Combo.name
  On Error GoTo 0
End Sub

' Typed input: UI-only filtering of dropdown items (forms may optionally implement ComboTyped)
Private Sub Combo_Change()
  On Error Resume Next
  CallByName mParent, "ComboTyped", VbMethod, Combo.name, Combo.Text
  On Error GoTo 0
End Sub

' Enter key confirms selection (optional UX)
Private Sub Combo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  On Error Resume Next
  If KeyCode = vbKeyReturn Then
    CallByName mParent, "FiltersChanged", VbMethod
  End If
  On Error GoTo 0
End Sub
