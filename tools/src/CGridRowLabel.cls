VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGridRowLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForm As Movelist
Private mRowIndex As Long
Private WithEvents mLbl As MSForms.label
Attribute mLbl.VB_VarHelpID = -1

Public Sub Init(ByVal frm As Movelist, ByVal lbl As MSForms.label, ByVal rowIndex As Long)
    Set mForm = frm
    Set mLbl = lbl
    mRowIndex = rowIndex
End Sub

Private Sub mLbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If Not mForm Is Nothing Then mForm.OnRowDoubleClick mRowIndex
    On Error GoTo 0
    Cancel = True
End Sub
