VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGridHeaderLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents HeaderLabel As MSForms.label
Attribute HeaderLabel.VB_VarHelpID = -1

Private mParent As Object
Private mColIndex As Long

Public Sub Init(ByVal parentForm As Object, ByVal lbl As MSForms.label, ByVal colIndex As Long)
    Set mParent = parentForm
    Set HeaderLabel = lbl
    mColIndex = colIndex
End Sub

Private Sub HeaderLabel_Click()
    On Error Resume Next
    CallByName mParent, "HeaderClicked", VbMethod, mColIndex
    On Error GoTo 0
End Sub
