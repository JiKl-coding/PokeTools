VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pokedex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' React only to GAME or PKMN_DEX
    If Intersect(Target, Union(Me.Range("GAME"), Me.Range("PKMN_DEX"))) Is Nothing Then Exit Sub

    On Error GoTo CleanFail
    Application.EnableEvents = False
    
    Functions.UnprotectForMacros Pokedex
    Functions.UnprotectForMacros Lists

    DexLogic.HandleGameChange Me        ' refresh PKMN_DEX list (and keep selection valid)
    DexLogic.HandleMoveListRefresh Me   ' refresh PKMN_MOVELIST list (and keep selection valid)

CleanExit:
    Functions.ProtectForMacros Pokedex
    Functions.ProtectForMacros Lists
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    ' Fail silently (or MsgBox for debugging)
    Resume CleanExit
End Sub
