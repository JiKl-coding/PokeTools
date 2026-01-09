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

Private Sub Worksheet_Change(ByVal target As Range)
    ' React only to GAME or PKMN_DEX
    If Intersect(target, Union(Me.Range("GAME"), Me.Range("PKMN_DEX"))) Is Nothing Then Exit Sub

    On Error GoTo CleanFail
    Application.EnableEvents = False
    
    Functions.UnprotectForMacros Pokedex

    DexLogic.HandleGameChange Me        ' refresh PKMN_DEX list (and keep selection valid)
    DexLogic.HandleMoveListRefresh Me   ' refresh PKMN_MOVELIST list (and keep selection valid)
    DexLogic.HandleAbilityRefresh Me    ' refresh ABILITY_1, ABILITY_2 AND HIDDEN_ABILITY description note
    
CleanExit:
    Functions.ProtectForMacros Pokedex
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    ' Fail silently (or MsgBox for debugging)
    Resume CleanExit
End Sub
