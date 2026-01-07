Attribute VB_Name = "Buttons"
' === Buttons.bas ===
Option Explicit

Sub AssignTypes()
    ' No auto-close: pokedata stays open (hidden) for the whole Excel session.
    On Error GoTo CleanFail

    Dim wsTypeChart As Worksheet
    Dim wsMoves As Worksheet
    Dim wsPokemons As Worksheet
    Dim wbPokedata As Workbook

    ' Variables
    Dim pokemon As String
    Dim move As String
    Dim pkmnType1 As String
    Dim pkmnType2 As String
    Dim MoveType As String

    ' Local sheet
    Set wsTypeChart = TypeChart

    ' External workbook + sheets (auto-opens in background if needed)
    Set wbPokedata = Functions.GetPokedataWb
    Set wsMoves = wbPokedata.Worksheets("Moves")
    Set wsPokemons = wbPokedata.Worksheets("Pokemon")

    ' Read inputs
    pokemon = Trim$(CStr(wsTypeChart.Range("PKMN").value))
    move = Trim$(CStr(wsTypeChart.Range("Move").value))

    ' Default outputs
    pkmnType1 = vbNullString
    pkmnType2 = vbNullString
    MoveType = vbNullString

    ' Pokemon lookup: table C:F, return E (col 3) and F (col 4)
    If Len(pokemon) > 0 Then
        pkmnType1 = NzText(Application.VLookup(pokemon, wsPokemons.Range("C:F"), 3, False)) ' E
        pkmnType2 = NzText(Application.VLookup(pokemon, wsPokemons.Range("C:F"), 4, False)) ' F
    End If

    ' Move lookup: table B:C, return C (col 2)
    If Len(move) > 0 Then
        MoveType = NzText(Application.VLookup(move, wsMoves.Range("B:C"), 2, False)) ' C
    End If

    ' Write outputs back (only when found / non-empty as requested)
    If pkmnType1 <> "" Then
        wsTypeChart.Range("PKMN_TYPE_1").value = pkmnType1
        wsTypeChart.Range("PKMN_TYPE_2").value = pkmnType2  ' may be empty
    End If

    If MoveType <> "" Then
        wsTypeChart.Range("MOVE_TYPE").value = MoveType
    End If

    Exit Sub

CleanFail:
    ' Fail silently (or replace with MsgBox for debugging)
    ' MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub


' Converts VLOOKUP result into a safe string:
' - returns "" if error
' - trims whitespace
' - converts to Proper Case (first letter of each word uppercase)
Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Then
        NzText = vbNullString
    Else
        NzText = StrConv(Trim$(CStr(v)), vbProperCase)
    End If
End Function

