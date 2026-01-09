Attribute VB_Name = "Buttons"
' === Buttons.bas ===
Option Explicit

Sub AssignTypes()
    ' No auto-close: pokedata stays open (hidden) for the whole Excel session.
    On Error GoTo CleanFail

    Dim wsTypeChart As Worksheet
    ' Variables
    Dim Pokemon As String
    Dim move As String
    Dim pkmnType1 As String
    Dim pkmnType2 As String
    Dim MoveType As String

    ' Local sheet
    Set wsTypeChart = TypeChart

    ' Read inputs
    Pokemon = Trim$(CStr(wsTypeChart.Range("PKMN").value))
    move = Trim$(CStr(wsTypeChart.Range("Move").value))

    ' Default outputs
    pkmnType1 = vbNullString
    pkmnType2 = vbNullString
    MoveType = vbNullString

    Dim pokemonNameCol As Long
    Dim type1Col As Long
    Dim type2Col As Long
    Dim moveNameCol As Long
    Dim moveTypeCol As Long

    GlobalTables.LoadPokemonTable
    GlobalTables.LoadMovesTable

    If Not IsEmpty(GlobalTables.PokemonTable) Then
        pokemonNameCol = GlobalTables.FindHeaderColumn(GlobalTables.PokemonTable, "DISPLAY_NAME")
        type1Col = GlobalTables.FindHeaderColumn(GlobalTables.PokemonTable, "TYPE1")
        type2Col = GlobalTables.FindHeaderColumn(GlobalTables.PokemonTable, "TYPE2")
    End If

    If Not IsEmpty(GlobalTables.movesTable) Then
        moveNameCol = GlobalTables.FindHeaderColumn(GlobalTables.movesTable, "DISPLAY_NAME")
        moveTypeCol = GlobalTables.FindHeaderColumn(GlobalTables.movesTable, "TYPE")
    End If

    ' Pokemon lookup via cached table
    If Len(Pokemon) > 0 And pokemonNameCol > 0 Then
        Dim pokemonRow As Long
        pokemonRow = GlobalTables.FindRowByValue(GlobalTables.PokemonTable, pokemonNameCol, Pokemon)
        If pokemonRow > 0 Then
            If type1Col > 0 Then pkmnType1 = NzText(GlobalTables.PokemonTable(pokemonRow, type1Col))
            If type2Col > 0 Then pkmnType2 = NzText(GlobalTables.PokemonTable(pokemonRow, type2Col))
        End If
    End If

    ' Move lookup via cached table
    If Len(move) > 0 And moveNameCol > 0 Then
        Dim MoveRow As Long
        MoveRow = GlobalTables.FindRowByValue(GlobalTables.movesTable, moveNameCol, move)
        If MoveRow > 0 And moveTypeCol > 0 Then
            MoveType = NzText(GlobalTables.movesTable(MoveRow, moveTypeCol))
        End If
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

