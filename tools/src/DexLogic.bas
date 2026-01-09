Attribute VB_Name = "DexLogic"
Option Explicit

' =====================================================================================
' DexLogic
' Refactor (2026-01):
' - pokedata workbook now stores movesets directly on sheet "Pokemon" in columns:
'     * Column C: Pokemon name
'     * Row 1: headers that include MOVESET_<game-version> (e.g., MOVESET_scarlet-violet)
' - PKMN_DEX list:
'     * For game "All": all non-empty Pokemon names from Pokemon!C
'     * For a specific game: only Pokemon where MOVESET_<game> cell is NOT empty
' - PKMN_MOVELIST list:
'     * For game "All": built from pokedata sheet "Learnsets" (all moves for the Pokemon)
'     * For a specific game: built from the selected Pokemon's MOVESET_<game> cell
'     * Moves are delimited by ';' (e.g., amnesia;attract;body-slam)
' - Normalization preserved:
'     * game version normalized via NormalizeGameVersion()
'     * move text normalized via NormalizeMoveText()
' - Performance:
'     * Skip "fake changes" (re-selecting the same game or same game+pokemon)
' =====================================================================================

' Tmp Pokemon list (Lists sheet)
Private Const TMP_COL As String = "O"
Private Const TMP_START_ROW As Long = 2
Private Const TMP_HEADER As String = "TmpPokmons"

' Tmp Move list (Lists sheet)
Private Const TMP_MOVE_COL As String = "P"
Private Const TMP_MOVE_START_ROW As Long = 2
Private Const TMP_MOVE_HEADER As String = "TmpMovelist"

' Cache per worksheet (keyed by ws.CodeName)
Private mLastGame As Object      ' Scripting.Dictionary: wsKey -> normalized game version
Private mLastDexKey As Object    ' Scripting.Dictionary: wsKey -> (game|pokemon)

Private Sub EnsureCaches()
    If mLastGame Is Nothing Then
        Set mLastGame = CreateObject("Scripting.Dictionary")
        mLastGame.CompareMode = vbTextCompare
    End If
    If mLastDexKey Is Nothing Then
        Set mLastDexKey = CreateObject("Scripting.Dictionary")
        mLastDexKey.CompareMode = vbTextCompare
    End If
End Sub

' =====================================================================================
' Public entry points (called from Pokedex Worksheet_Change)
' =====================================================================================

Public Sub HandleGameChange(ByVal ws As Worksheet)
    On Error GoTo CleanFail
        
    EnsureCaches

    Dim rngGame As Range, rngDex As Range
    Set rngGame = ws.Range("GAME")
    Set rngDex = ws.Range("PKMN_DEX")

    Dim gameRaw As String
    gameRaw = Trim$(CStr(rngGame.value))

    ' Empty => All
    If gameRaw = vbNullString Then
        rngGame.value = "All"
        gameRaw = "All"
    End If

    Dim gameVersion As String
    gameVersion = NormalizeGameVersion(gameRaw)

    ' Skip if game didn't actually change (prevents work when only Pokemon changed)
    Dim wsKey As String
    wsKey = ws.CodeName

    If mLastGame.Exists(wsKey) Then
        If StrComp(CStr(mLastGame(wsKey)), gameVersion, vbTextCompare) = 0 Then Exit Sub
    End If
    mLastGame(wsKey) = gameVersion

    ' Build Pokemon list for this game
    Dim pkmnList As Variant
    pkmnList = GetPokemonListForGame(gameVersion)

    ' If no data -> disable
    If IsEmpty(pkmnList) Then
        On Error Resume Next
        rngDex.Validation.Delete
        On Error GoTo 0
        rngDex.value = "-"
        Exit Sub
    End If

    ' Apply validation list
    SetDexValidationFromArray rngDex, pkmnList

    ' Keep selection valid
    Dim currentPkmn As String
    currentPkmn = Trim$(CStr(rngDex.value))

    If Len(currentPkmn) = 0 Or Not IsInArray(currentPkmn, pkmnList) Then
        rngDex.value = CStr(pkmnList(LBound(pkmnList)))
    End If

    Exit Sub

CleanFail:
    ' Fail silently (or MsgBox for debugging)
End Sub

Public Sub HandleMoveListRefresh(ByVal ws As Worksheet)
    On Error GoTo CleanFail
    EnsureCaches

    Dim rngGame As Range, rngDex As Range, rngMove As Range
    Set rngGame = ws.Range("GAME")
    Set rngDex = ws.Range("PKMN_DEX")
    Set rngMove = ws.Range("PKMN_MOVELIST")

    Dim gameRaw As String
    gameRaw = Trim$(CStr(rngGame.value))
    If gameRaw = vbNullString Then
        rngGame.value = "All"
        gameRaw = "All"
    End If

    Dim gameVersion As String
    gameVersion = NormalizeGameVersion(gameRaw)

    Dim pkmn As String
    pkmn = Trim$(CStr(rngDex.value))

    ' Skip if nothing relevant changed (fake re-select)
    Dim wsKey As String
    wsKey = ws.CodeName

    Dim dexKey As String
    dexKey = gameVersion & "|" & pkmn

    If mLastDexKey.Exists(wsKey) Then
        If StrComp(CStr(mLastDexKey(wsKey)), dexKey, vbTextCompare) = 0 Then Exit Sub
    End If
    mLastDexKey(wsKey) = dexKey

    ' Build move list
    Dim Movelist As Variant
    GetMoveListForPokemon gameVersion, pkmn, Movelist

    ' Write tmp lists and set validation
    SetMoveValidationFromArrays rngMove, Movelist
    ' Keep selection valid
    Dim currentMove As String
    currentMove = Trim$(CStr(rngMove.value))

    If Len(currentMove) = 0 Or Not IsInArray(currentMove, Movelist) Then
        rngMove.value = CStr(Movelist(LBound(Movelist)))
    End If

    Exit Sub

CleanFail:
    ' Fail silently (or MsgBox for debugging)
End Sub

' =====================================================================================
' Game / data helpers
' =====================================================================================

Public Function NormalizeGameVersion(ByVal s As String) As String
    Dim t As String
    t = Trim$(CStr(s))

    If Len(t) = 0 Then
        NormalizeGameVersion = vbNullString
        Exit Function
    End If

    t = Replace(t, " ", "-")
    t = Replace(t, "&", "-")
    t = Replace(t, ChrW(&H2019), "")

    ' Collapse multiple hyphens
    Do While InStr(1, t, "--", vbBinaryCompare) > 0
        t = Replace(t, "--", "-")
    Loop

    ' Trim leading/trailing hyphens
    Do While Left$(t, 1) = "-"
        t = Mid$(t, 2)
    Loop
    Do While Right$(t, 1) = "-"
        t = Left$(t, Len(t) - 1)
    Loop

    NormalizeGameVersion = t
End Function

Public Function DisplayNameFromGameKey(ByVal value As Variant) As String
    Dim raw As String
    If IsError(value) Or IsNull(value) Then
        DisplayNameFromGameKey = "All"
        Exit Function
    End If
    On Error GoTo CleanFallback
    raw = Trim$(CStr(value))

    Dim norm As String
    norm = NormalizeGameVersion(raw)

    Dim friendly As String
    If Len(norm) = 0 Then
        friendly = raw
    Else
        friendly = norm
    End If

    If Len(friendly) = 0 Then
        DisplayNameFromGameKey = "All"
        Exit Function
    End If

    friendly = Replace(friendly, "-", " ")
    friendly = Replace(friendly, "_", " ")

    Do While InStr(friendly, "  ") > 0
        friendly = Replace(friendly, "  ", " ")
    Loop

    DisplayNameFromGameKey = StrConv(Trim$(friendly), vbProperCase)
    Exit Function

CleanFallback:
    DisplayNameFromGameKey = "All"
End Function
Private Function GetPokemonListForGame(ByVal gameVersion As String) As Variant
    On Error GoTo CleanFail

    GlobalTables.LoadGameversionsTable
    If IsEmpty(GlobalTables.GameversionsTable) Then Exit Function

    Dim headerName As String
    If StrComp(gameVersion, "All", vbTextCompare) = 0 Then
        headerName = "POKEMON_ALL"
    Else
        headerName = "POKEMON_" & gameVersion
    End If

    Dim columnIndex As Long
    columnIndex = GlobalTables.FindHeaderColumn(GlobalTables.GameversionsTable, headerName)
    If columnIndex = 0 Then Exit Function

    Dim rawValues As Variant
    rawValues = GlobalTables.ExtractColumnValues(GlobalTables.GameversionsTable, columnIndex, True)
    If IsEmpty(rawValues) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim cleaned() As String
    Dim n As Long
    n = 0

    Dim i As Long
    Dim valueText As String
    For i = LBound(rawValues) To UBound(rawValues)
        valueText = Trim$(CStr(rawValues(i)))
        If Len(valueText) > 0 And valueText <> "0" Then
            If Not dict.Exists(valueText) Then
                dict.Add valueText, True
                n = n + 1
                If n = 1 Then
                    ReDim cleaned(1 To 1)
                Else
                    ReDim Preserve cleaned(1 To n)
                End If
                cleaned(n) = valueText
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    GetPokemonListForGame = cleaned
    Exit Function

CleanFail:
    ' Fail silently
End Function

Public Function FindMovesetColumn(ByVal wsPokemon As Worksheet, ByVal gameVersion As String) As Long
    ' Compatibility wrapper: column lookup now uses cached Pokemon table, wsPokemon is unused
    GlobalTables.LoadPokemonTable
    If IsEmpty(GlobalTables.PokemonTable) Then Exit Function

    Dim header As String
    header = "MOVESET_" & gameVersion

    FindMovesetColumn = GlobalTables.FindHeaderColumn(GlobalTables.PokemonTable, header)
End Function

Private Function FindPokemonRow(ByVal pokemonName As String) As Long
    GlobalTables.LoadPokemonTable
    If IsEmpty(GlobalTables.PokemonTable) Then Exit Function

    Dim nameCol As Long
    nameCol = GlobalTables.FindHeaderColumn(GlobalTables.PokemonTable, "DISPLAY_NAME")
    If nameCol = 0 Then Exit Function

    Dim firstRow As Long
    Dim lastRow As Long
    firstRow = LBound(GlobalTables.PokemonTable, 1) + 1 ' skip header
    lastRow = UBound(GlobalTables.PokemonTable, 1)

    Dim r As Long
    For r = firstRow To lastRow
        If StrComp(Trim$(CStr(GlobalTables.PokemonTable(r, nameCol))), pokemonName, vbTextCompare) = 0 Then
            FindPokemonRow = r
            Exit Function
        End If
    Next r
End Function

' =====================================================================================
' Validation helpers (Dex list)
' =====================================================================================

Private Sub SetDexValidationFromArray(ByVal rngDex As Range, ByVal values As Variant)
    Dim wsLists As Worksheet
    Set wsLists = Lists ' codename

    ' Header
    wsLists.Cells(1, TMP_COL).value = TMP_HEADER

    ' Clear old list
    wsLists.Range(wsLists.Cells(TMP_START_ROW, TMP_COL), _
                  wsLists.Cells(wsLists.Rows.Count, TMP_COL)).ClearContents

    Dim n As Long
    n = UBound(values) - LBound(values) + 1
    If n <= 0 Then Exit Sub

    ' Write list
    Dim outRng As Range
    Set outRng = wsLists.Range(wsLists.Cells(TMP_START_ROW, TMP_COL), _
                               wsLists.Cells(TMP_START_ROW + n - 1, TMP_COL))

    Dim i As Long
    For i = 1 To n
        outRng.Cells(i, 1).value = values(LBound(values) + i - 1)
    Next i
    
    Dim src As String

    ' Apply validation
    On Error Resume Next
        src = "=" & outRng.Address(External:=True)   ' => ='Pokedex.xlsm'!Lists!$Z$2:$Z$50 apod.
        
        With rngDex.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=src
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    On Error GoTo 0

End Sub

Private Function IsInArray(ByVal value As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(i)), value, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

' =====================================================================================
' Move list logic (from Pokemon!MOVESET_<game>)
' =====================================================================================

Private Sub GetAllMovesFromLearnsets(ByVal pokemonName As String, _
                                  ByRef outMoves As Variant)
    On Error GoTo CleanFail

    ' Default output
    Dim defArr(1 To 1) As String
    defArr(1) = "-"
    outMoves = defArr

    GlobalTables.LoadLearnsetsTable
    If IsEmpty(GlobalTables.LearnsetsTable) Then Exit Sub

    Dim nameCol As Long
    Dim moveCol As Long
    nameCol = GlobalTables.FindHeaderColumn(GlobalTables.LearnsetsTable, "DISPLAY_NAME")
    moveCol = GlobalTables.FindHeaderColumn(GlobalTables.LearnsetsTable, "MOVE_KEY")
    If nameCol = 0 Or moveCol = 0 Then Exit Sub

    Dim firstRow As Long
    Dim lastRow As Long
    firstRow = LBound(GlobalTables.LearnsetsTable, 1) + 1
    lastRow = UBound(GlobalTables.LearnsetsTable, 1)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim r As Long
    Dim mv As String

    For r = firstRow To lastRow
        If StrComp(Trim$(CStr(GlobalTables.LearnsetsTable(r, nameCol))), pokemonName, vbTextCompare) = 0 Then
            mv = NormalizeMoveText(Trim$(CStr(GlobalTables.LearnsetsTable(r, moveCol))))
            If Len(mv) > 0 Then
                If Not dict.Exists(mv) Then
                    dict.Add mv, True
                End If
            End If
        End If
    Next r

    If dict.Count = 0 Then Exit Sub

    Dim moves() As String
    ReDim moves(1 To dict.Count)

    Dim i As Long
    i = 1
    Dim key As Variant
    For Each key In dict.Keys
        moves(i) = CStr(key)
        i = i + 1
    Next key

    outMoves = moves
    Exit Sub

CleanFail:
    ' Fail silently
End Sub


Private Sub GetMoveListForPokemon(ByVal gameVersion As String, ByVal pokemonName As String, _
                                 ByRef outMoves As Variant)
    On Error GoTo CleanFail

    ' Default output
    Dim defArr(1 To 1) As String
    defArr(1) = "-"
    outMoves = defArr

    If Len(Trim$(pokemonName)) = 0 Then Exit Sub

    ' For GAME = All, return all moves for the Pokemon from Learnsets sheet
    If StrComp(gameVersion, "All", vbTextCompare) = 0 Then
        GetAllMovesFromLearnsets pokemonName, outMoves
        Exit Sub
    End If

    GlobalTables.LoadPokemonTable
    If IsEmpty(GlobalTables.PokemonTable) Then Exit Sub

    Dim movesetCol As Long
    movesetCol = FindMovesetColumn(Nothing, gameVersion)
    If movesetCol = 0 Then Exit Sub

    Dim rowP As Long
    rowP = FindPokemonRow(pokemonName)
    If rowP = 0 Then Exit Sub

    Dim movesetRaw As String
    movesetRaw = Trim$(CStr(GlobalTables.PokemonTable(rowP, movesetCol)))
    If Len(movesetRaw) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(movesetRaw, ";")

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim moves() As String
    Dim n As Long: n = 0

    Dim i As Long
    Dim token As String
    Dim mv As String

    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            mv = NormalizeMoveText(token)
            If Len(mv) > 0 Then
                If Not dict.Exists(mv) Then
                    dict.Add mv, True
                    n = n + 1
                    ReDim Preserve moves(1 To n)
                    moves(n) = mv
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Sub
    outMoves = moves
    Exit Sub

CleanFail:
    ' Fail silently, keep defaults
End Sub

Private Sub SetMoveValidationFromArrays(ByVal rngMove As Range, ByVal moves As Variant)
    Dim wsLists As Worksheet
    Set wsLists = Lists ' codename

    ' Headers
    wsLists.Cells(1, TMP_MOVE_COL).value = TMP_MOVE_HEADER

    ' Clear old lists
    wsLists.Range(wsLists.Cells(TMP_MOVE_START_ROW, TMP_MOVE_COL), _
                  wsLists.Cells(wsLists.Rows.Count, TMP_MOVE_COL)).ClearContents

    Dim n As Long
    n = UBound(moves) - LBound(moves) + 1
    If n <= 0 Then Exit Sub

    Dim outMovesRng As Range
    Set outMovesRng = wsLists.Range(wsLists.Cells(TMP_MOVE_START_ROW, TMP_MOVE_COL), _
                                    wsLists.Cells(TMP_MOVE_START_ROW + n - 1, TMP_MOVE_COL))


    Dim i As Long
    For i = 1 To n
        outMovesRng.Cells(i, 1).value = moves(LBound(moves) + i - 1)
    Next i

    Dim srcMoves As String

    ' Apply validation list to PKMN_MOVELIST (moves only)
    On Error Resume Next
    
        srcMoves = "=" & outMovesRng.Address(External:=True)
        
        With rngMove.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, formula1:=srcMoves
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        
    On Error GoTo 0
End Sub

Private Function NormalizeMoveText(ByVal s As String) As String
    Dim t As String
    t = Trim$(CStr(s))

    If Len(t) = 0 Then
        NormalizeMoveText = vbNullString
        Exit Function
    End If

    t = Replace(t, "-", " ")
    t = Application.WorksheetFunction.Proper(t)

    NormalizeMoveText = t
End Function

Public Sub HandleAbilityRefresh(ByVal ws As Worksheet)
    On Error GoTo CleanFail

    GlobalTables.LoadAbilitiesTable
    If IsEmpty(GlobalTables.AbilitiesTable) Then Exit Sub

    RefreshAbilityNote ws.Range("ABILITY_1")
    RefreshAbilityNote ws.Range("ABILITY_2")
    RefreshAbilityNote ws.Range("HIDDEN_ABILITY")
    Exit Sub

CleanFail:
    ' Fail silently
End Sub

Private Sub RefreshAbilityNote(ByVal rng As Range)
    On Error GoTo CleanFail

    Dim abilityName As String
    abilityName = Trim$(CStr(rng.Value2))

    ' empty -> clear note
    If Len(abilityName) = 0 Or abilityName = "-" Then
        ClearNote rng
        Exit Sub
    End If

    Dim desc As String
    desc = GetAbilityDescription(abilityName)

    ' no match -> clear note
    If Len(desc) = 0 Then
        ClearNote rng
        Exit Sub
    End If

    SetNote rng, desc
    Exit Sub

CleanFail:
    ' Fail silently
End Sub

Private Function GetAbilityDescription(ByVal abilityName As String) As String
    On Error GoTo CleanFail

    GlobalTables.LoadAbilitiesTable
    If IsEmpty(GlobalTables.AbilitiesTable) Then Exit Function

    Dim nameCol As Long
    Dim descCol As Long
    nameCol = GlobalTables.FindHeaderColumn(GlobalTables.AbilitiesTable, "DISPLAY_NAME")
    descCol = GlobalTables.FindHeaderColumn(GlobalTables.AbilitiesTable, "EFFECT_SHORT")
    If nameCol = 0 Or descCol = 0 Then Exit Function

    Dim rowIdx As Long
    rowIdx = GlobalTables.FindRowByValue(GlobalTables.AbilitiesTable, nameCol, abilityName)
    If rowIdx = 0 Then Exit Function

    GetAbilityDescription = Trim$(CStr(GlobalTables.AbilitiesTable(rowIdx, descCol)))
    Exit Function

CleanFail:
    ' empty
End Function

Private Sub ClearNote(ByVal rng As Range)
    On Error Resume Next
    If Not rng.Comment Is Nothing Then rng.Comment.Delete
    On Error GoTo 0
End Sub

Private Sub SetNote(ByVal rng As Range, ByVal noteText As String)
    On Error GoTo CleanFail

    ' Delete existing, then add clean
    If Not rng.Comment Is Nothing Then rng.Comment.Delete

    rng.AddComment Text:=noteText
    rng.Comment.Visible = False

    ' Format: bold text
    On Error Resume Next
    With rng.Comment.Shape
        .TextFrame.Characters.Font.Bold = True
        .Fill.Visible = msoTrue
        .Fill.Solid
    End With
    On Error GoTo 0

    Exit Sub

CleanFail:
    ' Fail silently
End Sub

