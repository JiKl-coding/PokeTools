Attribute VB_Name = "MoveDexLogic"
Option Explicit

Private Const MOVELIST_COL As String = "R"
Private Const MOVELIST_START_ROW As Long = 2
Private Const MOVELIST_HEADER As String = "tmpMoves"

Public Sub HandleWorksheetChange(ByVal ws As Worksheet, ByVal target As Range)
    Dim rngGame As Range
    Dim rngList As Range
    Set rngGame = ws.Range("GAMEVERSION")
    Set rngList = ws.Range("MVLIST")
    
    If Not Intersect(target, rngGame) Is Nothing Then
        RefreshMoveListForGame rngGame, rngList
    ElseIf Not Intersect(target, rngList) Is Nothing Then
        EnsureMoveSelectionValid rngGame, rngList
    End If
End Sub

Private Sub RefreshMoveListForGame(ByVal rngGame As Range, ByVal rngList As Range)
    Dim previousSelection As String
    previousSelection = Trim$(CStr(rngList.Value2))

    Dim normalizedGame As String
    normalizedGame = NormalizeGameCell(rngGame)

    Dim moves As Variant
    moves = BuildMoveList(normalizedGame)

    Dim cacheRange As Range
    Set cacheRange = WriteMovesToLists(moves)
    ApplyMoveValidation rngList, cacheRange
    RestoreSelection rngList, moves, previousSelection
End Sub

Private Sub EnsureMoveSelectionValid(ByVal rngGame As Range, ByVal rngList As Range)
    Dim moves As Variant
    moves = ReadCachedMoveList()

    If IsEmpty(moves) Then
        Dim normalizedGame As String
        normalizedGame = NormalizeGameCell(rngGame)
        moves = BuildMoveList(normalizedGame)
        Dim cacheRange As Range
        Set cacheRange = WriteMovesToLists(moves)
        ApplyMoveValidation rngList, cacheRange
    End If

    EnsureValueFromList rngList, moves

    If Not IsEmpty(moves) Then
        If Len(Trim$(CStr(rngList.Value2))) = 0 Then
            rngList.Value2 = moves(LBound(moves))
        End If
    End If
End Sub

Private Function NormalizeGameCell(ByVal rngGame As Range) As String
    Dim raw As String
    raw = Trim$(CStr(rngGame.Value2))
    If Len(raw) = 0 Then
        rngGame.Value2 = "All"
        raw = "All"
    End If

    Dim normalized As String
    normalized = SafeNormalizeGameVersion(raw)
    If Len(normalized) = 0 Then
        normalized = "All"
    End If

    NormalizeGameCell = normalized
End Function

Private Function SafeNormalizeGameVersion(ByVal value As String) As String
    On Error GoTo CleanFallback
    SafeNormalizeGameVersion = DexLogic.NormalizeGameVersion(value)
    Exit Function
CleanFallback:
    SafeNormalizeGameVersion = Trim$(CStr(value))
End Function

Private Function BuildMoveList(ByVal gameVersion As String) As Variant
    Dim values As Variant
    values = ExtractMovesFromTable(gameVersion)

    If IsEmpty(values) Then
        Dim fallback(1 To 1) As String
        fallback(1) = "-"
        BuildMoveList = fallback
    Else
        BuildMoveList = values
    End If
End Function

Private Function ExtractMovesFromTable(ByVal gameVersion As String) As Variant
    On Error GoTo CleanFail

    GlobalTables.LoadGameversionsTable
    If IsEmpty(GlobalTables.GameversionsTable) Then Exit Function

    Dim headerName As String
    If StrComp(gameVersion, "All", vbTextCompare) = 0 Then
        headerName = "MOVES_ALL"
    Else
        headerName = "MOVES_" & gameVersion
    End If

    Dim columnIndex As Long
    columnIndex = GlobalTables.FindHeaderColumn(GlobalTables.GameversionsTable, headerName)
    If columnIndex = 0 Then
        columnIndex = GlobalTables.FindHeaderColumn(GlobalTables.GameversionsTable, "MOVES_ALL")
    End If
    If columnIndex = 0 Then Exit Function

    Dim rawValues As Variant
    rawValues = GlobalTables.ExtractColumnValues(GlobalTables.GameversionsTable, columnIndex, True)
    If IsEmpty(rawValues) Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim cleaned() As String
    Dim count As Long

    Dim valueText As String
    Dim i As Long
    For i = LBound(rawValues) To UBound(rawValues)
        valueText = Trim$(CStr(rawValues(i)))
        If Len(valueText) > 0 And valueText <> "0" Then
            If Not dict.Exists(valueText) Then
                dict.Add valueText, True
                count = count + 1
                If count = 1 Then
                    ReDim cleaned(1 To 1)
                Else
                    ReDim Preserve cleaned(1 To count)
                End If
                cleaned(count) = valueText
            End If
        End If
    Next i

    If count = 0 Then Exit Function
    ExtractMovesFromTable = cleaned
    Exit Function

CleanFail:
    ' Leave Empty
End Function

Private Function WriteMovesToLists(ByVal moves As Variant) As Range
    Dim wsLists As Worksheet
    Set wsLists = Lists

    wsLists.Cells(1, MOVELIST_COL).value = MOVELIST_HEADER
    wsLists.Range(wsLists.Cells(MOVELIST_START_ROW, MOVELIST_COL), _
                  wsLists.Cells(wsLists.Rows.count, MOVELIST_COL)).ClearContents

    If IsEmpty(moves) Then Exit Function

    Dim count As Long
    count = UBound(moves) - LBound(moves) + 1
    If count <= 0 Then Exit Function

    Dim outRng As Range
    Set outRng = wsLists.Range(wsLists.Cells(MOVELIST_START_ROW, MOVELIST_COL), _
                               wsLists.Cells(MOVELIST_START_ROW + count - 1, MOVELIST_COL))

    Dim i As Long
    For i = 1 To count
        outRng.Cells(i, 1).value = moves(LBound(moves) + i - 1)
    Next i

    Set WriteMovesToLists = outRng
End Function

Private Sub ApplyMoveValidation(ByVal rngTarget As Range, ByVal sourceRange As Range)
    If sourceRange Is Nothing Then Exit Sub

    Dim formula As String
    formula = "=" & sourceRange.Address(External:=True)

    On Error Resume Next
    With rngTarget.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=formula
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Sub RestoreSelection(ByVal rngList As Range, ByVal moves As Variant, ByVal desiredValue As String)
    If IsEmpty(moves) Then Exit Sub

    Dim fallback As String
    fallback = CStr(moves(LBound(moves)))

    If Len(desiredValue) = 0 Then
        rngList.Value2 = fallback
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(moves) To UBound(moves)
        If StrComp(CStr(moves(i)), desiredValue, vbTextCompare) = 0 Then
            rngList.Value2 = moves(i)
            Exit Sub
        End If
    Next i

    rngList.Value2 = fallback
End Sub

Private Sub EnsureValueFromList(ByVal rngList As Range, ByVal moves As Variant)
    If IsEmpty(moves) Then Exit Sub

    Dim fallback As String
    fallback = CStr(moves(LBound(moves)))

    Dim current As String
    current = Trim$(CStr(rngList.Value2))
    If Len(current) = 0 Then
        rngList.Value2 = fallback
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(moves) To UBound(moves)
        If StrComp(CStr(moves(i)), current, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next i

    rngList.Value2 = fallback
End Sub

Private Function ReadCachedMoveList() As Variant
    On Error GoTo CleanFail

    Dim wsLists As Worksheet
    Set wsLists = Lists

    Dim colIndex As Long
    colIndex = wsLists.Columns(MOVELIST_COL).Column

    Dim lastRow As Long
    lastRow = wsLists.Cells(wsLists.Rows.count, colIndex).End(xlUp).row
    If lastRow < MOVELIST_START_ROW Then Exit Function

    Dim values() As Variant
    Dim count As Long
    Dim r As Long
    Dim cellValue As String

    For r = MOVELIST_START_ROW To lastRow
        cellValue = Trim$(CStr(wsLists.Cells(r, colIndex).value))
        If Len(cellValue) > 0 Then
            count = count + 1
            If count = 1 Then
                ReDim values(1 To 1)
            Else
                ReDim Preserve values(1 To count)
            End If
            values(count) = cellValue
        End If
    Next r

    If count = 0 Then Exit Function
    ReadCachedMoveList = values
    Exit Function

CleanFail:
    ' Leave Empty
End Function
