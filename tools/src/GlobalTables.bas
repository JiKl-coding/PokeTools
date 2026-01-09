Attribute VB_Name = "GlobalTables"
Option Explicit

' ===============================================================
' Global tables cache (lazy-loaded on demand)
' Each loader copies the full sheet (headers + rows) into a Variant
' array via Range.Value2 so downstream VBA can work on in-memory data
' without reopening the workbook.
' ===============================================================

Public PokemonTable As Variant
Public LearnsetsTable As Variant
Public MovesTable As Variant
Public ItemsTable As Variant
Public AbilitiesTable As Variant
Public NaturesTable As Variant
Public TypeChartTable As Variant
Public GameversionsTable As Variant
Public AssetsTable As Variant

Public Sub LoadPokemonTable()
    If TableHasData(PokemonTable) Then Exit Sub
    LoadSheetIntoArray "Pokemon", PokemonTable
End Sub

Public Sub LoadLearnsetsTable()
    If TableHasData(LearnsetsTable) Then Exit Sub
    LoadSheetIntoArray "Learnsets", LearnsetsTable
End Sub

Public Sub LoadMovesTable()
    If TableHasData(MovesTable) Then Exit Sub
    LoadSheetIntoArray "Moves", MovesTable
End Sub

Public Sub LoadItemsTable()
    If TableHasData(ItemsTable) Then Exit Sub
    LoadSheetIntoArray "Items", ItemsTable
End Sub

Public Sub LoadAbilitiesTable()
    If TableHasData(AbilitiesTable) Then Exit Sub
    LoadSheetIntoArray "Abilities", AbilitiesTable
End Sub

Public Sub LoadNaturesTable()
    If TableHasData(NaturesTable) Then Exit Sub
    LoadSheetIntoArray "Natures", NaturesTable
End Sub

Public Sub LoadTypeChartTable()
    If TableHasData(TypeChartTable) Then Exit Sub
    LoadSheetIntoArray "TypeChart", TypeChartTable
End Sub

Public Sub LoadGameversionsTable()
    If TableHasData(GameversionsTable) Then Exit Sub
    LoadSheetIntoArray "GAMEVERSIONS", GameversionsTable
End Sub

Public Sub LoadAssetsTable()
    If TableHasData(AssetsTable) Then Exit Sub
    LoadSheetIntoArray "Assets", AssetsTable
End Sub

Public Sub LoadAllGlobalTables()
    LoadPokemonTable
    LoadLearnsetsTable
    LoadMovesTable
    LoadItemsTable
    LoadAbilitiesTable
    LoadNaturesTable
    LoadTypeChartTable
    LoadGameversionsTable
    LoadAssetsTable
End Sub

Public Function FindHeaderColumn(ByRef tableArr As Variant, ByVal headerName As String) As Long
    ' Returns 1-based column index for the given header (row 1 comparison, case-insensitive)
    If IsEmpty(tableArr) Then Exit Function

    Dim headerRow As Long
    headerRow = LBound(tableArr, 1)

    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = LBound(tableArr, 2)
    lastCol = UBound(tableArr, 2)

    Dim col As Long
    For col = firstCol To lastCol
        If StrComp(SafeCellText(tableArr(headerRow, col)), headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = col
            Exit Function
        End If
    Next col
End Function

Public Function FindRowByValue(ByRef tableArr As Variant, ByVal columnIndex As Long, _
                               ByVal targetValue As String) As Long
    ' Finds the first row (including header) whose column value matches targetValue
    If IsEmpty(tableArr) Then Exit Function

    Dim normalizedTarget As String
    normalizedTarget = SafeCellText(targetValue)
    If Len(normalizedTarget) = 0 Then Exit Function

    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = LBound(tableArr, 2)
    lastCol = UBound(tableArr, 2)

    If columnIndex < firstCol Or columnIndex > lastCol Then Exit Function

    Dim firstRow As Long
    Dim lastRow As Long
    firstRow = LBound(tableArr, 1)
    lastRow = UBound(tableArr, 1)

    Dim r As Long
    For r = firstRow To lastRow
        If StrComp(SafeCellText(tableArr(r, columnIndex)), normalizedTarget, vbTextCompare) = 0 Then
            FindRowByValue = r
            Exit Function
        End If
    Next r
End Function

Public Function ExtractColumnValues(ByRef tableArr As Variant, ByVal columnIndex As Long, _
                                    Optional ByVal skipHeader As Boolean = True) As Variant
    ' Builds a 1-D Variant array with all values from the specified column
    If IsEmpty(tableArr) Then Exit Function

    Dim firstCol As Long
    Dim lastCol As Long
    firstCol = LBound(tableArr, 2)
    lastCol = UBound(tableArr, 2)

    If columnIndex < firstCol Or columnIndex > lastCol Then Exit Function

    Dim startRow As Long
    Dim lastRow As Long
    startRow = LBound(tableArr, 1)
    lastRow = UBound(tableArr, 1)

    If skipHeader Then
        startRow = startRow + 1
    End If

    If startRow > lastRow Then Exit Function

    Dim values() As Variant
    Dim count As Long
    count = 0

    Dim r As Long
    For r = startRow To lastRow
        count = count + 1
        If count = 1 Then
            ReDim values(1 To 1)
        Else
            ReDim Preserve values(1 To count)
        End If
        values(count) = tableArr(r, columnIndex)
    Next r

    If count > 0 Then
        ExtractColumnValues = values
    End If
End Function

Public Sub TestData(ByVal tableArr As Variant)
    On Error Resume Next
    Dim r As Long, c As Long
    Dim rowsCount As Long, colsCount As Long

    If IsEmpty(tableArr) Then
        Debug.Print "[TestData] Empty array"
        Exit Sub
    End If

    rowsCount = UBound(tableArr, 1) - LBound(tableArr, 1) + 1
    colsCount = UBound(tableArr, 2) - LBound(tableArr, 2) + 1

    Debug.Print "[TestData] Rows=" & rowsCount & ", Cols=" & colsCount

    For r = LBound(tableArr, 1) To UBound(tableArr, 1)
        Dim rowVals() As String
        ReDim rowVals(LBound(tableArr, 2) To UBound(tableArr, 2))
        For c = LBound(tableArr, 2) To UBound(tableArr, 2)
            rowVals(c) = CStr(tableArr(r, c))
        Next c
        Debug.Print Join(rowVals, " | ")
    Next r
End Sub

Private Sub LoadSheetIntoArray(ByVal sheetName As String, ByRef target As Variant)
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Functions.GetPokedataWb
    Set ws = wb.Worksheets(sheetName)

    Dim tableRange As Range
    Set tableRange = GetSheetTableRange(ws)

    If tableRange Is Nothing Then
        target = Empty
    Else
        target = tableRange.Value2
    End If
    Exit Sub

CleanFail:
    target = Empty
End Sub

Private Function GetSheetTableRange(ByVal ws As Worksheet) As Range
    Dim lastRowCell As Range
    Dim lastColCell As Range

    On Error Resume Next
    Set lastRowCell = ws.Cells.Find(What:="*", LookIn:=xlValues, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    Set lastColCell = ws.Cells.Find(What:="*", LookIn:=xlValues, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastRowCell Is Nothing Or lastColCell Is Nothing Then Exit Function

    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = lastRowCell.row
    lastCol = lastColCell.Column

    Set GetSheetTableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Private Function TableHasData(ByRef tableArr As Variant) As Boolean
    TableHasData = Not IsEmpty(tableArr)
End Function

Private Function SafeCellText(ByVal valueVariant As Variant) As String
    If IsError(valueVariant) Or IsNull(valueVariant) Or IsEmpty(valueVariant) Then Exit Function
    SafeCellText = Trim$(CStr(valueVariant))
End Function
