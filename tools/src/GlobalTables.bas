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

Public Sub LoadAllGlobalTables()
    LoadPokemonTable
    LoadLearnsetsTable
    LoadMovesTable
    LoadItemsTable
    LoadAbilitiesTable
    LoadNaturesTable
    LoadTypeChartTable
    LoadGameversionsTable
End Sub

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
    lastRow = lastRowCell.Row
    lastCol = lastColCell.Column

    Set GetSheetTableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Private Function TableHasData(ByRef tableArr As Variant) As Boolean
    TableHasData = Not IsEmpty(tableArr)
End Function
