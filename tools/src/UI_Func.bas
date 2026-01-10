Attribute VB_Name = "UI_Func"
Option Explicit

Public Function CleanSelection(ByVal rawValue As Variant, ByVal fallback As String) As String
    If IsError(rawValue) Or IsNull(rawValue) Then
        CleanSelection = fallback
        Exit Function
    End If

    Dim scalar As Variant
    scalar = rawValue

    If IsArray(scalar) Then
        Dim lb1 As Long
        Dim lb2 As Long
        lb1 = LBound(scalar, 1)
        lb2 = LBound(scalar, 2)
        On Error Resume Next
        scalar = scalar(lb1, lb2)
        On Error GoTo 0
    End If

    If IsError(scalar) Then GoTo CleanFallback

    Dim t As String
    On Error GoTo CleanFallback
    t = Trim$(CStr(scalar))
    On Error GoTo 0
    If Len(t) = 0 Then
        CleanSelection = fallback
    Else
        CleanSelection = t
    End If
    Exit Function

CleanFallback:
    CleanSelection = fallback
End Function

Public Function GameVersionKey(ByVal value As String, _
                               Optional ByVal allValue As String = "All", _
                               Optional ByVal allKey As String = "__all__") As String
    Dim norm As String
    norm = DexLogic.NormalizeGameVersion(CleanSelection(value, allValue))
    If Len(norm) = 0 Or StrComp(norm, allValue, vbTextCompare) = 0 Then
        GameVersionKey = allKey
    Else
        GameVersionKey = LCase$(norm)
    End If
End Function

Public Sub EnsureComboSelection(ByRef cbo As MSForms.ComboBox, ByVal desiredValue As String, _
                                Optional ByVal fallback As String = "All")
    Dim target As String
    target = CleanSelection(desiredValue, fallback)

    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(CStr(cbo.list(i)), target, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i

    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    Else
        cbo.value = target
    End If
End Sub

Public Sub HighlightComboText(ByRef cbo As MSForms.ComboBox)
    If cbo Is Nothing Then Exit Sub
    On Error Resume Next
    cbo.SelStart = 0
    cbo.SelLength = Len(cbo.text)
    cbo.DropDown
    On Error GoTo 0
End Sub

Public Function FormatTypeName(ByVal rawValue As Variant) As String
    Dim t As String
    t = CleanSelection(rawValue, vbNullString)
    If Len(t) = 0 Then Exit Function
    t = LCase$(t)
    FormatTypeName = StrConv(t, vbProperCase)
End Function

Public Function DictionaryToSortedArray(ByVal dict As Object) As Variant
    If dict Is Nothing Then Exit Function
    If dict.count = 0 Then Exit Function

    Dim arr() As String
    ReDim arr(1 To dict.count)

    Dim idx As Long
    Dim key As Variant
    For Each key In dict.keys
        idx = idx + 1
        arr(idx) = CStr(key)
    Next key

    SortStringArray arr
    DictionaryToSortedArray = arr
End Function

Private Sub SortStringArray(ByRef arr() As String)
    On Error GoTo CleanExit
    If Not IsArray(arr) Then Exit Sub
    QuickSortStrings arr, LBound(arr), UBound(arr)
CleanExit:
End Sub

Private Sub QuickSortStrings(ByRef arr() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim i As Long, j As Long
    i = lo
    j = hi
    Dim pivot As String
    pivot = arr((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            Dim tmp As String
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSortStrings arr, lo, j
    If i < hi Then QuickSortStrings arr, i, hi
End Sub

Public Function CollectMoveTypeOptions(Optional ByRef movesTable As Variant) As Variant
    Dim tbl As Variant
    If IsMissing(movesTable) Then
        GlobalTables.LoadMovesTable
        tbl = GlobalTables.movesTable
    Else
        tbl = movesTable
    End If

    If IsEmpty(tbl) Then Exit Function

    Dim typeCol As Long
    typeCol = GlobalTables.FindHeaderColumn(tbl, "TYPE")
    If typeCol = 0 Then Exit Function

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim headerRow As Long
    headerRow = LBound(tbl, 1)
    Dim firstRow As Long
    firstRow = headerRow + 1
    Dim lastRow As Long
    lastRow = UBound(tbl, 1)

    Dim r As Long
    For r = firstRow To lastRow
        Dim typeText As String
        typeText = FormatTypeName(tbl(r, typeCol))
        If Len(typeText) > 0 Then
            If Not dict.Exists(typeText) Then dict.Add typeText, True
        End If
    Next r

    CollectMoveTypeOptions = DictionaryToSortedArray(dict)
End Function
