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
        If StrComp(CStr(cbo.List(i)), target, vbTextCompare) = 0 Then
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
    cbo.SelLength = Len(cbo.Text)
    cbo.DropDown
    On Error GoTo 0
End Sub
