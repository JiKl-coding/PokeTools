Attribute VB_Name = "Functions"
' === Functions.bas ===
Option Explicit

' Absolute path to external pokedata workbook (initialized lazily)
Public POKEDATA_PATH As String
Public OWNS_POKEDATA As Boolean

' Cached reference to pokedata workbook (kept for this Excel session)
Private wbPokedata As Workbook

' Password
Public Const PASS As String = "pokemon"

Public Sub ProtectForMacros(ByVal ws As Worksheet)
    If Settings.Range("PROTECT_SHEETS").value = False Then
        Exit Sub
    End If
    On Error Resume Next
    ws.unprotect PASSWORD:=PASS
    ws.Protect PASSWORD:=PASS, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Public Sub UnprotectForMacros(ByVal ws As Worksheet)
    On Error Resume Next
    ws.unprotect PASSWORD:=PASS
    On Error GoTo 0
End Sub

Public Sub ProtectAllSheets()
    ProtectForMacros Pokedex
    ProtectForMacros Lists
    ProtectForMacros TypeChart
    ProtectForMacros Settings
End Sub

Public Sub UnprotectAllSheets()
    UnprotectForMacros Pokedex
    UnprotectForMacros Lists
    UnprotectForMacros TypeChart
    UnprotectForMacros Settings
End Sub

' Initialize relative paths
Private Sub InitPaths()
    Dim rawPath As String
    rawPath = ThisWorkbook.path & "\..\data\export\pokedata.xlsx"
    POKEDATA_PATH = NormalizePath(rawPath)
End Sub

' Normalize a path (resolves .. and returns absolute path)
Private Function NormalizePath(ByVal p As String) As String
    NormalizePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(p)
End Function

' Safely get workbook FullName (can fail in some cases)
Private Function SafeFullName(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeFullName = wb.FullName
    On Error GoTo 0
End Function

' Ensure paths are initialized before use
Public Sub CheckPaths()
    If Len(POKEDATA_PATH) = 0 Then
        InitPaths
    End If
End Sub


' Try to find already opened pokedata workbook.
' Returns Nothing if not opened.
Private Function FindOpenPokedataWb() As Workbook
    Dim wb As Workbook
    Dim fn As String

    ' 1) Match by absolute full path (best case)
    For Each wb In Application.Workbooks
        fn = LCase$(SafeFullName(wb))
        If Len(fn) > 0 Then
            If fn = LCase$(POKEDATA_PATH) Then
                Set FindOpenPokedataWb = wb
                Exit Function
            End If
        End If
    Next wb

    ' 2) Fallback: match by file name (handles OneDrive/URL FullName cases)
    For Each wb In Application.Workbooks
        If LCase$(wb.name) = "pokedata.xlsx" Then
            Set FindOpenPokedataWb = wb
            Exit Function
        End If
    Next wb
End Function

' Hide workbook window (keeps workbook open, prevents focus steal / UX jump)
Private Sub HideWorkbookWindow(ByVal wb As Workbook)
    On Error Resume Next
    If wb.Windows.Count > 0 Then
        wb.Windows(1).Visible = False
    End If
    On Error GoTo 0
End Sub

' Verify cached workbook reference is still valid.
Private Function IsCachedWbValid() As Boolean
    If wbPokedata Is Nothing Then
        IsCachedWbValid = False
        Exit Function
    End If

    On Error Resume Next
    Dim t As String
    t = wbPokedata.name
    IsCachedWbValid = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' Get pokedata workbook:
' - if already open: reuse it
' - if not open: open it silently in background (hidden window)
' - if user closed it during session: auto-reopen on next call
Public Function GetPokedataWb() As Workbook
    Dim wb As Workbook

    CheckPaths

    ' 1) Return cached workbook if still valid
    If IsCachedWbValid() Then
        Set GetPokedataWb = wbPokedata
        Exit Function
    Else
        Set wbPokedata = Nothing
    End If

    ' 2) Try to find already opened workbook
    Set wb = FindOpenPokedataWb()
    If Not wb Is Nothing Then
        Set wbPokedata = wb
        OWNS_POKEDATA = False
        Set GetPokedataWb = wbPokedata
        Exit Function
    End If

    ' 3) Open workbook as read-only in background (hidden)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    OWNS_POKEDATA = True

    Set wbPokedata = Workbooks.Open(POKEDATA_PATH, ReadOnly:=True)
    HideWorkbookWindow wbPokedata

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Bring focus back to tool window (extra safety)
    On Error Resume Next
    If ThisWorkbook.Windows.Count > 0 Then
        ThisWorkbook.Windows(1).Activate
    Else
        ThisWorkbook.Activate
    End If
    On Error GoTo 0

    Set GetPokedataWb = wbPokedata
End Function

' Convenience: ensure pokedata is available (call on Workbook_Open)
Public Sub EnsurePokedataOpen()
    Dim wb As Workbook
    Set wb = GetPokedataWb()
End Sub

